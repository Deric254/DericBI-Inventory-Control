"""Licence gate: a locked copy is refused everywhere that matters, the licence methods stay open, a licence opens,
renews and expires the copy, and the licence file never gets mixed up with business data or backups.
Throwaway keys only -- never the real owner key."""
import base64, http.client, io, json, os, re, threading, unittest, zipfile
from datetime import date, timedelta
from _env import (app, ROOT, TMP, fresh, make_licensing, use_owner_copy, licence_gate, Licensing, MemoryStore)
from licensing.licensing import FileStore

J = json.loads
UI = os.path.join(ROOT, "src", "ui.html")


class Http(unittest.TestCase):
    """A real server on a free port, so the gate is tested exactly where it runs."""
    @classmethod
    def setUpClass(cls):
        cls.eng, cls.api = fresh()
        cls.srv = app.LocalServer(("127.0.0.1", 0), app.make_handler(cls.api, UI))
        cls.port = cls.srv.server_port
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls): cls.srv.shutdown(); cls.srv.server_close()

    def tearDown(self): use_owner_copy(); self.api.lock_session()

    def call(self, method, args=()):
        c = http.client.HTTPConnection("127.0.0.1", self.port)
        c.request("POST", f"/api/{method}", json.dumps({"args": list(args)}),
                  {"Content-Type": "application/json", "Host": f"127.0.0.1:{self.port}", "X-ICT-Token": "test-token"})
        r = c.getresponse(); return r.status, J(r.read().decode())

    def get(self, path):
        c = http.client.HTTPConnection("127.0.0.1", self.port)
        c.request("GET", path, headers={"Host": f"127.0.0.1:{self.port}"})
        r = c.getresponse(); return r.status, r.read()

    def locked(self, today=None):
        """A copy with a valid build key but no licence: (licensing, owner seed, owner key text)."""
        lic, key_text, seed = make_licensing(today); licence_gate.use_licensing(lic); return lic, seed, key_text


class LockedCopy(Http):
    def test_api_calls_are_refused_with_402(self):
        self.locked()
        for m in ("load_data", "is_pin_set", "verify_pin", "set_pin", "export_full_backup", "remove_sample_data", "lock_session"):
            st, body = self.call(m, ["1234", "tester"] if m in ("verify_pin", "set_pin") else [])
            self.assertEqual(st, 402, m)
            self.assertEqual(body["licence"], {"ok": False, "kind": "unlicensed"}, m)
            self.assertFalse(body["ok"]); self.assertIn("licensed", body["error"])

    def test_creating_an_account_is_gated_too(self):
        self.locked(); self.assertFalse(os.path.exists(app.AUTH_FILE))
        self.assertEqual(self.call("set_pin", ["1234", "tester"])[0], 402)
        self.assertFalse(os.path.exists(app.AUTH_FILE), "nothing may be created while locked")

    def test_licence_methods_the_page_fonts_and_logo_stay_open(self):
        self.locked()
        st, body = self.call("licence_state")
        self.assertEqual(st, 200); self.assertFalse(body["licence"]["ok"])
        self.assertEqual(len(body["licence"]["install_id"].replace("-", "")), 12)
        self.assertEqual(self.get("/")[0], 200)
        self.assertEqual(self.get("/fonts/inter-latin-400-normal.woff2")[0], 200)

    def test_report_and_backup_downloads_are_refused(self):
        self.locked()
        self.assertEqual(self.get("/reports/x.pdf")[0], 402); self.assertEqual(self.get("/backups/x.db")[0], 402)
        use_owner_copy()
        self.assertEqual(self.get("/reports/x.pdf")[0], 404)       # past the gate: the file just doesn't exist

    def test_build_without_a_public_key_fails_closed(self):
        licence_gate.use_licensing(Licensing(licence_gate.APP_ID, licence_gate.PREFIX, "", MemoryStore()))
        st, body = self.call("load_data")
        self.assertEqual(st, 402); self.assertEqual(body["licence"]["kind"], "unconfigured")
        st, body = self.call("licence_state")
        self.assertEqual(st, 200); self.assertFalse(body["licence"]["ok"]); self.assertNotIn("install_id", body["licence"])
        st, body = self.call("activate_licence", ["x.y.z"]); self.assertEqual(st, 200); self.assertFalse(body["ok"])

    def test_unknown_methods_are_still_404_not_402(self):
        self.locked(); self.assertEqual(self.call("no_such_method")[0], 404); self.assertEqual(self.call("_throttle")[0], 404)


class Activation(Http):
    def activate(self, token): return self.call("activate_licence", [token])[1]

    def test_licence_for_this_computer_opens_the_app(self):
        lic, seed, _ = self.locked()
        code = self.call("licence_state")[1]["licence"]["install_id"]
        r = self.activate(lic.issue(seed, code, "Lucy Chemist", 30)["licence"])
        self.assertTrue(r["ok"]); self.assertEqual((r["licence"]["kind"], r["licence"]["user"]), ("licensed", "Lucy Chemist"))
        self.assertEqual(self.call("is_pin_set")[0], 200)           # past the gate now

    def test_licence_for_another_computer_is_rejected(self):
        lic, seed, _ = self.locked()
        r = self.activate(lic.issue(seed, "ABCD-EFGH-JKLM", "Lucy", 30)["licence"])
        self.assertFalse(r["ok"]); self.assertIn("different computer", r["error"])
        self.assertEqual(self.call("load_data")[0], 402)

    def test_garbage_and_oversize_are_rejected(self):
        self.locked()
        for bad in ("not a licence", "", "DIC1.a.b", 123, None, "x" * 5000):
            r = self.activate(bad); self.assertFalse(r["ok"], repr(bad)[:20]); self.assertTrue(r["error"])
        self.assertEqual(self.call("load_data")[0], 402)

    def test_a_licence_signed_by_someone_else_is_rejected(self):
        lic, _, _ = self.locked()
        other, _, oseed = make_licensing()                           # a different owner key
        token = other.issue(oseed, lic.install_id(), "Mallory", 30)["licence"]
        r = self.activate(token); self.assertFalse(r["ok"]); self.assertIn("wasn't signed by the owner", r["error"])

    def test_renewal_must_be_for_the_same_person(self):
        lic, seed, _ = self.locked(); code = lic.install_id()
        self.assertTrue(self.activate(lic.issue(seed, code, "Lucy", 30)["licence"])["ok"])
        r = self.activate(lic.issue(seed, code, "Someone Else", 30)["licence"]); self.assertFalse(r["ok"])
        self.assertTrue(self.activate(lic.issue(seed, code, "lucy", 60)["licence"])["ok"])

    def test_licence_stops_working_after_it_expires(self):
        today = {"d": date(2026, 1, 1)}
        lic, key_text, seed = make_licensing(); lic._today = lambda: today["d"]; licence_gate.use_licensing(lic)
        self.assertTrue(self.activate(lic.issue(seed, lic.install_id(), "Lucy", 1)["licence"])["ok"])
        self.assertEqual(self.call("is_pin_set")[0], 200)
        today["d"] += timedelta(days=3); licence_gate.reset_cache()
        st, body = self.call("is_pin_set"); self.assertEqual(st, 402); self.assertIn("expired", body["error"])
        today["d"] -= timedelta(days=3); licence_gate.reset_cache()      # clock back inside the term: still fine

    def test_renewal_notice_data_counts_days_left(self):
        lic, seed, _ = self.locked(); code = lic.install_id()
        r = self.activate(lic.issue(seed, code, "Lucy", 10)["licence"])
        self.assertEqual((r["licence"]["days_left"], r["licence"]["expires"]), (10, (date.today() + timedelta(days=10)).isoformat()))


class OwnerClaim(Http):
    def test_wrong_owner_key_is_rejected(self):
        self.locked(); other_text, _ = Licensing.make_owner_key(licence_gate.PREFIX)
        st, body = self.call("claim_owner", [other_text]); self.assertEqual(st, 200); self.assertFalse(body["ok"])
        self.assertEqual(self.call("load_data")[0], 402)

    def test_not_an_owner_key_and_oversize_are_rejected(self):
        self.locked()
        for bad in ("hello", "", 5, None, "x" * 600): self.assertFalse(self.call("claim_owner", [bad])[1]["ok"], repr(bad)[:20])

    def test_real_owner_key_opens_this_copy(self):
        _, _, key_text = self.locked()
        st, body = self.call("claim_owner", [key_text])
        self.assertTrue(body["ok"]); self.assertEqual(body["licence"]["kind"], "owner")
        self.assertEqual(self.call("is_pin_set")[0], 200)

    def test_owner_key_is_never_stored(self):
        lic, key_text, _ = make_licensing(); store = lic.store; licence_gate.use_licensing(lic)
        self.call("claim_owner", [key_text])
        self.assertNotIn(key_text.split("-OWNER-")[1], json.dumps(store.data))


class LoginStillApplies(Http):
    def test_a_licensed_copy_still_needs_the_pin(self):
        lic, seed, _ = self.locked()
        self.assertTrue(J(self.api.activate_licence(lic.issue(seed, lic.install_id(), "Lucy", 30)["licence"]))["ok"])
        self.assertEqual(self.call("set_pin", ["1234", "tester"])[0], 200)
        self.api.lock_session()
        self.assertEqual(self.call("load_data")[0], 401)             # licensed but signed out: 401, not 402
        self.assertEqual(self.call("verify_pin", ["1234", "tester"])[0], 200)
        self.assertEqual(self.call("load_data")[0], 200)

    def test_licence_screen_works_before_login_even_with_a_pin_set(self):
        use_owner_copy(); self.call("set_pin", ["1234", "tester"]); self.api.lock_session()
        lic, seed, _ = self.locked()                                 # now: PIN exists, nobody signed in, no licence
        self.assertEqual(self.call("licence_state")[0], 200)         # not 401
        self.assertTrue(self.call("activate_licence", [lic.issue(seed, lic.install_id(), "Lucy", 30)["licence"]])[1]["ok"])
        self.assertTrue(os.path.exists(app.AUTH_FILE))               # the account was left exactly as it was
        self.assertEqual(self.call("load_data")[0], 401)             # and signing in is still required


class StateHandling(Http):
    def test_state_is_cached_between_calls_and_reset_on_activation(self):
        lic, seed, _ = self.locked(); calls = []; original = lic.state
        lic.state = lambda: (calls.append(1), original())[1]
        licence_gate.current_state(); licence_gate.current_state(); self.assertEqual(len(calls), 1)
        licence_gate.reset_cache(); licence_gate.current_state(); self.assertEqual(len(calls), 2)
        self.assertFalse(licence_gate.current_state()["ok"])         # cached as locked...
        self.api.activate_licence(lic.issue(seed, lic.install_id(), "Lucy", 30)["licence"])
        self.assertTrue(licence_gate.current_state()["ok"])          # ...and unlocked the moment a licence is activated

    def test_static_files_do_not_trigger_a_licence_check(self):
        lic, _, _ = self.locked(); calls = []; original = lic.state
        lic.state = lambda: (calls.append(1), original())[1]
        self.get("/"); self.get("/fonts/inter-latin-400-normal.woff2"); self.get("/logo")
        self.assertEqual(calls, [])


class SeparateFromBusinessData(Http):
    def real_store_copy(self):
        lic, _, seed = make_licensing()
        public = lic.public.hex()
        real = Licensing(licence_gate.APP_ID, licence_gate.PREFIX, public, FileStore(TMP, licence_gate.STATE_FILENAME))
        licence_gate.use_licensing(real); return real, seed

    def test_licence_file_lives_in_the_data_folder_and_survives_a_restart(self):
        real, seed = self.real_store_copy()
        self.assertTrue(self.call("activate_licence", [real.issue(seed, real.install_id(), "Lucy", 30)["licence"]])[1]["ok"])
        path = os.path.join(TMP, "licence-state.json"); self.assertTrue(os.path.isfile(path))
        again = Licensing(licence_gate.APP_ID, licence_gate.PREFIX, real.public.hex(), FileStore(TMP, licence_gate.STATE_FILENAME))
        self.assertEqual(again.install_id(), real.install_id())
        self.assertEqual(again.state()["kind"], "licensed")

    def test_full_backup_never_contains_the_licence_and_restore_never_removes_it(self):
        real, seed = self.real_store_copy(); self.call("activate_licence", [real.issue(seed, real.install_id(), "Lucy", 30)["licence"]])
        bk = J(self.api.export_full_backup()); self.assertTrue(bk["ok"])
        names = zipfile.ZipFile(io.BytesIO(base64.b64decode(bk["b64"]))).namelist()
        self.assertFalse([n for n in names if "licence" in n.lower() or "licensing" in n.lower()], names)
        self.assertTrue(J(self.api.restore_full_backup(bk["b64"]))["ok"])
        self.assertTrue(os.path.isfile(os.path.join(TMP, "licence-state.json")))
        self.assertEqual(real.state()["kind"], "licensed")

    def test_default_build_state_is_never_ok_without_a_key_or_licence(self):
        licence_gate.use_licensing(None)                              # the real configuration, as shipped
        st = licence_gate.get_licensing().state()
        self.assertFalse(st["ok"]); self.assertIn(st["kind"], ("unconfigured", "unlicensed"))


class ShippedConfigurationIsConsistent(unittest.TestCase):
    def test_tool_config_matches_the_app(self):
        with open(os.path.join(ROOT, "src", "licensing", "licensing.json"), encoding="utf-8") as f: cfg = json.load(f)
        self.assertEqual((cfg["app_id"], cfg["prefix"]), (licence_gate.APP_ID, licence_gate.PREFIX))
        self.assertTrue(os.path.isfile(os.path.join(ROOT, cfg["public_key_file"])))
        self.assertEqual(cfg["public_key_var"], "OWNER_PUBLIC_KEY")

    def test_public_key_is_empty_or_a_valid_32_byte_hex(self):
        from licence_owner_key import OWNER_PUBLIC_KEY
        self.assertTrue(OWNER_PUBLIC_KEY == "" or (len(OWNER_PUBLIC_KEY) == 64 and bytes.fromhex(OWNER_PUBLIC_KEY)))

    def test_no_private_owner_key_is_in_the_project(self):
        pat = re.compile(r"\b[A-Z]{2,6}-OWNER-[A-Za-z0-9_\-]{43}\b")
        for dp, dn, fn in os.walk(ROOT):
            dn[:] = [d for d in dn if d not in ("node_modules", ".git", "python", "dist", "__pycache__")]
            for f in fn:
                self.assertFalse(f.endswith("-owner.key"), f)
                if f.endswith((".py", ".js", ".json", ".yml", ".md", ".html", ".bat", ".txt")):
                    with open(os.path.join(dp, f), encoding="utf-8", errors="ignore") as fh: t = fh.read()
                    self.assertIsNone(pat.search(t), f"private key text in {f}")


if __name__ == "__main__":
    unittest.main()
