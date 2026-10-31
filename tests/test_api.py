"""API layer: authentication, local-server protection, full backup/restore through the Api."""
import os, io, json, base64, zipfile, threading, unittest, http.client, hashlib
from _env import app, storage, TMP, ROOT, SEED, fresh, asset
import openpyxl


class Auth(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()

    def test_pin_flow_and_pbkdf2(self):
        r = json.loads(self.api.set_pin("1234")); self.assertTrue(r["ok"]); self.assertEqual(len(r["recovery_key"]), 19)
        d = json.load(open(os.path.join(TMP, "auth.json")))
        self.assertEqual(d["kdf"], "pbkdf2"); self.assertNotIn("rkey_plain", d)
        self.api.lock_session()
        self.assertFalse(json.loads(self.api.verify_pin("0000"))["valid"])
        self.assertTrue(json.loads(self.api.verify_pin("1234"))["valid"])

    def test_lockout_after_repeated_failures(self):
        self.api.set_pin("1234"); self.api.lock_session()
        for _ in range(5): self.api.verify_pin("bad")
        r = json.loads(self.api.verify_pin("1234")); self.assertFalse(r["valid"]); self.assertIn("Too many", r["error"])

    def test_cannot_change_pin_without_login_or_recovery(self):
        self.api.set_pin("1234"); self.api.lock_session()
        self.assertFalse(json.loads(self.api.set_pin("9999"))["ok"])
        self.assertTrue(json.loads(self.api.verify_pin("1234"))["valid"])

    def test_recovery_key_allows_reset(self):
        key = json.loads(self.api.set_pin("1234"))["recovery_key"]; self.api.lock_session()
        self.assertFalse(json.loads(self.api.verify_recovery_key("AAAA-AAAA-AAAA-AAAA"))["valid"])
        self.assertTrue(json.loads(self.api.verify_recovery_key(key.lower()))["valid"])
        self.assertTrue(json.loads(self.api.set_pin("5678"))["ok"]); self.api.lock_session()
        self.assertTrue(json.loads(self.api.verify_pin("5678"))["valid"])

    def test_legacy_sha256_pin_still_works_and_is_upgraded(self):
        salt = "ab" * 16
        json.dump({"hash": hashlib.sha256((salt + "4321").encode()).hexdigest(), "salt": salt}, open(os.path.join(TMP, "auth.json"), "w"))
        self.assertTrue(json.loads(self.api.verify_pin("4321"))["valid"])
        self.assertEqual(json.load(open(os.path.join(TMP, "auth.json")))["kdf"], "pbkdf2")
        self.api.lock_session(); self.assertTrue(json.loads(self.api.verify_pin("4321"))["valid"])


class ApiBehaviour(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()
    def J(self, s): return json.loads(s)

    def test_add_returns_id_and_errors_are_friendly(self):
        r = self.J(self.api.add_asset(json.dumps(asset(**{"Serial No.": "API-1"})))); self.assertTrue(r["ok"]); self.assertGreater(r["row"], 6)
        d = self.J(self.api.add_asset(json.dumps(asset(**{"Serial No.": "API-1"}))))
        self.assertFalse(d["ok"]); self.assertIn("already registered", d["error"])
        self.assertFalse(self.J(self.api.add_asset(json.dumps(asset(Remarks="x" * 40000))))["ok"])
        self.assertFalse(self.J(self.api.delete_asset(999999, "x"))["ok"])

    def test_ink_actions_through_api(self):
        self.assertTrue(self.J(self.api.ink_action("BK", "move_to_ict", 2, 2, ""))["ok"])
        r = self.J(self.api.ink_action("BK", "use", 50, 0, "")); self.assertFalse(r["ok"]); self.assertIn("available", r["error"])
        r = self.J(self.api.ink_action("BK", "move_to_ict", 50, 0, "")); self.assertIn("in store", r["error"])

    def test_edit_department_no_longer_overwrites_assigned_to(self):
        rid = self.J(self.api.add_asset(json.dumps(asset(**{"Serial No.": "BUG-1", "Assigned To": "Jane Doe", "Department/Location": "Lab"}))))["row"]
        self.api.update_asset(rid, json.dumps({"Department/Location": "Pharmacy"}))
        self.api.write_history("Laptop", "BUG-1", "Pharmacy", "reassigned", "Department changed: Lab → Pharmacy", "2026-01-01", "", "ICT Manager")
        a = [r for r in self.J(self.api.load_data())["data"]["inventory"] if r["_row"] == rid][0]
        self.assertEqual((a["Assigned To"], a["Department/Location"]), ("Jane Doe", "Pharmacy"))

    def test_pdf_report_generates(self):
        r = self.J(self.api.generate_pdf()); self.assertTrue(r["ok"], r)
        self.assertTrue(base64.b64decode(r["b64"])[:5] == b"%PDF-")

    def test_load_data_cache_hits_and_invalidates_across_writers(self):
        import time
        a = self.api.load_data(); t = time.time(); b = self.api.load_data()
        self.assertEqual(a, b); self.assertLess(time.time() - t, 0.05)
        other = storage.Storage(TMP)                                   # a SECOND engine = another window/process
        other.add_asset(asset(**{"Serial No.": "OTHER-WRITER"}))
        self.assertIn("OTHER-WRITER", self.api.load_data()); json.loads(self.api.load_data())

    def test_open_excel_exports_a_copy(self):
        os.environ["PATH"] += ""                                         # headless: just verify the export file is produced
        p = self.eng.export_for_user(); self.assertTrue(os.path.exists(p))


class HttpGuards(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.eng, cls.api = fresh()
        cls.srv = app.LocalServer(("127.0.0.1", 0), app.make_handler(cls.api, os.path.join(ROOT, "src", "ui.html")))
        cls.port = cls.srv.server_port
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
    @classmethod
    def tearDownClass(cls): cls.srv.shutdown(); cls.srv.server_close()
    def setUp(self): self.api.lock_session()

    def call(self, method, args=(), token="test-token", host=None, origin=None):
        c = http.client.HTTPConnection("127.0.0.1", self.port)
        h = {"Content-Type": "application/json", "Host": host or f"127.0.0.1:{self.port}"}
        if token is not None: h["X-ICT-Token"] = token
        if origin: h["Origin"] = origin
        c.request("POST", f"/api/{method}", json.dumps({"args": list(args)}), h)
        r = c.getresponse(); return r.status, r.read().decode()

    def get(self, path, host=None):
        c = http.client.HTTPConnection("127.0.0.1", self.port); c.request("GET", path, headers={"Host": host or f"127.0.0.1:{self.port}"})
        r = c.getresponse(); return r.status, r.read(), r

    def test_no_token_rejected(self): self.assertEqual(self.call("load_data", token=None)[0], 403)
    def test_wrong_token_rejected(self): self.assertEqual(self.call("load_data", token="nope")[0], 403)
    def test_foreign_origin_rejected(self): self.assertEqual(self.call("is_pin_set", origin="http://evil.example")[0], 403)
    def test_dns_rebinding_host_rejected(self): self.assertEqual(self.call("is_pin_set", host="evil.example:80")[0], 403)
    def test_public_method_ok_before_login(self): self.assertEqual(self.call("is_pin_set")[0], 200)

    def test_private_methods_blocked_until_login(self):
        self.api.set_pin("1234"); self.api.lock_session()
        for m in ("load_data", "get_recovery_hint", "export_full_backup", "delete_asset", "open_excel", "ink_action"):
            self.assertEqual(self.call(m, [1, "x"] if m == "delete_asset" else [])[0], 401, m)
        self.assertEqual(self.call("verify_pin", ["1234"])[0], 200); self.assertEqual(self.call("load_data")[0], 200)

    def test_internal_attributes_not_callable(self):
        for m in ("_throttle", "needs_login", "engine", "PUBLIC_METHODS"): self.assertEqual(self.call(m)[0], 404, m)

    def test_security_headers_fonts_and_logo(self):
        st, body, r = self.get("/")
        self.assertEqual(st, 200); self.assertIn("default-src 'self'", r.getheader("Content-Security-Policy")); self.assertEqual(r.getheader("X-Frame-Options"), "DENY")
        self.assertNotIn(b"googleapis", body)                           # fully offline-capable UI
        st, body, r = self.get("/fonts/inter-latin-400-normal.woff2"); self.assertEqual((st, r.getheader("Content-Type")), (200, "font/woff2")); self.assertTrue(body[:4] == b"wOF2")
        self.assertEqual(self.get("/fonts/../main.py")[0], 404); self.assertEqual(self.get("/fonts/x.txt")[0], 404)

    def test_downloads_require_login_when_pin_set(self):
        self.api.set_pin("1234"); self.api.lock_session()
        self.assertEqual(self.get("/reports/x.pdf")[0], 401); self.assertEqual(self.get("/backups/x.db")[0], 401)


class FullBackupRestore(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()
    def J(self, s): return json.loads(s)

    def test_export_then_restore_roundtrip(self):
        self.eng.add_asset(asset(**{"Serial No.": "BEFORE"}))
        bk = self.J(self.api.export_full_backup()); self.assertTrue(bk["ok"])
        z = zipfile.ZipFile(io.BytesIO(base64.b64decode(bk["b64"])))
        self.assertTrue({"inventory.db", "ICT_MASTER.xlsx", "ICT_MANIFEST.json"} <= set(z.namelist()))
        self.eng.add_asset(asset(**{"Serial No.": "AFTER"}))
        self.assertTrue(self.J(self.api.restore_full_backup(bk["b64"]))["ok"])
        sers = {r["Serial No."] for r in self.eng.load_all()["inventory"]}
        self.assertIn("BEFORE", sers); self.assertNotIn("AFTER", sers)

    def _zip(self, files, app_name="DericBI Inventory Control"):
        b = io.BytesIO()
        with zipfile.ZipFile(b, "w") as z:
            z.writestr("ICT_MANIFEST.json", json.dumps({"app": app_name}))
            for n, d in files.items(): z.writestr(n, d)
        return base64.b64encode(b.getvalue()).decode()

    def test_corrupt_backups_are_rejected_and_nothing_changes(self):
        self.eng.add_asset(asset(**{"Serial No.": "KEEP-ME"}))
        for files in ({"inventory.db": b"garbage" * 100}, {"ICT_MASTER.xlsx": b"not excel"}, {}):
            r = self.J(self.api.restore_full_backup(self._zip(files))); self.assertFalse(r["ok"], files)
            self.assertTrue(any(x["Serial No."] == "KEEP-ME" for x in self.eng.load_all()["inventory"]))
        self.assertFalse(self.J(self.api.restore_full_backup(self._zip({"inventory.db": b""}, app_name="Something Else")))["ok"])
        self.assertFalse(self.J(self.api.restore_full_backup("bm90IGEgemlw"))["ok"])

    def test_legacy_excel_backup_from_previous_version_restores(self):
        wb = openpyxl.load_workbook(SEED); wb["Inventory"].append(["", "", "Laptop", "Legacy", "LEG-1", "", "ICT", "Good"] + [""] * 8)
        buf = io.BytesIO(); wb.save(buf)
        r = self.J(self.api.restore_full_backup(self._zip({"ICT_MASTER.xlsx": buf.getvalue()}, app_name="ICT Command Centre")))
        self.assertTrue(r["ok"], r); self.assertTrue(any(x["Serial No."] == "LEG-1" for x in self.eng.load_all()["inventory"]))


if __name__ == "__main__":
    unittest.main()
