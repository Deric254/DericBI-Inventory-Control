"""Backend safety net: data integrity, auth, and local-API protection.
Run:  python -m unittest discover -s tests -v
"""
import os, sys, json, shutil, tempfile, threading, unittest, http.client, base64, io, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMP  = tempfile.mkdtemp(prefix="dericbi_test_")
os.environ["ICT_DATA_DIR"]   = TMP
os.environ["ICT_BUNDLE_DIR"] = os.path.join(ROOT, "assets")
os.environ["ICT_API_TOKEN"]  = "test-token"
sys.path.insert(0, os.path.join(ROOT, "src"))
os.environ.setdefault("PYWEBVIEW_NO_UIAUTOMATION", "1")
import main as app   # noqa: E402
import openpyxl      # noqa: E402

SEED = os.path.join(ROOT, "assets", "ICT_MASTER.xlsx")

def fresh():
    for f in ("ICT_MASTER.xlsx", "auth.json", "config.json"):
        p = os.path.join(TMP, f)
        if os.path.exists(p): os.remove(p)
    shutil.rmtree(os.path.join(TMP, "backups"), ignore_errors=True)
    shutil.copy2(SEED, os.path.join(TMP, "ICT_MASTER.xlsx"))
    eng = app.ExcelEngine(); return eng, app.Api(eng, app.ConfigEngine())

def asset(**kw):
    d = {"Equipment Type": "Laptop", "Brand/Model": "Test", "Serial No.": "T-1",
         "Department/Location": "ICT", "Condition/Status": "Good"}
    d.update(kw); return d

class Integrity(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()

    def test_add_update_delete_roundtrip(self):
        n0 = len(self.eng.load_all()["inventory"])
        row = self.eng.add_asset(asset())
        inv = self.eng.load_all()["inventory"]
        self.assertEqual(len(inv), n0 + 1)
        self.eng.update_asset(row, {"Brand/Model": "Changed"})
        self.assertEqual(self.eng.load_all()["inventory"][-1]["Brand/Model"], "Changed")
        self.eng.delete_asset(row, "x")
        self.assertEqual(len(self.eng.load_all()["inventory"]), n0)

    def test_header_row_and_out_of_range_rows_are_protected(self):
        for bad in (0, 1, 99999, -3):
            with self.assertRaises(ValueError): self.eng.update_asset(bad, {"Brand/Model": "x"})
            with self.assertRaises(ValueError): self.eng.delete_asset(bad, "x")
        wb = openpyxl.load_workbook(os.path.join(TMP, "ICT_MASTER.xlsx"))
        self.assertEqual(wb["Inventory"].cell(1, 1).value, app.INV_COLS[0])

    def test_formula_injection_stays_text(self):
        row = self.eng.add_asset(asset(Remarks="=HYPERLINK(\"http://evil\",\"x\")", **{"Assigned To": "=1+1"}))
        wb = openpyxl.load_workbook(os.path.join(TMP, "ICT_MASTER.xlsx"))
        ws = wb["Inventory"]
        for cell in ws[row]:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                self.assertEqual(cell.data_type, "s")
        self.eng.update_asset(row, {"Remarks": "=SUM(A1:A9)"})
        wb = openpyxl.load_workbook(os.path.join(TMP, "ICT_MASTER.xlsx"))
        c = [c for c in wb["Inventory"][row] if c.value == "=SUM(A1:A9)"]
        self.assertTrue(c and c[0].data_type == "s")

    def test_oversized_field_rejected_not_truncated(self):
        r = json.loads(self.api.add_asset(json.dumps(asset(Remarks="x" * 40000))))
        self.assertFalse(r["ok"])

    def test_failed_save_leaves_database_intact(self):
        before = open(os.path.join(TMP, "ICT_MASTER.xlsx"), "rb").read()
        real = openpyxl.Workbook.save
        def boom(self_, path): open(path, "wb").write(b"half-written"); raise OSError("disk full")
        openpyxl.Workbook.save = boom
        try:
            with self.assertRaises(OSError): self.eng.add_asset(asset())
        finally: openpyxl.Workbook.save = real
        self.assertEqual(open(os.path.join(TMP, "ICT_MASTER.xlsx"), "rb").read(), before)
        self.eng.load_all()   # still opens

    def test_no_temp_file_left_and_daily_backup_made(self):
        self.eng.add_asset(asset())
        self.assertFalse(os.path.exists(os.path.join(TMP, "ICT_MASTER.xlsx.tmp")))
        self.assertTrue(os.listdir(os.path.join(TMP, "backups", "daily")))

    def test_concurrent_writes_do_not_lose_rows(self):
        n0 = len(self.eng.load_all()["inventory"])
        ts = [threading.Thread(target=lambda i=i: self.eng.add_asset(asset(**{"Serial No.": f"C-{i}"}))) for i in range(12)]
        [t.start() for t in ts]; [t.join() for t in ts]
        inv = self.eng.load_all()["inventory"]
        self.assertEqual(len(inv), n0 + 12)
        self.assertEqual(len({r["Serial No."] for r in inv if str(r["Serial No."]).startswith("C-")}), 12)

    def test_audit_log_written(self):
        self.eng.add_asset(asset())
        self.assertTrue(any(l["action"] == "ADD" for l in self.eng.load_all()["log"]))

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
        r = json.loads(self.api.verify_pin("1234"))
        self.assertFalse(r["valid"]); self.assertIn("Too many", r["error"])

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
        import hashlib
        salt = "ab" * 16
        json.dump({"hash": hashlib.sha256((salt + "4321").encode()).hexdigest(), "salt": salt}, open(os.path.join(TMP, "auth.json"), "w"))
        self.assertTrue(json.loads(self.api.verify_pin("4321"))["valid"])
        self.assertEqual(json.load(open(os.path.join(TMP, "auth.json")))["kdf"], "pbkdf2")
        self.api.lock_session()
        self.assertTrue(json.loads(self.api.verify_pin("4321"))["valid"])

class HttpGuards(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.eng, cls.api = fresh()
        cls.srv = app.LocalServer(("127.0.0.1", 0), app.make_handler(cls.api, os.path.join(ROOT, "src", "ui.html")))
        cls.port = cls.srv.server_port
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
    @classmethod
    def tearDownClass(cls): cls.srv.shutdown(); cls.srv.server_close()
    def setUp(self): self.api.lock_session(); self.eng, _ = self.eng, None

    def call(self, method, args=(), token="test-token", host=None, origin=None, path=None):
        c = http.client.HTTPConnection("127.0.0.1", self.port)
        h = {"Content-Type": "application/json", "Host": host or f"127.0.0.1:{self.port}"}
        if token is not None: h["X-ICT-Token"] = token
        if origin: h["Origin"] = origin
        c.request("POST", path or f"/api/{method}", json.dumps({"args": list(args)}), h)
        r = c.getresponse(); body = r.read().decode(); return r.status, body

    def test_no_token_rejected(self): self.assertEqual(self.call("load_data", token=None)[0], 403)
    def test_wrong_token_rejected(self): self.assertEqual(self.call("load_data", token="nope")[0], 403)
    def test_foreign_origin_rejected(self): self.assertEqual(self.call("is_pin_set", origin="http://evil.example")[0], 403)
    def test_dns_rebinding_host_rejected(self): self.assertEqual(self.call("is_pin_set", host="evil.example:80")[0], 403)
    def test_public_method_ok_before_login(self): self.assertEqual(self.call("is_pin_set")[0], 200)
    def test_private_methods_blocked_until_login(self):
        self.api.set_pin("1234"); self.api.lock_session()
        for m in ("load_data", "get_recovery_hint", "export_full_backup", "delete_asset"):
            self.assertEqual(self.call(m, [2, "x"] if m == "delete_asset" else [])[0], 401, m)
        self.assertEqual(self.call("verify_pin", ["1234"])[0], 200)
        self.assertEqual(self.call("load_data")[0], 200)
    def test_internal_attributes_not_callable(self):
        for m in ("_throttle", "needs_login", "engine", "PUBLIC_METHODS"):
            self.assertEqual(self.call(m)[0], 404, m)
    def test_security_headers_and_logo(self):
        c = http.client.HTTPConnection("127.0.0.1", self.port); c.request("GET", "/", headers={"Host": f"127.0.0.1:{self.port}"})
        r = c.getresponse(); r.read()
        self.assertEqual(r.status, 200); self.assertIn("default-src 'self'", r.getheader("Content-Security-Policy"))
        self.assertEqual(r.getheader("X-Frame-Options"), "DENY")
    def test_reports_require_login_when_pin_set(self):
        self.api.set_pin("1234"); self.api.lock_session()
        c = http.client.HTTPConnection("127.0.0.1", self.port); c.request("GET", "/reports/x.pdf", headers={"Host": f"127.0.0.1:{self.port}"})
        r = c.getresponse(); r.read(); self.assertEqual(r.status, 401)

class Restore(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()
    def _zip(self, master_bytes):
        b = io.BytesIO()
        with zipfile.ZipFile(b, "w") as z:
            z.writestr("ICT_MANIFEST.json", json.dumps({"app": "ICT Command Centre"}))   # legacy name accepted
            z.writestr("ICT_MASTER.xlsx", master_bytes)
        return base64.b64encode(b.getvalue()).decode()
    def test_corrupt_backup_rejected_and_live_data_untouched(self):
        self.eng.add_asset(asset(**{"Serial No.": "KEEP-ME"}))
        r = json.loads(self.api.restore_full_backup(self._zip(b"not an excel file")))
        self.assertFalse(r["ok"])
        self.assertTrue(any(x["Serial No."] == "KEEP-ME" for x in self.eng.load_all()["inventory"]))
    def test_valid_backup_restores_and_snapshots_current(self):
        good = open(SEED, "rb").read()
        self.eng.add_asset(asset(**{"Serial No.": "WILL-GO"}))
        r = json.loads(self.api.restore_full_backup(self._zip(good)))
        self.assertTrue(r["ok"], r)
        self.assertFalse(any(x["Serial No."] == "WILL-GO" for x in self.eng.load_all()["inventory"]))
        self.assertTrue(os.listdir(os.path.join(TMP, "backups", "pre-restore")))

class ReadOnlyLoadMatches(unittest.TestCase):
    def test_load_all_shape(self):
        eng, _ = fresh(); d = eng.load_all()
        self.assertEqual(set(d), {"inventory", "ink", "log", "history"})
        self.assertTrue(all("_row" in r for r in d["inventory"]))
        self.assertEqual(d["inventory"][0]["_row"], 2)


class ReadCache(unittest.TestCase):
    def test_cache_hits_and_invalidates(self):
        import time
        eng, api = fresh()
        a = api.load_data(); t = time.time(); b = api.load_data(); fast = time.time() - t
        self.assertEqual(a, b); self.assertLess(fast, 0.05)
        eng.add_asset(asset(**{"Serial No.": "CACHE-1"}))
        self.assertIn("CACHE-1", api.load_data())                      # own write invalidates
        wb = openpyxl.load_workbook(os.path.join(TMP, "ICT_MASTER.xlsx"))
        wb["Inventory"].append([""] * 4 + ["EXTERNAL-EDIT"]); wb.save(os.path.join(TMP, "ICT_MASTER.xlsx"))
        self.assertIn("EXTERNAL-EDIT", api.load_data())                # edit made in Excel is picked up
        json.loads(api.load_data())                                    # always valid JSON

if __name__ == "__main__":
    unittest.main()
