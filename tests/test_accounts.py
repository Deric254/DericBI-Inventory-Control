"""Accounts (username + PIN) and removal of the shipped sample data."""
import os, json, hashlib, unittest
from _env import app, storage, TMP, SEED, fresh, asset

J = json.loads
AUTH = lambda: os.path.join(TMP, "auth.json")


class Accounts(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()

    def test_first_launch_requires_a_valid_username(self):
        for bad in (None, "", "   ", "ab", "x" * 33, "bad\nname", "a/b", "<script>", "...", "___"):
            self.assertFalse(J(self.api.set_pin("1234", bad))["ok"], repr(bad))
            self.assertFalse(os.path.exists(AUTH()), "nothing may be written for an invalid username")
        r = J(self.api.set_pin("1234", "  Dr. A_b-1  "))
        self.assertTrue(r["ok"]); self.assertEqual(r["username"], "Dr. A_b-1")
        self.assertEqual(json.load(open(AUTH()))["username"], "Dr. A_b-1")

    def test_signing_in_needs_username_and_pin(self):
        self.api.set_pin("1234", "Alice"); self.api.lock_session()
        self.assertTrue(J(self.api.is_pin_set())["has_username"])
        self.assertFalse(J(self.api.verify_pin("1234"))["valid"])                 # PIN alone is not enough
        self.assertFalse(J(self.api.verify_pin("1234", "bob"))["valid"])          # wrong user
        self.assertFalse(J(self.api.verify_pin("0000", "alice"))["valid"])        # wrong PIN
        self.assertTrue(self.api.needs_login())
        r = J(self.api.verify_pin("1234", "  aLiCe "))                            # case/space-insensitive
        self.assertTrue(r["valid"]); self.assertEqual(r["username"], "Alice"); self.assertFalse(self.api.needs_login())

    def test_wrong_username_counts_toward_lockout(self):
        self.api.set_pin("1234", "alice"); self.api.lock_session()
        for _ in range(5): self.api.verify_pin("1234", "mallory")                 # right PIN, wrong user
        r = J(self.api.verify_pin("1234", "alice")); self.assertFalse(r["valid"]); self.assertIn("Too many", r["error"])

    def test_username_is_not_a_hint_before_login(self):
        self.api.set_pin("1234", "alice"); self.api.lock_session()
        for resp in (self.api.verify_pin("1234", "nobody"), self.api.verify_pin("0000", "alice")):
            self.assertEqual(set(J(resp)), {"ok", "valid"})                       # same shape, no username leaked

    def test_changing_and_resetting_pin_keep_the_account_name(self):
        self.api.set_pin("1234", "alice")
        key = J(self.api.set_pin("5678"))["recovery_key"]                         # change while signed in
        self.api.lock_session(); self.assertTrue(J(self.api.verify_pin("5678", "alice"))["valid"])
        self.api.lock_session()
        r = J(self.api.verify_recovery_key(key)); self.assertTrue(r["valid"]); self.assertEqual(r["username"], "alice")
        self.assertEqual(J(self.api.set_pin("9999"))["username"], "alice")        # reset via recovery key
        self.api.lock_session(); self.assertTrue(J(self.api.verify_pin("9999", "alice"))["valid"])

    def test_username_cannot_be_changed_or_added_without_login(self):
        self.api.set_pin("1234", "alice"); self.api.lock_session()
        self.assertFalse(J(self.api.set_pin("9999", "mallory"))["ok"])
        self.assertTrue(J(self.api.verify_pin("1234", "alice"))["valid"])

    def test_session_user_is_forgotten_on_lock(self):
        self.api.set_pin("1234", "alice"); self.assertEqual(self.api._user, "alice")
        self.api.lock_session(); self.assertIsNone(self.api._user)

    def test_install_from_before_accounts_still_signs_in_with_pin_only(self):
        salt = "ab" * 16
        json.dump({"hash": hashlib.sha256((salt + "4321").encode()).hexdigest(), "salt": salt}, open(AUTH(), "w"))
        r = J(self.api.is_pin_set()); self.assertTrue(r["set"]); self.assertFalse(r["has_username"])
        v = J(self.api.verify_pin("4321")); self.assertTrue(v["valid"]); self.assertNotIn("username", v)
        J(self.api.set_pin("8765"))                                               # nothing is forced on it
        self.assertNotIn("username", json.load(open(AUTH())))
        self.api.lock_session(); self.assertTrue(J(self.api.verify_pin("8765"))["valid"])

    def test_backup_restore_keeps_the_account(self):
        self.api.set_pin("1234", "alice")
        b64 = J(self.api.export_full_backup())["b64"]
        os.remove(AUTH()); self.api.lock_session()
        self.assertTrue(J(self.api.restore_full_backup(b64))["ok"])
        self.api.lock_session(); self.assertTrue(J(self.api.verify_pin("1234", "alice"))["valid"])


class SampleData(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()                                 # seeded with the shipped sample

    def data(self): return J(self.api.load_data())["data"]
    def serials(self): return sorted(r["Serial No."] for r in self.data()["inventory"])
    def snapshots(self):
        d = os.path.join(TMP, "backups"); return [f for f in os.listdir(d) if f.endswith(".db")] if os.path.isdir(d) else []

    def test_fresh_install_has_exactly_the_six_sample_assets(self):
        self.assertEqual(self.serials(), [f"SAMPLE-000{i}" for i in range(1, 7)])

    def test_remove_deletes_only_sample_rows_and_saves_a_safety_copy(self):
        J(self.api.add_asset(json.dumps(asset(**{"Serial No.": "REAL-1"}))))
        J(self.api.add_asset(json.dumps(asset(**{"Serial No.": "sample-lower"}))))   # not an exact SAMPLE- prefix
        before = len(self.snapshots())
        r = J(self.api.remove_sample_data()); self.assertEqual((r["ok"], r["removed"]), (True, 6))
        self.assertEqual(self.serials(), ["REAL-1", "sample-lower"])
        self.assertGreater(len(self.snapshots()), before)

    def test_history_of_sample_assets_goes_but_the_audit_log_stays(self):
        a = self.data()["inventory"][0]
        J(self.api.transfer_asset(a["_row"], "x", a["Serial No."], a["Department/Location"], "Some Other Dept", "n", "2026-01-01"))
        self.assertEqual(len(self.data()["history"]), 1)
        log_before = len(self.data()["log"])
        self.api.remove_sample_data()
        d = self.data(); self.assertEqual(len(d["history"]), 0); self.assertGreater(len(d["log"]), log_before)

    def test_remove_is_idempotent_and_takes_no_pointless_copy(self):
        self.api.remove_sample_data(); n = len(self.snapshots())
        r = J(self.api.remove_sample_data()); self.assertEqual((r["ok"], r["removed"]), (True, 0))
        self.assertEqual(len(self.snapshots()), n)

    def test_samples_do_not_come_back_after_a_restart(self):
        self.api.remove_sample_data(); J(self.api.add_asset(json.dumps(asset(**{"Serial No.": "REAL-1"}))))
        api2 = app.Api(storage.Storage(TMP, os.path.join(TMP, "backups"), seed_xlsx=SEED), app.ConfigEngine())
        self.assertEqual([r["Serial No."] for r in J(api2.load_data())["data"]["inventory"]], ["REAL-1"])

    def test_empty_database_still_loads_everything(self):
        self.api.remove_sample_data(); d = self.data()
        self.assertEqual((d["inventory"], d["history"]), ([], []))


if __name__ == "__main__":
    unittest.main()
