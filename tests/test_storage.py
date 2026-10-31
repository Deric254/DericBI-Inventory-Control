"""Storage engine: ACID guarantees, verified migration, backups, export, performance.
Run:  python -m unittest discover -s tests -v
"""
import os, sys, json, time, shutil, sqlite3, subprocess, threading, unittest, datetime
from _env import app, storage, TMP, SRC, SEED, fresh, wipe, asset
import openpyxl


def raw(sql, args=()):
    c = sqlite3.connect(os.path.join(TMP, "inventory.db")); c.row_factory = sqlite3.Row
    try: return [dict(r) for r in c.execute(sql, args)]
    finally: c.close()

def count(table): return raw(f"SELECT COUNT(*) n FROM {table}")[0]["n"]


# ───────────────────────────── migration ─────────────────────────────
class Migration(unittest.TestCase):
    def test_fresh_install_imports_sample_and_leaves_no_clutter(self):
        eng, _ = fresh()
        inv = eng.load_all()["inventory"]
        self.assertEqual(len(inv), 6)
        self.assertTrue(all(r["Serial No."].startswith("SAMPLE-") for r in inv))
        self.assertEqual([r["_row"] for r in inv], list(range(1, 7)))
        self.assertFalse(os.path.exists(os.path.join(TMP, "ICT_MASTER.xlsx")))
        self.assertFalse(os.path.exists(os.path.join(TMP, "ICT_MASTER.pre-sqlite.xlsx")))
        self.assertEqual(eng.load_all()["log"], [])               # no noise in a new install's audit log

    def _hostile_workbook(self, path):
        wb = openpyxl.Workbook(); wb.remove(wb.active)
        ws = wb.create_sheet("Inventory")
        # shuffled columns, 'Transfer History' column missing entirely, extra unknown column
        ws.append(["Serial No.", "Equipment Type", "Mystery", "Department/Location", "Date Collected", "Brand/Model", "RAM", "Condition/Status", "Remarks"])
        ws.append(["SN-1", "Laptop", "ignored", "Pharmacy", datetime.datetime(2025, 3, 4), "HP", 8, "Good", "ok"])
        ws.append([None] * 9)                                                          # blank row
        ws.append(["SN-1", "Printer", "x", "Lab", "2025-01-02 00:00:00", "Canon", "nan", "Fair", "dup serial"])
        ws.append(["n/a", "Router", "", "ICT", None, "MikroTik", None, None, "é ü 日本語 🚀"])
        ws.append(["n/a", "Switch", "", "ICT", None, "TP", None, None, "x" * 20000])
        ws.append(["SN-9", "Server", "", "ICT", None, "Dell", None, "Good", "=HYPERLINK(\"http://evil\")"])
        ws.cell(ws.max_row, 9).data_type = "s"                       # a TEXT cell that merely starts with '='
        h = wb.create_sheet("History"); h.append(list(storage.HIST_HEADERS))
        h.append(["2025-01-01", "Laptop HP", "SN-1", "A", "B", "transferred", "note", "ICT Manager"])
        h.append(["2025-01-02", "Laptop HP", "SN-1", "B", "C", "transferred", "note2", "ICT Manager"])
        k = wb.create_sheet("Ink"); k.append(list(storage.INK_HEADERS))
        k.append(["2025-01-01 10:00:00", "BK", "Black", "restock", 5, None, 0, ""])
        k.append(["2025-01-02 10:00:00", "BK", "Black", "move_to_ict", 3, 3, 0, ""])
        k.append(["2025-01-03 10:00:00", "BK", "Black", "use", 1, 2, 0, ""])
        l = wb.create_sheet("Audit Log"); l.append(list(storage.LOG_HEADERS))
        l.append(["2025-01-01 09:00:00", "ADD", "Added thing", "ICT Manager"])
        l.append([datetime.datetime(2025, 1, 2, 9, 30), "EDIT", "Edited thing", "Bob"])
        wb.create_sheet("Replacements").append(list(storage.REPL_HEADERS))
        wb.create_sheet("Config").append(["Key", "Value"])
        wb.save(path)

    def test_hostile_excel_is_migrated_faithfully_and_archived(self):
        wipe(); p = os.path.join(TMP, "ICT_MASTER.xlsx"); self._hostile_workbook(p)
        eng = storage.Storage(TMP, os.path.join(TMP, "backups"))
        d = eng.load_all(); inv = d["inventory"]
        self.assertEqual(len(inv), 5)
        self.assertEqual([r["_row"] for r in inv], [1, 2, 3, 4, 5])
        self.assertEqual(inv[0]["Date Collected"], "2025-03-04")
        self.assertEqual(inv[0]["RAM"], "8")
        self.assertEqual(inv[1]["RAM"], "")                       # 'nan' -> blank
        self.assertEqual(inv[1]["Date Collected"], "2025-01-02")  # '2025-01-02 00:00:00' -> date
        self.assertEqual(inv[2]["Date Collected"], "")
        self.assertEqual(inv[1]["Condition/Status"], "Fair")
        self.assertEqual(inv[2]["Remarks"], "é ü 日本語 🚀")
        self.assertEqual(len(inv[3]["Remarks"]), 20000)
        self.assertEqual(inv[4]["Remarks"], "=HYPERLINK(\"http://evil\")")      # stays text
        self.assertTrue(all(r["Transfer History"] == "" for r in inv))
        self.assertEqual(d["ink"]["BK"], {"store": 9 + 5 - 3, "ictRoom": 2})
        self.assertEqual(len(d["history"]), 2)
        self.assertEqual(d["history"][0]["note"], "note2")                      # newest first
        self.assertEqual(d["log"][0]["action"], "MIGRATE")
        self.assertIn("1 duplicate serial", d["log"][0]["desc"])
        self.assertEqual([x["time"] for x in d["log"][1:]], ["2025-01-02 09:30:00", "2025-01-01 09:00:00"])
        # original is kept (never deleted) + a verified copy
        self.assertFalse(os.path.exists(p))
        self.assertTrue(os.path.exists(os.path.join(TMP, "ICT_MASTER.pre-sqlite.xlsx")))
        self.assertTrue(os.listdir(os.path.join(TMP, "backups", "pre-sqlite")))
        # restart must NOT import again
        n = count("inventory"); storage.Storage(TMP, os.path.join(TMP, "backups")); self.assertEqual(count("inventory"), n)

    def test_oldest_worksheet_format(self):
        wipe(); wb = openpyxl.Workbook(); ws = wb.active; ws.title = "work sheet"
        ws.append(["Equipment Type", "Brand/Model", "Serial No.", "RAM Size", "RAM Type", "Disk Size", "Disk Type", "IP Address", "Department/Location"])
        ws.append(["Desktop", "Dell", "OLD-1", "8GB", "DDR4", "500GB", "HDD", "192.068.1.5", " Lab "])
        wb.save(os.path.join(TMP, "ICT_MASTER.xlsx"))
        r = storage.Storage(TMP).load_all()["inventory"][0]
        self.assertEqual((r["RAM"], r["Disk"], r["IP Address"], r["Department/Location"], r["Condition/Status"]),
                         ("8GB DDR4", "500GB HDD", "192.168.1.5", "Lab", "Good"))

    def test_corrupt_excel_aborts_cleanly_and_can_be_retried(self):
        wipe(); p = os.path.join(TMP, "ICT_MASTER.xlsx"); open(p, "wb").write(os.urandom(5000))
        before = open(p, "rb").read()
        with self.assertRaises(storage.StorageError) as cm: storage.Storage(TMP)
        self.assertIn("NOT changed", str(cm.exception))
        self.assertEqual(open(p, "rb").read(), before)            # untouched
        self._hostile_workbook(p)                                  # fix it -> retry works
        self.assertEqual(len(storage.Storage(TMP).load_all()["inventory"]), 5)

    def test_workbook_without_inventory_data_columns_is_rejected(self):
        wipe(); wb = openpyxl.Workbook(); wb.active.title = "Inventory"; wb.active.append(["Foo", "Bar"]); wb.active.append([1, 2])
        wb.save(os.path.join(TMP, "ICT_MASTER.xlsx"))
        with self.assertRaises(storage.StorageError): storage.Storage(TMP)


# ───────────────────────────── ACID ─────────────────────────────
class Atomicity(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()

    def _trigger(self, table):
        c = sqlite3.connect(os.path.join(TMP, "inventory.db"))
        c.execute(f"CREATE TRIGGER boom BEFORE INSERT ON {table} BEGIN SELECT RAISE(ABORT,'injected failure'); END"); c.commit(); c.close()

    def test_failed_history_write_rolls_back_the_transfer(self):
        rid = self.eng.add_asset(asset(**{"Serial No.": "A-1", "Department/Location": "Lab"}))
        self._trigger("history")
        with self.assertRaises(Exception): self.eng.transfer_asset(rid, "x", "A-1", "Lab", "Pharmacy", "n", "2026-01-01")
        r = raw('SELECT "Department/Location" d, "Transfer History" h FROM inventory WHERE id=?', (rid,))[0]
        self.assertEqual((r["d"], r["h"]), ("Lab", ""))          # inventory update was undone too
        self.assertFalse(any(x["action"] == "TRANSFER" for x in self.eng.load_all()["log"]))

    def test_failed_audit_log_rolls_back_the_new_asset(self):
        n = count("inventory"); self._trigger("log")
        with self.assertRaises(Exception): self.eng.add_asset(asset())
        self.assertEqual(count("inventory"), n)

    def test_failed_replacement_leaves_old_asset_untouched(self):
        rid = self.eng.add_asset(asset(**{"Serial No.": "R-1"})); self._trigger("replacements")
        with self.assertRaises(Exception):
            self.eng.record_replacement(rid, "Laptop", {"date": "2026-01-01", "note": "dead", "newSerial": "R-2", "newBrand": "N"})
        self.assertEqual(raw('SELECT "Condition/Status" c FROM inventory WHERE id=?', (rid,))[0]["c"], "Good")
        self.assertEqual(raw('SELECT COUNT(*) n FROM inventory WHERE "Serial No."=?', ("R-2",))[0]["n"], 0)

    def test_restore_is_all_or_nothing(self):
        self.eng.add_asset(asset(**{"Serial No.": "KEEP"}))
        bad = os.path.join(TMP, "bad.db"); open(bad, "wb").write(b"not a database at all, just text" * 50)
        with self.assertRaises(ValueError): self.eng.restore_from_db(bad)
        self.assertTrue(any(r["Serial No."] == "KEEP" for r in self.eng.load_all()["inventory"]))


class Consistency(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()

    def test_duplicate_serials_rejected_everywhere(self):
        self.eng.add_asset(asset(**{"Serial No.": "DUP-1"}))
        with self.assertRaises(ValueError): self.eng.add_asset(asset(**{"Serial No.": " dup-1 "}))      # case/space-insensitive
        other = self.eng.add_asset(asset(**{"Serial No.": "DUP-2"}))
        with self.assertRaises(ValueError): self.eng.update_asset(other, {"Serial No.": "DUP-1"})
        self.eng.update_asset(other, {"Serial No.": "DUP-2", "Remarks": "same serial, fine"})            # own serial is fine
        with self.assertRaises(ValueError):
            self.eng.record_replacement(other, "x", {"date": "2026-01-01", "note": "", "newSerial": "DUP-1"})

    def test_placeholder_serials_and_retired_assets_may_repeat(self):
        for _ in range(3): self.eng.add_asset(asset(**{"Serial No.": "N/A"}))
        for _ in range(2): self.eng.add_asset(asset(**{"Serial No.": ""}))
        rid = self.eng.add_asset(asset(**{"Serial No.": "OLD-1"}))
        self.eng.write_history_event("x", "OLD-1", "ICT", "disposed", "scrapped", "2026-01-01")
        self.eng.add_asset(asset(**{"Serial No.": "OLD-1"}))                                              # disposed serial can be reused

    def test_required_fields_and_limits(self):
        with self.assertRaises(ValueError): self.eng.add_asset(asset(**{"Equipment Type": "  "}))
        with self.assertRaises(ValueError): self.eng.add_asset(asset(Remarks="x" * 40000))
        rid = self.eng.add_asset(asset())
        with self.assertRaises(ValueError): self.eng.update_asset(rid, {"Equipment Type": ""})
        self.assertEqual(self.eng.load_all()["inventory"][-1]["Equipment Type"], "Laptop")

    def test_ids_are_stable_and_never_reused(self):
        a = self.eng.add_asset(asset(**{"Serial No.": "I-1"})); b = self.eng.add_asset(asset(**{"Serial No.": "I-2"}))
        self.eng.delete_asset(a, "")
        self.assertEqual(self.eng.load_all()["inventory"][-1]["_row"], b)                                  # b did not shift
        c = self.eng.add_asset(asset(**{"Serial No.": "I-3"})); self.assertGreater(c, b)
        for bad in (0, -1, 10**9, "abc", None, True):
            with self.assertRaises(ValueError): self.eng.update_asset(bad, {"Remarks": "x"})
            with self.assertRaises(ValueError): self.eng.delete_asset(bad, "x")

    def test_transfer_rules_and_history(self):
        rid = self.eng.add_asset(asset(**{"Serial No.": "TR-1", "Department/Location": "Lab"}))
        with self.assertRaises(ValueError): self.eng.transfer_asset(rid, "x", "TR-1", "Lab", "lab", "", "2026-01-01")
        with self.assertRaises(ValueError): self.eng.transfer_asset(rid, "x", "TR-1", "Lab", "", "", "2026-01-01")
        self.eng.transfer_asset(rid, "Laptop Test", "TR-1", "WRONG-CLIENT-VALUE", "Pharmacy", "moved", "2026-01-01")
        self.eng.transfer_asset(rid, "Laptop Test", "TR-1", "Pharmacy", "Maternity", "", "2026-02-01")
        a = [r for r in self.eng.load_all()["inventory"] if r["_row"] == rid][0]
        self.assertEqual(a["Department/Location"], "Maternity")
        self.assertIn("Lab → Pharmacy: moved", a["Transfer History"]); self.assertIn("Pharmacy → Maternity", a["Transfer History"])
        tl = self.eng.get_transfer_log()
        self.assertEqual([t["to"] for t in tl], ["Maternity", "Pharmacy"])
        self.assertEqual(tl[1]["from"], "Lab")                   # history records the TRUE previous department

    def test_history_events_update_assets(self):
        rid = self.eng.add_asset(asset(**{"Serial No.": "H-1", "Condition/Status": "Fair (Needs Repair)"}))
        g = lambda: [r for r in self.eng.load_all()["inventory"] if r["_row"] == rid][0]
        self.eng.write_history_event("x", "H-1", "ICT", "repaired", "fixed", "2026-01-01", "1500"); self.assertEqual(g()["Condition/Status"], "Good")
        self.assertIn("Cost: KES 1500", self.eng.load_all()["history"][0]["note"])                         # cost is no longer lost
        self.eng.update_asset(rid, {"Assigned To": "Jane Doe"})
        # THE OLD BUG: editing the department logs this note; it must NOT touch 'Assigned To'
        self.eng.write_history_event("x", "H-1", "Lab", "reassigned", "Department changed: ICT → Lab", "2026-01-02"); self.assertEqual(g()["Assigned To"], "Jane Doe")
        self.eng.write_history_event("x", "H-1", "Lab", "reassigned", "Status changed: Good → Fair", "2026-01-02"); self.assertEqual(g()["Assigned To"], "Jane Doe")
        self.eng.write_history_event("x", "H-1", "Lab", "reassigned", "Assignment changed: Jane Doe → Peter Kamau", "2026-01-02"); self.assertEqual(g()["Assigned To"], "Peter Kamau")
        self.eng.write_history_event("x", "H-1", "Lab", "reassigned", "Assignment changed: Peter Kamau → (none)", "2026-01-02"); self.assertEqual(g()["Assigned To"], "")
        self.eng.write_history_event("x", "H-1", "Lab", "reassigned", "Reassigned to John Smith [by ICT]", "2026-01-02"); self.assertEqual(g()["Assigned To"], "John Smith")
        self.eng.write_history_event("x", "H-1", "Lab", "reassigned", "Monitor moved next door", "2026-01-02"); self.assertEqual(g()["Assigned To"], "John Smith")   # 'Monitor' contains 'to'
        self.eng.update_asset(rid, {"Assigned To": "Jane Doe"})
        self.eng.write_history_event("x", "", "ICT", "disposed", "no serial given", "2026-01-03")           # blank serial never matches other assets
        self.assertNotEqual(g()["Condition/Status"], "Disposed")

    def test_replacement_flow(self):
        rid = self.eng.add_asset(asset(**{"Serial No.": "RP-1", "Department/Location": "Lab", "Remarks": "old"}))
        self.eng.record_replacement(rid, "Laptop Test", {"date": "2026-03-01", "note": "dead psu", "newType": "", "newBrand": "Lenovo",
                                                         "newSerial": "RP-2", "newCond": "New", "cost": "45000"})
        inv = {r["Serial No."]: r for r in self.eng.load_all()["inventory"]}
        self.assertEqual(inv["RP-1"]["Condition/Status"], "Replaced"); self.assertIn("REPLACED 2026-03-01: dead psu", inv["RP-1"]["Remarks"])
        self.assertEqual((inv["RP-2"]["Equipment Type"], inv["RP-2"]["Department/Location"], inv["RP-2"]["Condition/Status"]), ("Laptop", "Lab", "New"))
        self.assertIn("45000", inv["RP-2"]["Remarks"])
        self.assertEqual(count("replacements"), 1); self.assertEqual(self.eng.load_all()["history"][0]["event"], "replaced")

    def test_ink_rules(self):
        st = lambda: self.eng.load_all()["ink"]["BK"]
        self.assertEqual(st(), {"store": 9, "ictRoom": 0})
        self.eng.ink_apply("BK", "move_to_ict", 4, 4, ""); self.assertEqual(st(), {"store": 5, "ictRoom": 4})
        self.eng.ink_apply("BK", "use", 6, 0, "");          self.assertEqual(st(), {"store": 3, "ictRoom": 0})
        for args in (("BK", "use", 99, 0), ("BK", "move_to_ict", 99, 0), ("ZZ", "use", 1, 0), ("BK", "explode", 1, 0), ("BK", "use", -1, 0)):
            with self.assertRaises(ValueError): self.eng.ink_apply(*args)
        self.assertEqual(st(), {"store": 3, "ictRoom": 0})

    def test_user_text_is_sanitised_not_executed(self):
        rid = self.eng.add_asset(asset(Remarks="=1+1\x00\x07 ok \ud800", **{"Assigned To": "'; DROP TABLE inventory;--"}))
        r = [x for x in self.eng.load_all()["inventory"] if x["_row"] == rid][0]
        self.assertEqual(r["Remarks"], "=1+1 ok ?"); self.assertEqual(r["Assigned To"], "'; DROP TABLE inventory;--")
        self.assertGreater(count("inventory"), 0)


# ───────────────────────────── isolation ─────────────────────────────
class Isolation(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh(seed=False)

    def test_many_threads_never_lose_or_duplicate_rows(self):
        errs = []
        def work(t):
            try:
                for i in range(15): self.eng.add_asset(asset(**{"Serial No.": f"TH-{t}-{i}"}))
            except Exception as e: errs.append(repr(e))
        ts = [threading.Thread(target=work, args=(t,)) for t in range(8)]; [t.start() for t in ts]; [t.join() for t in ts]
        self.assertEqual(errs, []); self.assertEqual(count("inventory"), 120)
        self.assertEqual(len({r["Serial No."] for r in self.eng.load_all()["inventory"]}), 120)
        self.assertEqual(sum(1 for l in self.eng.load_all()["log"] if l["action"] == "ADD"), 120)

    def test_concurrent_writers_never_break_backups(self):
        import io, contextlib
        err = io.StringIO()
        with contextlib.redirect_stderr(err):
            ts = [threading.Thread(target=lambda t=t: [self.eng.add_asset(asset(**{"Serial No.": f"BK-{t}-{i}"})) for i in range(6)]) for t in range(8)]
            [t.start() for t in ts]; [t.join() for t in ts]
        self.assertNotIn("snapshot skipped", err.getvalue())
        snaps = [f for f in os.listdir(os.path.join(TMP, "backups")) if f.endswith(".db")]
        self.assertGreaterEqual(len(snaps), 1)
        self.assertEqual([f for f in os.listdir(os.path.join(TMP, "backups")) if f.endswith(".tmp")], [])
        for f in snaps:
            c = sqlite3.connect(os.path.join(TMP, "backups", f)); self.assertEqual(c.execute("PRAGMA integrity_check").fetchone()[0], "ok"); c.close()

    def test_concurrent_duplicate_serial_only_one_wins(self):
        res = []
        def go():
            try: self.eng.add_asset(asset(**{"Serial No.": "RACE-1"})); res.append("ok")
            except ValueError: res.append("dup")
        ts = [threading.Thread(target=go) for _ in range(10)]; [t.start() for t in ts]; [t.join() for t in ts]
        self.assertEqual((res.count("ok"), res.count("dup")), (1, 9))

    def test_ink_cannot_be_oversold_by_concurrent_users(self):
        res = []
        def use():
            try: self.eng.ink_apply("BK", "use", 1, 0, ""); res.append(1)
            except ValueError: res.append(0)
        ts = [threading.Thread(target=use) for _ in range(25)]; [t.start() for t in ts]; [t.join() for t in ts]
        self.assertEqual(sum(res), 9)                                     # stock was 9
        self.assertEqual(self.eng.load_all()["ink"]["BK"], {"store": 0, "ictRoom": 0})

    def test_separate_processes_share_one_database_safely(self):
        code = ("import sys; sys.path.insert(0, %r); import storage\n"
                "e = storage.Storage(%r)\n"
                "for i in range(25): e.add_asset({'Equipment Type':'Laptop','Serial No.': sys.argv[1]+'-'+str(i)})\n") % (SRC, TMP)
        ps = [subprocess.Popen([sys.executable, "-c", code, f"P{n}"], stderr=subprocess.PIPE) for n in range(4)]
        for p in ps: self.assertEqual(p.wait(timeout=120), 0, p.stderr.read().decode())
        self.assertEqual(count("inventory"), 100)
        self.assertEqual(raw("PRAGMA integrity_check")[0]["integrity_check"], "ok")


# ───────────────────────────── durability ─────────────────────────────
class Durability(unittest.TestCase):
    def test_kill_9_mid_write_never_corrupts_or_loses_committed_data(self):
        wipe()
        code = ("import sys; sys.path.insert(0, %r); import storage\n"
                "e = storage.Storage(%r)\ni = 0\n"
                "while True:\n    print(e.add_asset({'Equipment Type':'Laptop','Serial No.':sys.argv[1]+'-'+str(i)}), flush=True); i += 1\n") % (SRC, TMP)
        committed = []
        for rnd in range(3):                                              # three crashes in a row
            p = subprocess.Popen([sys.executable, "-c", code, f"K{rnd}"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            got = 0
            while got < 25 + rnd * 7:
                line = p.stdout.readline()
                if not line: self.fail("writer died early: " + p.stderr.read().decode())
                committed.append(int(line)); got += 1
            p.kill(); p.wait()
            eng = storage.Storage(TMP)                                    # re-open == crash recovery
            ids = {r["_row"] for r in eng.load_all()["inventory"]}
            self.assertTrue(set(committed) <= ids, "a committed row was lost")
            self.assertEqual(raw("PRAGMA integrity_check")[0]["integrity_check"], "ok")
            adds = sum(1 for l in eng.load_all()["log"] if l["action"] == "ADD")
            self.assertEqual(adds, len(ids), "an asset and its audit entry must always exist together")
        self.assertGreaterEqual(count("inventory"), len(committed))


# ───────────────────────────── backups / export ─────────────────────────────
class BackupsAndExport(unittest.TestCase):
    def setUp(self): self.eng, self.api = fresh()

    def test_snapshot_is_valid_and_restorable(self):
        self.eng.add_asset(asset(**{"Serial No.": "SNAP-1"}))
        snap = self.eng.snapshot(force=True); self.assertTrue(snap and os.path.exists(snap))
        self.eng.add_asset(asset(**{"Serial No.": "AFTER"}))
        self.eng.restore_from_db(snap)
        sers = {r["Serial No."] for r in self.eng.load_all()["inventory"]}
        self.assertIn("SNAP-1", sers); self.assertNotIn("AFTER", sers)
        self.assertTrue(os.listdir(os.path.join(TMP, "backups", "pre-restore")))     # the 'undo' copy of what was replaced
        self.assertEqual(self.eng.load_all()["log"][0]["action"], "RESTORE")
        rid = self.eng.add_asset(asset(**{"Serial No.": "NEW-AFTER-RESTORE"})); self.assertGreater(rid, 7)

    def test_destructive_actions_snapshot_first(self):
        rid = self.eng.add_asset(asset())
        before = len(os.listdir(os.path.join(TMP, "backups")))
        self.eng.delete_asset(rid, "x")
        self.assertGreater(len(os.listdir(os.path.join(TMP, "backups"))), before - 1)
        self.assertTrue(any(b["name"].endswith(".db") for b in self.eng.list_backups()))

    def test_auto_snapshots_are_rate_limited(self):
        for i in range(5): self.eng.add_asset(asset(**{"Serial No.": f"RL-{i}"}))
        self.assertLessEqual(len([f for f in os.listdir(os.path.join(TMP, "backups")) if f.endswith(".db")]), 1)

    def test_excel_export_roundtrips_exactly(self):
        self.eng.add_asset(asset(Remarks="=SUM(A1:A2)", **{"Serial No.": "X-1", "Assigned To": "José 日本"}))
        self.eng.transfer_asset(7, "Laptop", "X-1", "ICT", "Lab", "n", "2026-01-01"); self.eng.ink_apply("C", "restock", 2, 0, "")
        p = self.eng.export_for_user()
        wb = openpyxl.load_workbook(p)                                    # formulas NOT evaluated here
        c = [x for x in wb["Inventory"][8] if x.value == "=SUM(A1:A2)"][0]
        self.assertEqual(c.data_type, "s")                                # text, not a live formula
        want = self.eng.load_all()
        import tempfile
        e2 = storage.Storage(tempfile.mkdtemp(prefix="dericbi_rt_")); e2.restore_from_xlsx(p); got = e2.load_all()
        strip = lambda d: [{k: v for k, v in r.items() if k != "_row"} for r in d["inventory"]]
        self.assertEqual(strip(got), strip(want))
        self.assertEqual(got["ink"], want["ink"])
        self.assertEqual([h["event"] for h in got["history"]], [h["event"] for h in want["history"]])

    def test_export_survives_illegal_characters_in_old_data(self):
        c = sqlite3.connect(os.path.join(TMP, "inventory.db"))
        c.execute('UPDATE inventory SET "Remarks"=? WHERE id=1', ("bad\x01\x02char",)); c.commit(); c.close()
        self.assertTrue(os.path.exists(self.eng.export_for_user()))

    def test_exports_are_pruned(self):
        for _ in range(7): self.eng.export_for_user(); time.sleep(1.05)
        self.assertLessEqual(len(os.listdir(os.path.join(TMP, "exports"))), 5)

    def test_corrupt_database_is_reported_not_silently_used(self):
        self.eng.snapshot(force=True)
        del self.eng
        db = os.path.join(TMP, "inventory.db")
        for ext in ("-wal", "-shm"):
            if os.path.exists(db + ext): os.remove(db + ext)
        data = bytearray(open(db, "rb").read())
        for i in range(4096 * 2, min(len(data), 4096 * 5)): data[i] = 0xFF
        open(db, "wb").write(bytes(data))
        with self.assertRaises(Exception): storage.Storage(TMP)


# ───────────────────────────── performance ─────────────────────────────
class Performance(unittest.TestCase):
    N = 20000

    def test_writes_stay_fast_with_a_large_database(self):
        eng, _ = fresh(seed=False)
        c = sqlite3.connect(os.path.join(TMP, "inventory.db"))
        cols = ",".join(f'"{x}"' for x in storage.INV_COLS)
        c.executemany(f"INSERT INTO inventory({cols}) VALUES({','.join('?' * 16)})",
                      [("2025-01-01", "", "Laptop", f"Model {i}", f"PERF-{i:07d}", "User", f"Dept{i % 20}", "Good", "", "", "", "", "", "", "r", "") for i in range(self.N)])
        c.commit(); c.close()
        def t(fn):
            s = time.perf_counter(); fn(); return time.perf_counter() - s
        add = t(lambda: eng.add_asset(asset(**{"Serial No.": "PERF-NEW"})))
        upd = t(lambda: eng.update_asset(self.N // 2, {"Remarks": "edited"}))
        tr  = t(lambda: eng.transfer_asset(self.N // 3, "x", "", "", "NewDept", "n", "2026-01-01"))
        dup = t(lambda: self.assertRaises(ValueError, eng.add_asset, asset(**{"Serial No.": "perf-0000100"})))
        first = t(eng.load_all_json); again = t(eng.load_all_json)
        print(f"\n   [{self.N:,} assets] add={add*1000:.0f}ms update={upd*1000:.0f}ms transfer={tr*1000:.0f}ms dup-check={dup*1000:.0f}ms "
              f"full-load={first*1000:.0f}ms cached={again*1000:.1f}ms")
        for name, v in (("add", add), ("update", upd), ("transfer", tr), ("dup", dup)):
            self.assertLess(v, 0.5, f"{name} too slow: {v:.3f}s")
        self.assertLess(again, 0.05)


if __name__ == "__main__":
    unittest.main()
