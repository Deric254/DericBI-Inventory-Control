"""
storage.py - DericBI Inventory Control storage engine (SQLite).

Why SQLite: every operation is a real ACID transaction.
  Atomic     - an operation (e.g. a transfer = inventory update + history + audit log)
               either fully happens or not at all, even if the app is killed mid-write.
  Consistent - constraints are enforced inside the same transaction (unique serials,
               required fields, valid ink stock, existing asset ids).
  Isolated   - SQLite locks across threads AND processes; two writers can never
               overwrite each other (no lost updates).
  Durable    - WAL journal + synchronous=FULL: once a call returns, the data is on disk.

Excel is still supported as an import source (one-time, verified migration) and as an
export format ("Export to Excel"), but it is no longer the live database.
"""
import os, re, sys, json, time, uuid, shutil, sqlite3, threading, datetime, contextlib, pathlib

SCHEMA_VERSION = 1
DB_NAME        = "inventory.db"
LEGACY_XLSX    = "ICT_MASTER.xlsx"
MAX_FIELD_CHARS = 32000

SHEET_INV, SHEET_HISTORY, SHEET_REPLACE = "Inventory", "History", "Replacements"
SHEET_INK, SHEET_LOG, SHEET_CONFIG      = "Ink", "Audit Log", "Config"
ALL_SHEETS = [SHEET_INV, SHEET_HISTORY, SHEET_REPLACE, SHEET_INK, SHEET_LOG, SHEET_CONFIG]

INV_COLS = [
    "Date Collected", "Purchase Date", "Equipment Type", "Brand/Model", "Serial No.",
    "Assigned To", "Department/Location", "Condition/Status", "OS/Firmware",
    "IP Address", "MAC Address", "CPU", "RAM", "Disk", "Remarks", "Transfer History",
]
HIST_HEADERS = ["Date", "Asset", "Serial No.", "From Department", "To Department", "Event", "Description", "By"]
REPL_HEADERS = ["Date Replaced", "Old Asset", "Old Serial", "Department", "Replaced With", "New Serial", "New Condition", "Note", "By"]
INK_HEADERS  = ["Timestamp", "Colour Code", "Colour Name", "Action", "Quantity", "Big ICT", "Small ICT", "Note"]
LOG_HEADERS  = ["Timestamp", "Action", "Description", "By"]

INK_CODES    = ["BK", "C", "LC", "M", "LM", "Y"]
INK_NAMES    = {"BK": "Black", "C": "Cyan", "LC": "Light Cyan", "M": "Magenta", "LM": "Light Magenta", "Y": "Yellow"}
INK_DEFAULTS = {"BK": 9, "C": 3, "LC": 5, "M": 1, "LM": 5, "Y": 11}
INK_ACTIONS  = {"restock", "move_to_ict", "use", "set_printer"}

# Serial numbers that mean "no serial" - never treated as duplicates of each other.
NO_SERIAL = {"", "n/a", "na", "n.a", "none", "nil", "null", "-", "--", "---", "unknown", "?",
             "not available", "unavailable", "no serial", "tbd", "nan"}
CLOSED_CONDITIONS = ("Disposed", "Replaced")      # retired assets may reuse a serial

TABLES = ["inventory", "history", "replacements", "ink", "log"]

_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def today(): return datetime.date.today().strftime("%Y-%m-%d")
def now():   return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def txt(v):
    """Normalise any value to clean text safe for SQLite and Excel."""
    if v is None: return ""
    if isinstance(v, bool): return "Yes" if v else "No"
    if isinstance(v, (datetime.datetime,)): return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.date): return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer(): v = int(v)
    s = str(v)
    s = s.encode("utf-8", "replace").decode("utf-8")      # lone surrogates -> '?'
    return _CTRL_RE.sub("", s).strip()


def clean_cell(v):
    """Excel cell -> text (dates as YYYY-MM-DD, 'nan'/'None' placeholders -> blank)."""
    if v is None: return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.date): return v.strftime("%Y-%m-%d")
    s = txt(v)
    if re.match(r"\d{4}-\d{2}-\d{2}[ T]", s) and re.fullmatch(r"\d{4}-\d{2}-\d{2}[ T]00:00:00(\.0+)?", s): s = s[:10]
    return "" if s.lower() in ("nan", "none", "nat", "<na>") else s


def _int(v, default=0):
    try: return int(float(str(v)))
    except Exception: return default


def _int_or_none(v):
    try: return int(float(str(v)))
    except Exception: return None


def _schema():
    inv = ",\n  ".join(f'"{c}" TEXT NOT NULL DEFAULT \'\'' for c in INV_COLS)
    return [
        "CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT NOT NULL)",
        f"CREATE TABLE IF NOT EXISTS inventory(\n  id INTEGER PRIMARY KEY AUTOINCREMENT,\n  {inv}\n)",
        """CREATE TABLE IF NOT EXISTS history(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  date TEXT NOT NULL DEFAULT '', asset TEXT NOT NULL DEFAULT '', serial TEXT NOT NULL DEFAULT '',
  from_dept TEXT NOT NULL DEFAULT '', to_dept TEXT NOT NULL DEFAULT '', event TEXT NOT NULL DEFAULT '',
  description TEXT NOT NULL DEFAULT '', actor TEXT NOT NULL DEFAULT '')""",
        """CREATE TABLE IF NOT EXISTS replacements(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  date TEXT NOT NULL DEFAULT '', old_asset TEXT NOT NULL DEFAULT '', old_serial TEXT NOT NULL DEFAULT '',
  dept TEXT NOT NULL DEFAULT '', new_desc TEXT NOT NULL DEFAULT '', new_serial TEXT NOT NULL DEFAULT '',
  new_cond TEXT NOT NULL DEFAULT '', note TEXT NOT NULL DEFAULT '', actor TEXT NOT NULL DEFAULT '')""",
        """CREATE TABLE IF NOT EXISTS ink(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL DEFAULT '', code TEXT NOT NULL DEFAULT '', name TEXT NOT NULL DEFAULT '',
  action TEXT NOT NULL DEFAULT '', qty INTEGER NOT NULL DEFAULT 0 CHECK(qty >= 0),
  ict_room INTEGER, aux INTEGER NOT NULL DEFAULT 0, note TEXT NOT NULL DEFAULT '')""",
        """CREATE TABLE IF NOT EXISTS log(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL DEFAULT '', action TEXT NOT NULL DEFAULT '',
  description TEXT NOT NULL DEFAULT '', actor TEXT NOT NULL DEFAULT '')""",
        'CREATE INDEX IF NOT EXISTS ix_inv_serial ON inventory(lower(trim("Serial No.")))',
        'CREATE INDEX IF NOT EXISTS ix_inv_dept ON inventory("Department/Location")',
        'CREATE INDEX IF NOT EXISTS ix_inv_type ON inventory("Equipment Type")',
        'CREATE INDEX IF NOT EXISTS ix_inv_cond ON inventory("Condition/Status")',
        "CREATE INDEX IF NOT EXISTS ix_hist_serial ON history(serial)",
        "CREATE INDEX IF NOT EXISTS ix_hist_event ON history(event)",
        "CREATE INDEX IF NOT EXISTS ix_ink_code ON ink(code)",
    ]


_COLS = {
    "inventory":    ["id"] + [f'"{c}"' for c in INV_COLS],
    "history":      ["id", "date", "asset", "serial", "from_dept", "to_dept", "event", "description", "actor"],
    "replacements": ["id", "date", "old_asset", "old_serial", "dept", "new_desc", "new_serial", "new_cond", "note", "actor"],
    "ink":          ["id", "ts", "code", "name", "action", "qty", "ict_room", "aux", "note"],
    "log":          ["id", "ts", "action", "description", "actor"],
}


class StorageError(Exception):
    pass


class Storage:
    def __init__(self, base_dir, backup_dir=None, seed_xlsx=None):
        self.base_dir   = base_dir
        self.path       = os.path.join(base_dir, DB_NAME)
        self.backup_dir = backup_dir or os.path.join(base_dir, "backups")
        self.export_dir = os.path.join(base_dir, "exports")
        self._jcache    = None
        self._snap_lock = threading.Lock()
        os.makedirs(base_dir, exist_ok=True)
        os.makedirs(self.backup_dir, exist_ok=True)
        self._remove_stale_tmp()
        self._init_schema()
        self._first_run_import(seed_xlsx)
        self._startup_check()

    # ── connections / transactions ─────────────────────────────────────────────
    def _connect(self):
        c = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        c.execute("PRAGMA busy_timeout=30000")
        c.execute("PRAGMA foreign_keys=ON")
        c.execute("PRAGMA synchronous=FULL")
        return c

    @contextlib.contextmanager
    def _tx(self, write=True, snapshot=None):
        """One ACID transaction. snapshot: None | 'auto' (rate-limited) | 'force'."""
        if write and snapshot:
            self.snapshot(force=(snapshot == "force"))
        c = self._connect()
        try:
            c.execute("BEGIN IMMEDIATE" if write else "BEGIN")
            try:
                yield c
                if write:
                    c.execute("UPDATE meta SET value = CAST(value AS INTEGER) + 1 WHERE key='rev'")
                c.execute("COMMIT")
            except BaseException:
                try: c.execute("ROLLBACK")
                except sqlite3.Error: pass
                raise
        finally:
            c.close()

    def _init_schema(self):
        c = self._connect()
        try:
            c.execute("PRAGMA journal_mode=WAL")
            ver = c.execute("PRAGMA user_version").fetchone()[0]
            if ver > SCHEMA_VERSION:
                raise StorageError("This database was created by a newer version of DericBI Inventory Control. "
                                   "Please update the application.")
            c.execute("BEGIN IMMEDIATE")
            try:
                for stmt in _schema():
                    c.execute(stmt)
                c.execute("INSERT OR IGNORE INTO meta(key,value) VALUES('rev','0')")
                c.execute(f"PRAGMA user_version={SCHEMA_VERSION}")
                c.execute("COMMIT")
            except BaseException:
                try: c.execute("ROLLBACK")
                except sqlite3.Error: pass
                raise
        finally:
            c.close()

    def _meta(self, c, key, default=None):
        r = c.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return r[0] if r else default

    def _startup_check(self):
        c = self._connect()
        try:
            res = c.execute("PRAGMA quick_check").fetchone()[0]
        finally:
            c.close()
        if res != "ok":
            newest = self.list_backups()[:1]
            hint = f" The newest backup is '{newest[0]['name']}' in the backups folder." if newest else ""
            raise StorageError("The database failed its integrity check (" + str(res)[:120] + ")." + hint +
                               " Nothing was changed. Restore a backup from Security > Backups.")

    # ── first-run import (Excel -> SQLite) ─────────────────────────────────────
    def _first_run_import(self, seed_xlsx):
        c = self._connect()
        try:
            initialised = self._meta(c, "initialised")
        finally:
            c.close()
        if initialised:
            return
        live = os.path.join(self.base_dir, LEGACY_XLSX)
        source = live if os.path.exists(live) else (seed_xlsx if seed_xlsx and os.path.exists(seed_xlsx) else None)
        summary = None
        if source:
            try:
                with self._tx() as c:
                    summary = self._import_xlsx(c, source, replace=False)
                    c.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('initialised',?)", (now(),))
                    c.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('imported_from',?)",
                              (json.dumps({"file": os.path.basename(source), "at": now(), "counts": summary["counts"]}),))
                    if not summary["pristine_seed"]:
                        dup, sus = summary["duplicates"], summary["suspect_assignee"]
                        self._log(c, "MIGRATE",
                                  f"Imported {summary['counts']['inventory']} assets, {summary['counts']['history']} history, "
                                  f"{summary['counts']['log']} log and {summary['counts']['ink']} ink records from Excel into the new "
                                  f"database." + (f" {dup} duplicate serial number(s) found - please review." if dup else "") +
                                  (f" {sus} asset(s) have 'Assigned To' equal to their department - an old bug could have overwritten the "
                                   f"assignee when a department was edited; please review." if sus else ""))
            except Exception as e:
                raise StorageError(f"Could not import '{os.path.basename(source)}' into the new database: {e}. "
                                   f"Your Excel file was NOT changed. Nothing was lost.") from e
            # Imported successfully: archive the Excel file (kept, never deleted) unless it was the pristine sample.
            if source == live:
                try:
                    if summary["pristine_seed"]:
                        os.remove(live)
                    else:
                        arch = os.path.join(self.backup_dir, "pre-sqlite"); os.makedirs(arch, exist_ok=True)
                        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        shutil.copy2(live, os.path.join(arch, f"ICT_MASTER_{stamp}.xlsx"))
                        os.replace(live, os.path.join(self.base_dir, "ICT_MASTER.pre-sqlite.xlsx"))
                except Exception as e:
                    sys.stderr.write(f"[storage] could not archive Excel file: {e}\n")
        else:
            with self._tx() as c:
                c.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('initialised',?)", (now(),))

    # ── Excel import ───────────────────────────────────────────────────────────
    def _import_xlsx(self, c, path, replace):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            names = {n.casefold(): n for n in wb.sheetnames}
            def sheet(name):
                n = names.get(name.casefold()); return wb[n] if n else None
            def rows_of(ws):
                return [r for r in ws.iter_rows(values_only=True)] if ws is not None else []

            inv_ws = sheet(SHEET_INV)
            inv_rows = []
            if inv_ws is not None:
                inv_rows = self._read_inventory_rows(rows_of(inv_ws))
            else:
                legacy = next((names[k] for k in ("work sheet", "worksheet") if k in names), None)
                if legacy is None:
                    raise ValueError("no 'Inventory' sheet found in the workbook")
                inv_rows = self._read_old_format(rows_of(wb[legacy]))

            def nonblank(r): return any(v is not None and str(v).strip() != "" for v in r)
            hist = [r for r in rows_of(sheet(SHEET_HISTORY))[1:] if nonblank(r)]
            repl = [r for r in rows_of(sheet(SHEET_REPLACE))[1:] if nonblank(r)]
            ink  = [r for r in rows_of(sheet(SHEET_INK))[1:] if nonblank(r)]
            logs = [r for r in rows_of(sheet(SHEET_LOG))[1:] if nonblank(r)]
        finally:
            wb.close()

        def cell(r, i): return clean_cell(r[i]) if i < len(r) else ""
        if replace:
            for t in TABLES: c.execute(f"DELETE FROM {t}")
            c.execute("DELETE FROM sqlite_sequence")

        inv_sql = f'INSERT INTO inventory({", ".join(_COLS["inventory"][1:])}) VALUES({",".join("?" * len(INV_COLS))})'
        c.executemany(inv_sql, inv_rows)
        c.executemany("INSERT INTO history(date,asset,serial,from_dept,to_dept,event,description,actor) VALUES(?,?,?,?,?,?,?,?)",
                      [tuple(cell(r, i) for i in range(8)) for r in hist])
        c.executemany("INSERT INTO replacements(date,old_asset,old_serial,dept,new_desc,new_serial,new_cond,note,actor) VALUES(?,?,?,?,?,?,?,?,?)",
                      [tuple(cell(r, i) for i in range(9)) for r in repl])
        c.executemany("INSERT INTO ink(ts,code,name,action,qty,ict_room,aux,note) VALUES(?,?,?,?,?,?,?,?)",
                      [(cell(r, 0), cell(r, 1), cell(r, 2), cell(r, 3), max(0, _int(r[4] if len(r) > 4 else 0)),
                        (None if (len(r) <= 5 or r[5] in (None, "", "None")) else _int_or_none(r[5])),
                        _int(r[6] if len(r) > 6 else 0), cell(r, 7)) for r in ink])
        c.executemany("INSERT INTO log(ts,action,description,actor) VALUES(?,?,?,?)",
                      [tuple(cell(r, i) for i in range(4)) for r in logs])

        counts = {"inventory": len(inv_rows), "history": len(hist), "replacements": len(repl), "ink": len(ink), "log": len(logs)}
        for t, expected in (("inventory", counts["inventory"]), ("history", counts["history"]), ("replacements", counts["replacements"]),
                            ("ink", counts["ink"]), ("log", counts["log"])):
            got = c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            if got != expected:
                raise ValueError(f"verification failed for {t}: read {expected} rows, stored {got}")
        sidx = INV_COLS.index("Serial No.")
        seen, dups = set(), 0
        for r in inv_rows:
            s = r[sidx].strip().lower()
            if s in NO_SERIAL: continue
            if s in seen: dups += 1
            seen.add(s)
        suspect = sum(1 for r in inv_rows if r[INV_COLS.index("Assigned To")] and
                      r[INV_COLS.index("Assigned To")].casefold() == r[INV_COLS.index("Department/Location")].casefold())
        pristine = (counts["inventory"] > 0 and all(r[sidx].startswith("SAMPLE-") for r in inv_rows)
                    and not (counts["history"] or counts["replacements"] or counts["ink"] or counts["log"]))
        return {"counts": counts, "duplicates": dups, "suspect_assignee": suspect, "pristine_seed": pristine}

    @staticmethod
    def _read_inventory_rows(rows):
        if not rows: return []
        header = [str(h).strip().casefold() if h is not None else "" for h in rows[0]]
        idx = {}
        for col in INV_COLS:
            try: idx[col] = header.index(col.casefold())
            except ValueError: idx[col] = None
        if idx["Equipment Type"] is None:
            raise ValueError("the Inventory sheet has no 'Equipment Type' column")
        out = []
        for r in rows[1:]:
            if not any(v is not None and str(v).strip() != "" for v in r): continue
            out.append(tuple(clean_cell(r[idx[c]]) if idx[c] is not None and idx[c] < len(r) else "" for c in INV_COLS))
        return out

    @staticmethod
    def _read_old_format(rows):
        """Very first app version: a single 'work sheet' with different column names."""
        if not rows: return []
        old_h = [str(c).strip() if c else "" for c in rows[0]]
        def g(row, name):
            try:
                i = old_h.index(name)
                return clean_cell(row[i]) if i < len(row) else ""
            except ValueError:
                return ""
        def fix_ip(ip): return (ip or "").replace("192.068.", "192.168.").replace("192.0.168.", "192.168.")
        out = []
        for row in rows[1:]:
            if not any(v is not None and str(v).strip() != "" for v in row): continue
            ram  = " ".join(p for p in [g(row, "RAM Size"), g(row, "RAM Type")] if p)
            disk = " ".join(p for p in [g(row, "Disk Size"), g(row, "Disk Type"), g(row, "Disk Health")] if p)
            out.append((g(row, "Date Collected"), g(row, "Purchase Date"), g(row, "Equipment Type"), g(row, "Brand/Model"),
                        g(row, "Serial No."), g(row, "Assigned To"), g(row, "Department/Location"),
                        g(row, "Condition/Status") or "Good", g(row, "OS/Firmware"), fix_ip(g(row, "IP Address")),
                        g(row, "MAC Address"), g(row, "CPU"), ram, disk, g(row, "Remarks"), ""))
        return out

    # ── helpers ────────────────────────────────────────────────────────────────
    def _log(self, c, action, desc, by="ICT Manager"):
        c.execute("INSERT INTO log(ts,action,description,actor) VALUES(?,?,?,?)", (now(), txt(action), txt(desc), txt(by) or "ICT Manager"))

    def _asset(self, c, row_id):
        try: rid = int(row_id)
        except (TypeError, ValueError): raise ValueError("Invalid asset id")
        if isinstance(row_id, bool) or rid < 1: raise ValueError("Invalid asset id")
        r = c.execute("SELECT id, " + ", ".join(f'"{x}"' for x in INV_COLS) + " FROM inventory WHERE id=?", (rid,)).fetchone()
        if not r: raise ValueError("This asset no longer exists - refresh and try again")
        d = {"_row": r[0]}; d.update(zip(INV_COLS, r[1:])); return d

    @staticmethod
    def _check_len(data):
        for k, v in data.items():
            if isinstance(v, str) and len(v) > MAX_FIELD_CHARS:
                raise ValueError(f"'{k}' is too long ({len(v):,} characters; limit {MAX_FIELD_CHARS:,})")

    def _check_serial(self, c, serial, exclude_id=None):
        s = txt(serial).lower()
        if s in NO_SERIAL: return
        q = ('SELECT id, "Equipment Type", "Brand/Model" FROM inventory WHERE lower(trim("Serial No."))=? '
             'AND "Condition/Status" NOT IN (?,?)')
        args = [s, *CLOSED_CONDITIONS]
        if exclude_id is not None: q += " AND id<>?"; args.append(exclude_id)
        r = c.execute(q + " LIMIT 1", args).fetchone()
        if r:
            who = f"{r[1]} {r[2]}".strip()
            raise ValueError(f"Serial number '{txt(serial)}' is already registered to asset #{r[0]} ({who}). "
                             f"Serial numbers must be unique.")

    # ── READ ───────────────────────────────────────────────────────────────────
    def load_all(self):
        return self._load(with_rev=False)

    def _load(self, with_rev):
        with self._tx(write=False) as c:
            rev = int(self._meta(c, "rev", "0"))
            sel = "SELECT id, " + ", ".join(f'"{x}"' for x in INV_COLS) + " FROM inventory ORDER BY id"
            inv = []
            for r in c.execute(sel):
                d = {"_row": r[0]}
                for i, col in enumerate(INV_COLS): d[col] = r[i + 1]
                inv.append(d)
            ink = self._ink_state(c.execute("SELECT code, action, qty, ict_room FROM ink ORDER BY id"))
            log = [{"time": r[0], "action": r[1], "desc": r[2]} for r in c.execute("SELECT ts, action, description FROM log ORDER BY id DESC")]
            hist = [{"date": r[0], "asset": r[1], "serial": r[2], "from": r[3], "to": r[4], "event": r[5], "note": r[6], "by": r[7]}
                    for r in c.execute("SELECT date, asset, serial, from_dept, to_dept, event, description, actor FROM history ORDER BY id DESC")]
        data = {"inventory": inv, "ink": ink, "log": log, "history": hist}
        return (rev, data) if with_rev else data

    @staticmethod
    def _ink_state(rows):
        state = {k: {"store": INK_DEFAULTS[k], "ictRoom": 0} for k in INK_CODES}
        for code, action, qty, snap in rows:
            code = (code or "").strip()
            if code not in state: continue
            qty = qty or 0
            if action == "restock":
                state[code]["store"] += qty
            elif action == "move_to_ict":
                state[code]["store"] = max(0, state[code]["store"] - qty)
                state[code]["ictRoom"] += qty
            elif action == "use":
                from_ict = min(qty, state[code]["ictRoom"])
                state[code]["ictRoom"] = max(0, state[code]["ictRoom"] - from_ict)
                state[code]["store"] = max(0, state[code]["store"] - (qty - from_ict))
            if snap is not None and action in ("move_to_ict", "use", "set_printer"):
                state[code]["ictRoom"] = snap
        return state

    def load_all_json(self):
        """JSON for the UI. Cached until the database changes (by this app, another
        window/process, or a restore) - detected through the transaction counter."""
        with self._tx(write=False) as c:
            rev = int(self._meta(c, "rev", "0"))
        cached = self._jcache
        if cached and cached[0] == rev:
            return cached[1]
        rev, data = self._load(with_rev=True)
        text = json.dumps(data)
        self._jcache = (rev, text)
        return text

    def get_transfer_log(self):
        with self._tx(write=False) as c:
            return [{"date": r[0], "asset": r[1], "serial": r[2], "from": r[3], "to": r[4], "note": r[5], "by": r[6]}
                    for r in c.execute("SELECT date, asset, serial, from_dept, to_dept, description, actor FROM history "
                                       "WHERE event='transferred' ORDER BY id DESC")]

    # ── WRITE ──────────────────────────────────────────────────────────────────
    def log_session(self, action, desc, by="ICT Manager"):
        with self._tx() as c:
            self._log(c, action, desc, by)

    def add_asset(self, data):
        if not isinstance(data, dict): raise ValueError("Invalid data")
        self._check_len(data)
        vals = {col: txt(data.get(col, "")) for col in INV_COLS}
        if not vals["Equipment Type"]: raise ValueError("Equipment Type is required")
        if not vals["Date Collected"]: vals["Date Collected"] = today()
        with self._tx(snapshot="auto") as c:
            self._check_serial(c, vals["Serial No."])
            cur = c.execute(f'INSERT INTO inventory({", ".join(_COLS["inventory"][1:])}) VALUES({",".join("?" * len(INV_COLS))})',
                            [vals[col] for col in INV_COLS])
            self._log(c, "ADD", f"Added: {vals['Equipment Type']} — {vals['Brand/Model']} | {vals['Department/Location']}")
            return cur.lastrowid

    def update_asset(self, row_id, data):
        if not isinstance(data, dict): raise ValueError("Invalid data")
        self._check_len(data)
        changes = {col: txt(data[col]) for col in INV_COLS if col in data}
        with self._tx(snapshot="auto") as c:
            cur = self._asset(c, row_id)
            if "Equipment Type" in changes and not changes["Equipment Type"]:
                raise ValueError("Equipment Type is required")
            if "Serial No." in changes and changes["Serial No."].lower() != cur["Serial No."].strip().lower():
                self._check_serial(c, changes["Serial No."], exclude_id=cur["_row"])
            if changes:
                c.execute("UPDATE inventory SET " + ", ".join(f'"{k}"=?' for k in changes) + " WHERE id=?", [*changes.values(), cur["_row"]])
            merged = {**cur, **changes}
            self._log(c, "EDIT", f"Edited asset #{cur['_row']}: {merged['Equipment Type']} — {merged['Brand/Model']}")

    def delete_asset(self, row_id, desc):
        with self._tx(snapshot="force") as c:
            cur = self._asset(c, row_id)
            c.execute("DELETE FROM inventory WHERE id=?", (cur["_row"],))
            label = txt(desc) or f"{cur['Equipment Type']} {cur['Brand/Model']} (S/N:{cur['Serial No.']})"
            self._log(c, "DELETE", f"Deleted: {label}")

    def remove_sample_data(self):
        """Remove the demo assets a fresh install ships with (Serial No. starts with 'SAMPLE-') and their
        transfer history. Everything else is left alone, the audit log is never edited (a new entry records
        the removal) and a safety copy of the database is taken first. Returns how many assets were removed."""
        with self._tx(write=False) as c:
            if not c.execute('SELECT 1 FROM inventory WHERE substr("Serial No.", 1, 7) = ? LIMIT 1', ("SAMPLE-",)).fetchone():
                return 0
        with self._tx(snapshot="force") as c:
            n = c.execute('DELETE FROM inventory WHERE substr("Serial No.", 1, 7) = ?', ("SAMPLE-",)).rowcount
            c.execute("DELETE FROM history WHERE substr(serial, 1, 7) = ?", ("SAMPLE-",))
            if n: self._log(c, "DELETE", f"Removed {n} sample asset{'s' if n != 1 else ''}")
        return n

    def transfer_asset(self, row_id, asset_desc, serial, from_dept, to_dept, note, date, by="ICT Manager"):
        to_dept = txt(to_dept)
        if not to_dept: raise ValueError("Destination department is required")
        note, date = txt(note), txt(date) or today()
        with self._tx(snapshot="auto") as c:
            cur = self._asset(c, row_id)
            real_from = cur["Department/Location"]
            if real_from.casefold() == to_dept.casefold():
                raise ValueError(f"This asset is already in {to_dept}")
            entry = f"[{date}] {real_from} → {to_dept}" + (f": {note}" if note else "")
            history = (cur["Transfer History"] + " | " + entry).strip(" |")
            c.execute('UPDATE inventory SET "Department/Location"=?, "Transfer History"=? WHERE id=?', (to_dept, history, cur["_row"]))
            ser = txt(serial) or cur["Serial No."]
            desc = txt(asset_desc) or f"{cur['Equipment Type']} {cur['Brand/Model']}".strip()
            c.execute("INSERT INTO history(date,asset,serial,from_dept,to_dept,event,description,actor) VALUES(?,?,?,?,?,?,?,?)",
                      (date, desc, ser, real_from, to_dept, "transferred", f"Transferred from {real_from} to {to_dept}. Note: {note}", txt(by) or "ICT Manager"))
            self._log(c, "TRANSFER", f"Transferred {desc} (S/N:{ser}) from {real_from} → {to_dept}")

    def record_replacement(self, old_row, old_desc, rep):
        if not isinstance(rep, dict): raise ValueError("Invalid data")
        self._check_len(rep)
        date, note = txt(rep.get("date")) or today(), txt(rep.get("note"))
        with self._tx(snapshot="force") as c:
            old = self._asset(c, old_row)
            new_type = txt(rep.get("newType")) or old["Equipment Type"]
            new_serial = txt(rep.get("newSerial"))
            c.execute('UPDATE inventory SET "Condition/Status"=?, "Remarks"=? WHERE id=?',
                      ("Replaced", (old["Remarks"] + f" | REPLACED {date}: {note}").strip(" |"), old["_row"]))
            self._check_serial(c, new_serial)                      # old asset is now 'Replaced', so it may share a serial
            dept = txt(rep.get("dept")) or old["Department/Location"]
            cost = txt(rep.get("cost"))
            remarks = f"Replaced: {txt(old_desc)} on {date}. {note}".strip() + (f" Cost: KES {cost}." if cost not in ("", "0") else "")
            new_vals = {"Date Collected": date, "Purchase Date": date, "Equipment Type": new_type,
                        "Brand/Model": txt(rep.get("newBrand")), "Serial No.": new_serial, "Assigned To": txt(rep.get("assigned")),
                        "Department/Location": dept, "Condition/Status": txt(rep.get("newCond")) or "New",
                        "IP Address": txt(rep.get("ip")), "Remarks": remarks}
            vals = [new_vals.get(col, "") for col in INV_COLS]
            c.execute(f'INSERT INTO inventory({", ".join(_COLS["inventory"][1:])}) VALUES({",".join("?" * len(INV_COLS))})', vals)
            new_desc = f"{new_type} {new_vals['Brand/Model']}".strip()
            c.execute("INSERT INTO replacements(date,old_asset,old_serial,dept,new_desc,new_serial,new_cond,note,actor) VALUES(?,?,?,?,?,?,?,?,?)",
                      (date, txt(old_desc), txt(rep.get("oldSerial")) or old["Serial No."], dept, new_desc, new_serial, new_vals["Condition/Status"], note, "ICT Manager"))
            c.execute("INSERT INTO history(date,asset,serial,from_dept,to_dept,event,description,actor) VALUES(?,?,?,?,?,?,?,?)",
                      (date, txt(old_desc), txt(rep.get("oldSerial")) or old["Serial No."], dept, dept, "replaced",
                       f"Replaced with: {new_desc} S/N:{new_serial}. Note: {note}", "ICT Manager"))
            self._log(c, "REPLACE", f"Replaced {txt(old_desc)} → {new_vals['Brand/Model']} on {date}")

    def write_history_event(self, asset_desc, serial, dept, event, note, date, cost="", by="ICT Manager"):
        event, serial, note = txt(event), txt(serial), txt(note)
        if not event: raise ValueError("Event type is required")
        cost = txt(cost)
        desc = note + (f" (Cost: KES {cost})" if cost not in ("", "0", "0.0") else "")
        with self._tx(snapshot="auto") as c:
            c.execute("INSERT INTO history(date,asset,serial,from_dept,to_dept,event,description,actor) VALUES(?,?,?,?,?,?,?,?)",
                      (txt(date) or today(), txt(asset_desc), serial, txt(dept), txt(dept), event, desc, txt(by) or "ICT Manager"))
            if serial and event in ("repaired", "disposed", "replaced", "upgraded", "reassigned"):
                r = c.execute('SELECT id, "Condition/Status" FROM inventory WHERE trim("Serial No.")=? '
                              'ORDER BY ("Condition/Status" IN (?,?)), id LIMIT 1', (serial, *CLOSED_CONDITIONS)).fetchone()
                if r:
                    rid, cond = r
                    if event == "repaired" and ("Repair" in cond or cond in ("Fair (Needs Repair)", "Needs Repair")):
                        c.execute('UPDATE inventory SET "Condition/Status"=? WHERE id=?', ("Good", rid))
                    elif event == "disposed":
                        c.execute('UPDATE inventory SET "Condition/Status"=? WHERE id=?', ("Disposed", rid))
                    elif event == "replaced":
                        c.execute('UPDATE inventory SET "Condition/Status"=? WHERE id=?', ("Replaced", rid))
                    elif event == "upgraded":
                        c.execute('UPDATE inventory SET "Condition/Status"=? WHERE id=?', ("Good", rid))
                    elif event == "reassigned":
                        who = self._new_assignee(note)
                        if who is not None:
                            c.execute('UPDATE inventory SET "Assigned To"=? WHERE id=?', (who, rid))
            self._log(c, event.upper(), f"{event} on {txt(asset_desc)}: {note[:80]}")

    @staticmethod
    def _new_assignee(note):
        """Which 'Assigned To' value (if any) does a 'reassigned' history note ask for?
        The old app wrote the text after ANY arrow into Assigned To - so a note like
        'Department changed: Lab → Pharmacy' wrongly replaced the person with the department."""
        note = txt(note)
        low = note.casefold()
        if low.startswith("department changed") or low.startswith("status changed"):
            return None                                           # not an assignee change
        if low.startswith("assignment changed"):
            tail = note.rsplit("→", 1)[-1].strip()
            return "" if tail.casefold() in ("(none)", "none", "") else tail
        ms = list(re.finditer(r"(?:\bto\b|→)\s*(.+?)(?:\s*\[|$)", note, re.IGNORECASE))   # free text: 'Reassigned to John'
        return txt(ms[-1].group(1)) if ms else None

    def ink_apply(self, code, action, qty, ict_room_after=0, note=""):
        """Validate against current stock AND record, in ONE transaction (no check-then-act race)."""
        code, action = txt(code), txt(action)
        if code not in INK_CODES: raise ValueError("Unknown ink colour")
        if action not in INK_ACTIONS: raise ValueError("Unknown ink action")
        qty, snap = _int(qty), _int(ict_room_after)
        if qty < 0 or snap < 0: raise ValueError("Quantity cannot be negative")
        with self._tx() as c:
            state = self._ink_state(c.execute("SELECT code, action, qty, ict_room FROM ink ORDER BY id")).get(code, {})
            if action == "use":
                avail = state.get("ictRoom", 0) + state.get("store", 0)
                if avail < qty: raise ValueError(f"Only {avail} available (ICT Room + Store)")
            elif action == "move_to_ict":
                if state.get("store", 0) < qty: raise ValueError(f"Only {state.get('store', 0)} in store")
            c.execute("INSERT INTO ink(ts,code,name,action,qty,ict_room,aux,note) VALUES(?,?,?,?,?,?,0,?)",
                      (now(), code, INK_NAMES.get(code, code), action, qty, snap, txt(note)))
            nm = f"{qty}x {INK_NAMES.get(code, code)} ({code})"
            self._log(c, "INK", {"restock": f"Ink restocked to store: {nm}", "move_to_ict": f"Ink moved Store→ICT Room: {nm}",
                                 "use": f"Ink used: {nm}"}.get(action, f"Ink {action}: {nm}"))

    # ── backups ────────────────────────────────────────────────────────────────
    def _backup_to(self, dst):
        src = self._connect()
        out = sqlite3.connect(dst)
        try:
            src.backup(out)
            out.execute("PRAGMA journal_mode=DELETE")      # backups are single self-contained files
            ok = out.execute("PRAGMA quick_check").fetchone()[0]
            if ok != "ok": raise StorageError("backup verification failed: " + str(ok))
        finally:
            out.close(); src.close()

    def _remove_stale_tmp(self):
        """Leftovers of a snapshot/export that was interrupted by a crash."""
        for d in (self.backup_dir, os.path.join(self.backup_dir, "pre-restore"), self.export_dir):
            try:
                for f in os.listdir(d):
                    fp = os.path.join(d, f)
                    if f.endswith(".tmp") and time.time() - os.path.getmtime(fp) > 600: os.remove(fp)
            except OSError:
                pass

    def snapshot(self, force=False, min_interval=300, subdir=None):
        """Point-in-time copy of the database (SQLite online-backup API: consistent even while in use).
        Serialised: concurrent callers never write the same file. An automatic snapshot is simply
        skipped if another one is already running; a forced one (delete/replace/restore) waits."""
        if not self._snap_lock.acquire(blocking=force):
            return None
        try:
            d = self.backup_dir if subdir is None else os.path.join(self.backup_dir, subdir)
            os.makedirs(d, exist_ok=True)
            if not force and subdir is None:
                newest = max((os.path.getmtime(os.path.join(d, f)) for f in os.listdir(d)
                              if f.startswith("ICT_backup_") and f.endswith(".db")), default=0)
                if newest and time.time() - newest < min_interval:
                    return None
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            dst, n = os.path.join(d, f"ICT_backup_{stamp}.db"), 1
            while os.path.exists(dst):
                dst = os.path.join(d, f"ICT_backup_{stamp}_{n}.db"); n += 1
            tmp = f"{dst}.{os.getpid()}.{uuid.uuid4().hex[:8]}.tmp"
            try:
                self._backup_to(tmp)
                if os.path.exists(dst): dst = f"{dst[:-3]}_{uuid.uuid4().hex[:6]}.db"       # another PROCESS used this name
                os.replace(tmp, dst)
            finally:
                if os.path.exists(tmp):
                    try: os.remove(tmp)
                    except OSError: pass
            self._prune(d, "ICT_backup_", ".db", 10)
            if subdir is None:                               # one long-term copy per day, 30 days
                daily_dir = os.path.join(self.backup_dir, "daily"); os.makedirs(daily_dir, exist_ok=True)
                daily = os.path.join(daily_dir, f"ICT_daily_{datetime.date.today():%Y%m%d}.db")
                if not os.path.exists(daily): shutil.copy2(dst, daily)
                self._prune(daily_dir, "ICT_daily_", ".db", 30)
            return dst
        except Exception as e:
            sys.stderr.write(f"[storage] snapshot skipped: {e}\n")
            return None
        finally:
            self._snap_lock.release()

    @staticmethod
    def _prune(d, prefix, ext, keep):
        files = sorted(f for f in os.listdir(d) if f.startswith(prefix) and f.endswith(ext))
        for old in files[:-keep]:
            try: os.remove(os.path.join(d, old))
            except OSError: pass

    def list_backups(self):
        out = []
        if os.path.isdir(self.backup_dir):
            for f in sorted(os.listdir(self.backup_dir), reverse=True):
                fp = os.path.join(self.backup_dir, f)
                if os.path.isfile(fp) and f.endswith((".db", ".xlsx")):
                    out.append({"name": f, "size_kb": round(os.path.getsize(fp) / 1024, 1),
                                "modified": datetime.datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M:%S")})
        return out

    def snapshot_to(self, dst):
        """Consistent copy of the live database to an arbitrary path (used by full-backup export)."""
        self._backup_to(dst)

    # ── restore ────────────────────────────────────────────────────────────────
    def _validate_db_file(self, path):
        try:
            v = sqlite3.connect(pathlib.Path(path).resolve().as_uri() + "?mode=ro", uri=True)
        except sqlite3.Error as e:
            raise ValueError("That file is not a valid database.") from e
        try:
            try:
                if v.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                    raise ValueError("The database inside this backup is damaged - nothing was changed.")
                tables = {r[0] for r in v.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                missing = [t for t in TABLES if t not in tables]
                if missing: raise ValueError("Backup database is missing tables: " + ", ".join(missing) + " - nothing was changed.")
                ver = v.execute("PRAGMA user_version").fetchone()[0]
                if ver > SCHEMA_VERSION: raise ValueError("This backup was made by a newer version of the app - please update first.")
                for t in TABLES:
                    cols = [r[1] for r in v.execute(f"PRAGMA table_info({t})")]
                    want = [c.strip('"') for c in _COLS[t]]
                    if cols != want: raise ValueError(f"Backup table '{t}' has an unexpected layout - nothing was changed.")
            except sqlite3.DatabaseError as e:
                raise ValueError("The database inside this backup is damaged - nothing was changed.") from e
        finally:
            v.close()

    def restore_from_db(self, path):
        """Replace ALL data with the contents of a backup database, atomically."""
        self._validate_db_file(path)
        self.snapshot(force=True, subdir="pre-restore")
        c = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        try:
            c.execute("PRAGMA busy_timeout=30000"); c.execute("PRAGMA synchronous=FULL")
            c.execute("ATTACH DATABASE ? AS bk", (path,))
            c.execute("BEGIN IMMEDIATE")
            try:
                for t in TABLES: c.execute(f"DELETE FROM main.{t}")
                c.execute("DELETE FROM main.sqlite_sequence")
                for t in TABLES:
                    cols = ", ".join(_COLS[t])
                    c.execute(f"INSERT INTO main.{t}({cols}) SELECT {cols} FROM bk.{t}")
                self._log(c, "RESTORE", "Database restored from a backup")
                c.execute("UPDATE meta SET value = CAST(value AS INTEGER) + 1 WHERE key='rev'")
                c.execute("COMMIT")
            except BaseException:
                try: c.execute("ROLLBACK")
                except sqlite3.Error: pass
                raise
            finally:
                try: c.execute("DETACH DATABASE bk")
                except sqlite3.Error: pass
        finally:
            c.close()
        self._jcache = None

    def restore_from_xlsx(self, path):
        """Restore from a legacy/exported Excel workbook (all-or-nothing)."""
        self.snapshot(force=True, subdir="pre-restore")
        with self._tx() as c:
            s = self._import_xlsx(c, path, replace=True)
            self._log(c, "RESTORE", f"Database restored from an Excel backup ({s['counts']['inventory']} assets)")
        self._jcache = None

    # ── export ─────────────────────────────────────────────────────────────────
    def export_xlsx(self, dst):
        """Write a complete Excel copy of the data (never the live database)."""
        from openpyxl import Workbook
        from openpyxl.cell import WriteOnlyCell
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
        wb = Workbook(write_only=True)
        hfont, hfill, hal = Font(bold=True, color="FFFFFF", size=11), PatternFill("solid", fgColor="1F2937"), Alignment(horizontal="center", vertical="center")
        def header(ws, names):
            row = []
            for n in names:
                cl = WriteOnlyCell(ws, value=n); cl.font, cl.fill, cl.alignment = hfont, hfill, hal; row.append(cl)
            ws.append(row)
        def safe(ws, v):
            if isinstance(v, str):
                v = ILLEGAL_CHARACTERS_RE.sub("", v)
                cl = WriteOnlyCell(ws, value=v)
                if v[:1] == "=": cl.data_type = "s"            # text, never a live formula
                return cl
            return v
        widths = {"Date Collected": 13, "Purchase Date": 13, "Equipment Type": 22, "Brand/Model": 20, "Serial No.": 22, "Assigned To": 20,
                  "Department/Location": 24, "Condition/Status": 22, "OS/Firmware": 16, "IP Address": 16, "MAC Address": 18,
                  "CPU": 14, "RAM": 12, "Disk": 20, "Remarks": 32, "Transfer History": 40}
        sheets = {}
        for name in ALL_SHEETS:
            ws = wb.create_sheet(name); sheets[name] = ws
        for i, col in enumerate(INV_COLS, 1):
            sheets[SHEET_INV].column_dimensions[get_column_letter(i)].width = widths.get(col, 16)
        header(sheets[SHEET_INV], INV_COLS); header(sheets[SHEET_HISTORY], HIST_HEADERS); header(sheets[SHEET_REPLACE], REPL_HEADERS)
        header(sheets[SHEET_INK], INK_HEADERS); header(sheets[SHEET_LOG], LOG_HEADERS); header(sheets[SHEET_CONFIG], ["Key", "Value"])
        with self._tx(write=False) as c:
            for r in c.execute("SELECT " + ", ".join(f'"{x}"' for x in INV_COLS) + " FROM inventory ORDER BY id"):
                sheets[SHEET_INV].append([safe(sheets[SHEET_INV], v) for v in r])
            for r in c.execute("SELECT date,asset,serial,from_dept,to_dept,event,description,actor FROM history ORDER BY id"):
                sheets[SHEET_HISTORY].append([safe(sheets[SHEET_HISTORY], v) for v in r])
            for r in c.execute("SELECT date,old_asset,old_serial,dept,new_desc,new_serial,new_cond,note,actor FROM replacements ORDER BY id"):
                sheets[SHEET_REPLACE].append([safe(sheets[SHEET_REPLACE], v) for v in r])
            for r in c.execute("SELECT ts,code,name,action,qty,ict_room,aux,note FROM ink ORDER BY id"):
                sheets[SHEET_INK].append([safe(sheets[SHEET_INK], v) for v in r])
            for r in c.execute("SELECT ts,action,description,actor FROM log ORDER BY id"):
                sheets[SHEET_LOG].append([safe(sheets[SHEET_LOG], v) for v in r])
        os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
        tmp = f"{dst}.{os.getpid()}.{uuid.uuid4().hex[:8]}.tmp"
        try:
            wb.save(tmp)
            os.replace(tmp, dst)
        finally:
            if os.path.exists(tmp):
                try: os.remove(tmp)
                except OSError: pass
        return dst

    def export_for_user(self):
        """Excel export into <data>/exports/ (keeps the newest 5) - used by 'Export to Excel'."""
        os.makedirs(self.export_dir, exist_ok=True)
        dst = os.path.join(self.export_dir, f"DericBI_Inventory_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx")
        self.export_xlsx(dst)
        self._prune(self.export_dir, "DericBI_Inventory_", ".xlsx", 5)
        return dst
