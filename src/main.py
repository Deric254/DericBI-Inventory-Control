"""
DericBI Inventory Control v3
Local app: HTML UI + Python backend + SQLite (ACID) storage.
Excel is supported for one-time import of older data and for "Export to Excel".
"""
import os as _os
_os.environ.setdefault("PYWEBVIEW_GUI", "edgechromium")
_os.environ.setdefault("WEBKIT_DISABLE_COMPOSITING_MODE", "1")
_os.environ.setdefault("PYWEBVIEW_NO_UIAUTOMATION", "1")

import json, os, sys, datetime, shutil, threading, re, hashlib, secrets, hmac, time
import http.server, socketserver, webbrowser, urllib.parse

# ── Portable EXE / Electron path resolution ───────────────────────────────────
# Three launch modes:
#   1. Electron:      env vars ICT_BUNDLE_DIR + ICT_DATA_DIR are set by main.js
#   2. PyInstaller:   sys._MEIPASS == bundle dir, sys.executable dir == data dir
#   3. Script (dev):  both dirs == directory of main.py

def _meipass():
    """Return the directory for read-only bundled assets (ui.html, seed Excel)."""
    if os.environ.get('ICT_BUNDLE_DIR'):
        return os.environ['ICT_BUNDLE_DIR']
    return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

def _datadir():
    """Return the directory for mutable data files (live Excel, auth, reports)."""
    if os.environ.get('ICT_DATA_DIR'):
        return os.environ['ICT_DATA_DIR']
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BUNDLE_DIR = _meipass()   # read-only: ui.html, seed ICT_MASTER.xlsx
BASE_DIR   = _datadir()   # read-write: live data files

os.makedirs(BASE_DIR, exist_ok=True)

# On first run from a fresh install, seed the data files from the bundle
def _seed_file(name):
    dst = os.path.join(BASE_DIR, name)
    if not os.path.exists(dst):
        src = os.path.join(BUNDLE_DIR, name)
        if os.path.exists(src):
            shutil.copy2(src, dst)

for _f in ("auth.json", "logo.png"):
    _seed_file(_f)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from storage import Storage

import licence_gate
from licensing.licensing import LicenceError
licence_gate.set_data_dir(BASE_DIR)

BACKUP_DIR = os.path.join(BASE_DIR, "backups")
AUTH_FILE  = os.path.join(BASE_DIR, "auth.json")
CONFIG_FILE= os.path.join(BASE_DIR, "config.json")

# Default lists — users can extend these
DEFAULT_DEPTS = [
    "ICT Department","Administration","Outpatient","Inpatient","Pharmacy",
    "Laboratory","Radiology","Theatre","Maternity","Accounts","Records","Store"
]
DEFAULT_TYPES = [
    "Desktop Computer","Laptop","Printer","Scanner","Server","Switch","Router",
    "UPS","Monitor","Keyboard","Mouse","External HDD","Flash Drive","Projector",
    "Tablet","Phone","Photocopier","Camera","Other"
]
DEFAULT_CONDITIONS = [
    "New","Good","Fair","Fair (Needs Repair)","Needs Repair","Replaced","Disposed"
]

def today(): return datetime.date.today().strftime("%Y-%m-%d")
def now():   return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ── Integrity helpers ─────────────────────────────────────────────────────────
MAX_CELL_CHARS = 32000

def _atomic_write_bytes(path, data):
    """Crash-safe small-file write: temp file, flush to disk, atomic rename."""
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data); f.flush(); os.fsync(f.fileno())
    os.replace(tmp, path)

def _atomic_write_json(path, obj, **kw):
    _atomic_write_bytes(path, json.dumps(obj, **kw).encode("utf-8"))

def _check_fields(data):
    if not isinstance(data, dict):
        raise ValueError("Invalid data")
    for k, v in data.items():
        if isinstance(v, str) and len(v) > MAX_CELL_CHARS:
            raise ValueError(f"'{k}' is too long ({len(v):,} characters; limit {MAX_CELL_CHARS:,})")

# ════════════════════════════════════════════
# CONFIG ENGINE
# ════════════════════════════════════════════
# ── Credential hashing ────────────────────────────────────────────────────────
_PBKDF2_ITERS = 600_000

def _kdf(secret, salt_hex, iters=None, legacy=False):
    """PBKDF2-HMAC-SHA256 (current). legacy=True verifies old single-pass SHA-256 hashes."""
    if legacy:
        return hashlib.sha256((salt_hex + secret).encode()).hexdigest()
    return hashlib.pbkdf2_hmac("sha256", secret.encode(), bytes.fromhex(salt_hex), int(iters or _PBKDF2_ITERS)).hex()

class ConfigEngine:
    def __init__(self):
        self._lock = threading.Lock()
        self._data = self._load()

    def _load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE) as f:
                    return json.load(f)
            except: pass
        return {"departments": DEFAULT_DEPTS[:], "types": DEFAULT_TYPES[:], "conditions": DEFAULT_CONDITIONS[:]}

    def _save(self):
        _atomic_write_json(CONFIG_FILE, self._data, indent=2)

    def get(self):
        with self._lock:
            return dict(self._data)

    def add_item(self, key, value):
        with self._lock:
            v = value.strip()
            if v and v not in self._data.get(key, []):
                self._data.setdefault(key, []).append(v)
                self._save()
                return True
            return False

    def remove_item(self, key, value):
        with self._lock:
            lst = self._data.get(key, [])
            if value in lst:
                lst.remove(value)
                self._save()
                return True
            return False

_USERNAME_MSG = "Username must be 3-32 characters: letters, numbers, spaces and . _ - @"

def _clean_username(u):
    """Validate an account name. Returns it trimmed, or raises ValueError."""
    u = u.strip() if isinstance(u, str) else ""
    if not re.fullmatch(r"[\w .@-]{3,32}", u) or not u.strip(" ._-@"):
        raise ValueError(_USERNAME_MSG)
    return u

# ════════════════════════════════════════════
# JS API BRIDGE
# ════════════════════════════════════════════
class Api:
    def __init__(self, engine, cfg: ConfigEngine):
        self.engine = engine
        self.cfg    = cfg
        self.window = None
        self._authed      = False     # True after a correct PIN this session
        self._user        = None      # account name of the signed-in user (None for PIN-only installs)
        self._recovery_ok = False     # True after a correct recovery key (allows PIN reset)
        self._fails       = 0
        self._locked_until = 0.0

    # Methods callable before login. Everything else needs an authenticated session.
    PUBLIC_METHODS = frozenset({"is_pin_set", "verify_pin", "verify_recovery_key",
                                "set_pin", "get_logo_path", "lock_session",
                                "licence_state", "activate_licence", "claim_owner"})

    def needs_login(self):
        return os.path.exists(AUTH_FILE) and not self._authed

    def lock_session(self):
        self._authed = False
        self._recovery_ok = False
        self._user = None
        return json.dumps({"ok": True})

    def _throttle(self):
        wait = self._locked_until - time.time()
        if wait > 0:
            return f"Too many attempts. Try again in {int(wait) + 1} seconds."
        return None

    def _register_fail(self):
        self._fails += 1
        if self._fails >= 5:
            self._locked_until = time.time() + min(900, 30 * 2 ** (self._fails - 5))

    def _register_ok(self):
        self._fails = 0; self._locked_until = 0.0

    def load_data(self):
        try:
            return '{"ok": true, "data": ' + self.engine.load_all_json() + '}'
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def get_config(self):
        try:
            return json.dumps({"ok": True, "config": self.cfg.get()})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def add_config_item(self, key, value):
        try:
            ok = self.cfg.add_item(key, value)
            return json.dumps({"ok": ok, "msg": "Added" if ok else "Already exists"})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def remove_config_item(self, key, value):
        try:
            ok = self.cfg.remove_item(key, value)
            return json.dumps({"ok": ok})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def add_asset(self, data_json):
        try:
            data = json.loads(data_json)
            _check_fields(data)
            if not data.get('Equipment Type'):
                return json.dumps({"ok": False, "error": "Equipment Type is required"})
            if not data.get('Date Collected'):
                data['Date Collected'] = today()
            return json.dumps({"ok": True, "row": self.engine.add_asset(data)})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def update_asset(self, row_num, data_json):
        try:
            _d = json.loads(data_json); _check_fields(_d)
            self.engine.update_asset(int(row_num), _d)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def delete_asset(self, row_num, desc):
        try:
            self.engine.delete_asset(int(row_num), desc)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def remove_sample_data(self):
        try:
            return json.dumps({"ok": True, "removed": self.engine.remove_sample_data()})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def transfer_asset(self, row_num, asset_desc, serial, from_dept, to_dept, note, date, by='ICT Manager'):
        try:
            if not to_dept:
                return json.dumps({"ok": False, "error": "Target department is required"})
            if from_dept == to_dept:
                return json.dumps({"ok": False, "error": "Source and target departments must be different"})
            self.engine.transfer_asset(int(row_num), asset_desc, serial, from_dept, to_dept, note, date or today(), by)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def get_transfer_log(self):
        try:
            return json.dumps({"ok": True, "data": self.engine.get_transfer_log()})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def record_replacement(self, rep_json):
        try:
            rep = json.loads(rep_json)
            if not rep.get('note'):
                return json.dumps({"ok": False, "error": "Replacement note is required"})
            self.engine.record_replacement(int(rep['oldRow']), rep['oldDesc'], rep)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def write_history(self, asset_desc, serial, dept, event, note, date, cost='', by='ICT Manager'):
        try:
            self.engine.write_history_event(asset_desc, serial, dept, event, note, date, cost, by)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def ink_action(self, code, action, qty, ict_room_after=0, note=''):
        try:
            self.engine.ink_apply(code, action, qty, ict_room_after, note)   # validate + record in ONE transaction
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def log_action(self, action, desc, by="ICT Manager"):
        """Client-side initiated audit log entry."""
        try:
            self.engine.log_session(action, desc, by)
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def generate_pdf(self, dept_filter='', type_filter='', date_from='', date_to=''):
        """Generate PDF report, return base64 content for browser download."""
        try:
            import os, base64
            out_dir = os.path.join(BASE_DIR, 'reports')
            os.makedirs(out_dir, exist_ok=True)
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            fname = f'ICT_Report_{ts}.pdf'
            out_path = os.path.join(out_dir, fname)
            logo_path = ''
            for lname in ['logo.png','logo.jpg','logo.jpeg','LOGO.PNG','LOGO.JPG']:
                candidate = os.path.join(BASE_DIR, lname)
                if os.path.exists(candidate):
                    logo_path = candidate
                    break
            result = generate_pdf_report(self.engine, out_path, dept_filter, type_filter, logo_path, date_from, date_to)
            if result['ok']:
                result['filename'] = fname
                with open(out_path, 'rb') as f:
                    result['b64'] = base64.b64encode(f.read()).decode('ascii')
            return json.dumps(result)
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def get_logo_path(self):
        """Return a URL to the logo — HTTP in Electron/browser mode, file:/// in desktop mode."""
        for lname in ['logo.png','logo.jpg','logo.jpeg','LOGO.PNG','LOGO.JPG']:
            p = os.path.join(BASE_DIR, lname)
            if os.path.exists(p):
                # In Electron/browser mode the HTTP server is running — serve via /logo
                port = os.environ.get('ICT_HTTP_PORT')
                if port:
                    return f'http://127.0.0.1:{port}/logo'
                # Desktop (pywebview) mode — file:// works fine
                uri = p.replace(os.sep, '/'); return 'file:///' + uri
        return ''

    def get_backups(self):
        try:
            return json.dumps({"ok": True, "backups": self.engine.list_backups()})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def export_full_backup(self):
        """Package ALL system data into one zip for full portability."""
        try:
            import zipfile, base64, io, tempfile
            ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            buf = io.BytesIO()
            with tempfile.TemporaryDirectory() as td, zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                dbp = os.path.join(td, "inventory.db");   self.engine.snapshot_to(dbp)       # consistent copy, even while in use
                xlp = os.path.join(td, "ICT_MASTER.xlsx"); self.engine.export_xlsx(xlp)     # readable copy / downgrade path
                zf.write(dbp, "inventory.db")
                zf.write(xlp, "ICT_MASTER.xlsx")
                for fname, fpath in [("auth.json", AUTH_FILE), ("config.json", CONFIG_FILE)]:
                    if os.path.exists(fpath):
                        zf.write(fpath, fname)
                if os.path.exists(BACKUP_DIR):
                    for f in os.listdir(BACKUP_DIR):
                        if f.endswith(('.db', '.xlsx')) and os.path.isfile(os.path.join(BACKUP_DIR, f)):
                            zf.write(os.path.join(BACKUP_DIR, f), f"backups/{f}")
                manifest = {"version": 2, "format": "sqlite", "exported": ts,
                            "app": "DericBI Inventory Control", "tagline": "Cut Through Noise"}
                zf.writestr("ICT_MANIFEST.json", json.dumps(manifest, indent=2))
            b64 = base64.b64encode(buf.getvalue()).decode('ascii')
            return json.dumps({"ok": True, "b64": b64, "filename": f"ICT_FullBackup_{ts}.zip"})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def restore_full_backup(self, b64_zip):
        """Restore all system data from a full backup zip (all-or-nothing)."""
        try:
            import zipfile, base64, io, tempfile
            raw = base64.b64decode(b64_zip)
            buf = io.BytesIO(raw)
            if not zipfile.is_zipfile(buf):
                return json.dumps({"ok": False, "error": "Invalid file - not a zip archive"})
            buf.seek(0)
            with zipfile.ZipFile(buf, 'r') as zf:
                names = zf.namelist()
                if "ICT_MANIFEST.json" not in names:
                    return json.dumps({"ok": False, "error": "This is not a DericBI Inventory Control backup."})
                try:
                    manifest = json.loads(zf.read("ICT_MANIFEST.json").decode("utf-8"))
                except Exception:
                    return json.dumps({"ok": False, "error": "The backup manifest is damaged - nothing was changed."})
                if manifest.get("app") not in ("DericBI Inventory Control", "ICT Command Centre"):
                    return json.dumps({"ok": False, "error": "This backup is not from DericBI Inventory Control."})
                if sum(i.file_size for i in zf.infolist()) > 8 * 1024**3:
                    return json.dumps({"ok": False, "error": "Backup archive is too large to restore safely."})
                kind = "db" if "inventory.db" in names else ("xlsx" if "ICT_MASTER.xlsx" in names else None)
                if kind is None:
                    return json.dumps({"ok": False, "error": "Backup contains no database - cannot restore."})
                with tempfile.TemporaryDirectory() as td:
                    p = os.path.join(td, "restore." + kind)
                    with open(p, "wb") as f: f.write(zf.read("inventory.db" if kind == "db" else "ICT_MASTER.xlsx"))
                    try:
                        (self.engine.restore_from_db if kind == "db" else self.engine.restore_from_xlsx)(p)
                    except ValueError as e:
                        return json.dumps({"ok": False, "error": str(e)})
                    except Exception:
                        return json.dumps({"ok": False, "error": "The database inside this backup could not be read - nothing was changed."})
                # Data restored successfully - now the settings files and older backups
                for fname, fpath in [("auth.json", AUTH_FILE), ("config.json", CONFIG_FILE)]:
                    if fname in names:
                        _atomic_write_bytes(fpath, zf.read(fname))
                os.makedirs(BACKUP_DIR, exist_ok=True)
                for name in names:
                    if name.startswith("backups/") and name.endswith((".db", ".xlsx")):
                        fname = os.path.basename(name)
                        if fname and not fname.startswith("."):
                            _atomic_write_bytes(os.path.join(BACKUP_DIR, fname), zf.read(name))
            return json.dumps({"ok": True})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def open_excel(self):
        """Export an up-to-date Excel copy and open it (the live data lives in the database)."""
        import subprocess
        try:
            path = self.engine.export_for_user()
            if sys.platform == 'win32':    os.startfile(path)
            elif sys.platform == 'darwin': subprocess.Popen(['open', path])
            else:                          subprocess.Popen(['xdg-open', path])
            return json.dumps({"ok": True, "path": path})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    # ── AUTH ──────────────────────────────────
    # ── Licence (per computer, offline; see licence_gate.py) ────────────────
    def licence_state(self):
        return json.dumps({"ok": True, "licence": licence_gate.public_state()})

    def activate_licence(self, licence):
        """Activate this copy, or renew it. A renewal must be for the same person who holds the current licence."""
        try:
            if not isinstance(licence, str) or len(licence) > 4096:
                raise LicenceError("That doesn't look like a licence. Paste the whole text the owner sent you.")
            lic = licence_gate.get_licensing()
            current = lic.state()
            holder = current.get("user") if current.get("kind") == "licensed" else None
            lic.activate(licence, expect_user=holder)
            licence_gate.reset_cache()
            return json.dumps({"ok": True, "licence": licence_gate.public_state()})
        except LicenceError as e:
            return json.dumps({"ok": False, "error": str(e)})

    def claim_owner(self, owner_key):
        """The owner proves it is their own computer, once. The key is used in memory and never stored."""
        try:
            if not isinstance(owner_key, str) or len(owner_key) > 512:
                raise LicenceError("That isn't an owner key.")
            licence_gate.get_licensing().claim_owner(owner_key)
            licence_gate.reset_cache()
            return json.dumps({"ok": True, "licence": licence_gate.public_state()})
        except LicenceError as e:
            return json.dumps({"ok": False, "error": str(e)})

    def is_pin_set(self):
        has_user = False
        if os.path.exists(AUTH_FILE):
            try:
                with open(AUTH_FILE) as f: has_user = bool(json.load(f).get("username"))
            except Exception: pass
        return json.dumps({"ok": True, "set": os.path.exists(AUTH_FILE), "has_username": has_user})

    def set_pin(self, pin, username=None):
        try:
            existing = os.path.exists(AUTH_FILE)
            # Changing an existing PIN needs a logged-in session or a verified recovery key.
            if existing and not (self._authed or self._recovery_ok):
                return json.dumps({"ok": False, "error": "Not authorised"})
            if not isinstance(pin, str) or len(pin.strip()) < 4:
                return json.dumps({"ok": False, "error": "PIN must be at least 4 characters"})
            if existing:
                # Changing / resetting a PIN keeps the account name (installs made before accounts have none).
                try:
                    with open(AUTH_FILE) as f: user = json.load(f).get("username") or None
                except Exception: user = None
            else:
                user = _clean_username(username)      # first launch: the account name is required
            salt = secrets.token_hex(16)
            h    = _kdf(pin, salt)
            # Recovery key: XXXX-XXXX-XXXX-XXXX (shown once; only a hash is stored)
            raw_key = secrets.token_hex(8).upper()
            rkey_display = '-'.join([raw_key[i:i+4] for i in range(0, 16, 4)])
            rkey_salt = secrets.token_hex(16)
            rkey_hash = _kdf(rkey_display, rkey_salt)
            rkey_hint = '-'.join([p[0]+'***' for p in rkey_display.split('-')])
            auth = {"hash": h, "salt": salt, "kdf": "pbkdf2", "iters": _PBKDF2_ITERS,
                    "rkey_hash": rkey_hash, "rkey_salt": rkey_salt, "rkey_kdf": "pbkdf2",
                    "rkey_iters": _PBKDF2_ITERS, "rkey_hint": rkey_hint}
            if user: auth["username"] = user
            _atomic_write_json(AUTH_FILE, auth)
            self._authed = True; self._recovery_ok = False; self._user = user; self._register_ok()
            return json.dumps({"ok": True, "recovery_key": rkey_display, "username": user})
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})

    def get_recovery_hint(self):
        try:
            if not os.path.exists(AUTH_FILE):
                return json.dumps({"ok": False, "key": None})
            with open(AUTH_FILE) as f:
                data = json.load(f)
            # Older installs stored the key readable; new PINs store only a masked hint.
            key = data.get('rkey_plain') or data.get('rkey_hint') or None
            return json.dumps({"ok": True, "key": key})
        except Exception as e:
            return json.dumps({"ok": False, "key": None, "error": str(e)})

    def verify_recovery_key(self, key):
        try:
            msg = self._throttle()
            if msg: return json.dumps({"valid": False, "error": msg})
            if not os.path.exists(AUTH_FILE):
                return json.dumps({"valid": False})
            with open(AUTH_FILE) as f:
                data = json.load(f)
            rkey_hash = data.get('rkey_hash'); rkey_salt = data.get('rkey_salt')
            if not rkey_hash or not rkey_salt:
                return json.dumps({"valid": False})
            cand = _kdf(str(key).strip().upper(), rkey_salt, data.get('rkey_iters'), legacy=(data.get('rkey_kdf') != 'pbkdf2'))
            ok = hmac.compare_digest(cand, rkey_hash)
            if ok: self._recovery_ok = True; self._register_ok()
            else:  self._register_fail()
            resp = {"valid": ok}
            if ok and data.get("username"): resp["username"] = data["username"]
            return json.dumps(resp)
        except Exception as e:
            return json.dumps({"valid": False, "error": str(e)})

    def verify_pin(self, pin, username=None):
        try:
            msg = self._throttle()
            if msg: return json.dumps({"ok": True, "valid": False, "error": msg})
            if not os.path.exists(AUTH_FILE):
                return json.dumps({"ok": False, "error": "No PIN set"})
            with open(AUTH_FILE) as f:
                data = json.load(f)
            legacy = data.get('kdf') != 'pbkdf2'
            cand = _kdf(str(pin), data['salt'], data.get('iters'), legacy=legacy)
            ok = hmac.compare_digest(cand, data['hash'])
            stored_user = data.get('username') or None
            if stored_user:       # accounts need the right username AND PIN (same error either way)
                given = str(username or "").strip().casefold().encode("utf-8")
                ok = hmac.compare_digest(given, stored_user.strip().casefold().encode("utf-8")) and ok
            if ok:
                self._authed = True; self._user = stored_user; self._register_ok()
                if legacy:   # silently upgrade an old SHA-256 PIN hash to PBKDF2
                    try:
                        salt = secrets.token_hex(16)
                        data.update({"hash": _kdf(str(pin), salt), "salt": salt, "kdf": "pbkdf2", "iters": _PBKDF2_ITERS})
                        _atomic_write_json(AUTH_FILE, data)
                    except Exception: pass
            else:
                self._register_fail()
            resp = {"ok": True, "valid": ok}
            if ok and stored_user: resp["username"] = stored_user
            return json.dumps(resp)
        except Exception as e:
            return json.dumps({"ok": False, "error": str(e)})


# ════════════════════════════════════════════
# LAUNCH
# ════════════════════════════════════════════
# Per-session secret. In Electron it is generated by the main process (ICT_API_TOKEN)
# and added to every call over IPC; in browser mode it is injected into the page.
API_TOKEN = os.environ.get("ICT_API_TOKEN") or secrets.token_hex(32)
MAX_BODY_BYTES = 768 * 1024 * 1024

BROWSER_BRIDGE = """
<script>
(function(){
  window.pywebview = window.pywebview || {};
  window.pywebview.api = new Proxy({}, {
    get: function(_, name) {
      return function() {
        return fetch('/api/' + encodeURIComponent(name), {
          method: 'POST',
          headers: {'Content-Type': 'application/json', 'X-ICT-Token': '__ICT_TOKEN__'},
          body: JSON.stringify({args: Array.prototype.slice.call(arguments)})
        }).then(function(r) { return r.text(); });
      };
    }
  });
  // Mark that we are running in browser mode (not pywebview desktop)
  window.__IS_BROWSER_MODE__ = true;
  // Wait for DOMContentLoaded, then yield with setTimeout(0) so that
  // all synchronous <script> blocks at the bottom of <body> have had a
  // chance to register their 'pywebviewready' listeners before we fire.
  function dispatchReady() {
    window.dispatchEvent(new Event('pywebviewready'));
  }
  if (document.readyState === 'loading') {
    window.addEventListener('DOMContentLoaded', function(){ setTimeout(dispatchReady, 0); });
  } else {
    setTimeout(dispatchReady, 0);
  }
})();
</script>
"""

def ecosystem_card():
    """The app's public "capability card" (who it is and what it offers), read as written from src/ecosystem.json so it can be edited without touching
    code. Returns the JSON bytes, or None when the file is missing or not a card, in which case the address simply answers 404. No data is in it."""
    try:
        with open(os.path.join(BUNDLE_DIR, "ecosystem.json"), "rb") as f:
            raw = f.read(64 * 1024)
        card = json.loads(raw)
        if isinstance(card, dict) and all(isinstance(card.get(k), str) and card[k] for k in ("contract", "id", "name")):
            return json.dumps(card).encode("utf-8")
    except (OSError, ValueError):
        pass
    return None


def make_handler(api, ui_path, browser_mode=False):
    class AppHandler(http.server.BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            return

        CSP = ("default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; "
               "img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'; "
               "object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'")

        def _security_headers(self):
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("X-Frame-Options", "DENY")
            self.send_header("Referrer-Policy", "no-referrer")
            self.send_header("Cache-Control", getattr(self, "_cache_ctl", "no-store"))
            self.send_header("Content-Security-Policy", self.CSP)

        def end_headers(self):
            self._security_headers()
            super().end_headers()

        def _send(self, status, body, content_type="text/plain; charset=utf-8"):
            if isinstance(body, str):
                body = body.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _origin_ok(self):
            """Block DNS-rebinding (bad Host) and cross-site requests (foreign Origin)."""
            port = self.server.server_port
            hosts = {f"127.0.0.1:{port}", f"localhost:{port}"}
            if self.headers.get("Host", "") not in hosts:
                return False
            origin = self.headers.get("Origin")
            if origin and origin not in {f"http://{h}" for h in hosts}:
                return False
            return True

        def do_GET(self):
            if not self._origin_ok():
                self._send(403, "Forbidden"); return
            path = urllib.parse.urlparse(self.path).path
            # The public capability card for other DericBI systems (see docs/ECOSYSTEM.md): no data, so no login or licence needed.
            if path == "/api/ecosystem/card":
                card = ecosystem_card()
                if card is None:
                    self._send(404, "Not found"); return
                self._send(200, card, "application/json; charset=utf-8"); return
            if (path.startswith("/reports/") or path.startswith("/backups/")) and not licence_gate.allows(None, licence_gate.current_state()):
                self._send(402, "Licence required"); return
            if (path.startswith("/reports/") or path.startswith("/backups/")) and api.needs_login():
                self._send(401, "Locked"); return
            if path in ("/", "/ui.html"):
                with open(ui_path, "r", encoding="utf-8") as f:
                    html = f.read()
                # Only inject the fetch-proxy bridge in browser mode.
                # In desktop (pywebview) mode pywebview injects its own native
                # js_api bridge — injecting ours would overwrite it.
                if browser_mode:
                    html = html.replace("</head>", BROWSER_BRIDGE.replace("__ICT_TOKEN__", API_TOKEN) + "\n</head>", 1)
                self._send(200, html, "text/html; charset=utf-8")
                return
            # Self-hosted UI fonts (whitelisted file names only)
            if path.startswith("/fonts/"):
                fname = urllib.parse.unquote(path[len("/fonts/"):])
                fpath = os.path.join(BUNDLE_DIR, "fonts", fname)
                if re.fullmatch(r"[A-Za-z0-9._-]+\.woff2", fname) and os.path.isfile(fpath):
                    with open(fpath, "rb") as ff:
                        data = ff.read()
                    self._cache_ctl = "public, max-age=604800, immutable"
                    self.send_response(200)
                    self.send_header("Content-Type", "font/woff2")
                    self.send_header("Content-Length", str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                    return
                self._send(404, "Not found"); return
            # Serve logo image over HTTP (needed in Electron mode — file:// blocked)
            if path == "/logo":
                for lname in ['logo.png','logo.jpg','logo.jpeg','LOGO.PNG','LOGO.JPG']:
                    lp = os.path.join(BASE_DIR, lname)
                    if os.path.isfile(lp):
                        ext = os.path.splitext(lname)[1].lower()
                        mime = 'image/jpeg' if ext in ('.jpg','.jpeg') else 'image/png'
                        with open(lp, 'rb') as lf:
                            data = lf.read()
                        self.send_response(200)
                        self.send_header('Content-Type', mime)
                        self.send_header('Content-Length', str(len(data)))
                        self.end_headers()
                        self.wfile.write(data)
                        return
                self._send(404, "No logo")
                return
            # Serve generated PDF reports for browser download
            if path.startswith("/reports/"):
                fname = urllib.parse.unquote(path[len("/reports/"):])
                # Sanitise: no path traversal
                if fname and "/" not in fname and "\\" not in fname:
                    fpath = os.path.join(BASE_DIR, "reports", fname)
                    if os.path.isfile(fpath):
                        with open(fpath, "rb") as pf:
                            data = pf.read()
                        self.send_response(200)
                        self.send_header("Content-Type", "application/pdf")
                        self.send_header("Content-Length", str(len(data)))
                        self.send_header("Content-Disposition",
                                         f'inline; filename="{fname}"')
                        self.end_headers()
                        self.wfile.write(data)
                        return
            # Serve backup files for download
            if path.startswith("/backups/"):
                fname = urllib.parse.unquote(path[len("/backups/"):])
                if fname and "/" not in fname and "\\" not in fname and fname.endswith((".xlsx", ".db")):
                    fpath = os.path.join(BACKUP_DIR, fname)
                    if os.path.isfile(fpath):
                        with open(fpath, "rb") as bf:
                            data = bf.read()
                        self.send_response(200)
                        self.send_header("Content-Type", "application/octet-stream" if fname.endswith(".db") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        self.send_header("Content-Length", str(len(data)))
                        self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
                        self.end_headers()
                        self.wfile.write(data)
                        return
            self._send(404, "Not found")

        def do_POST(self):
            if not self._origin_ok():
                self._send(403, json.dumps({"ok": False, "error": "Forbidden"})); return
            if not hmac.compare_digest(self.headers.get("X-ICT-Token", ""), API_TOKEN):
                self._send(403, json.dumps({"ok": False, "error": "Forbidden"})); return
            path = urllib.parse.urlparse(self.path).path
            if not path.startswith("/api/"):
                self._send(404, "Not found")
                return
            name = urllib.parse.unquote(path.split("/api/", 1)[1])
            if name.startswith("_") or name.isupper() or not callable(getattr(api, name, None)) or name in ("needs_login",):
                self._send(404, json.dumps({"ok": False, "error": "Unknown API method"}))
                return
            if name not in licence_gate.LICENCE_METHODS:
                lic_state = licence_gate.current_state()
                if not licence_gate.allows(name, lic_state):
                    self._send(402, json.dumps(licence_gate.refusal(lic_state, name))); return
            if name not in api.PUBLIC_METHODS and api.needs_login():
                self._send(401, json.dumps({"ok": False, "locked": True, "error": "Session locked — please log in again"}))
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length < 0 or length > MAX_BODY_BYTES:
                    self._send(413, json.dumps({"ok": False, "error": "Request too large"})); return
                payload = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
                args = payload.get("args", [])
                if not isinstance(args, list):
                    self._send(400, json.dumps({"ok": False, "error": "Bad request"})); return
                result = getattr(api, name)(*args)
                self._send(200, result)
            except (ValueError, TypeError) as e:
                self._send(400, json.dumps({"ok": False, "error": str(e)}))
            except Exception as e:
                self._send(500, json.dumps({"ok": False, "error": str(e)}))

    return AppHandler

HOME_PORT = 47615   # fixed local address, http://127.0.0.1:47615 (the ecosystem card is at /api/ecosystem/card)

class LocalServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True

def run_browser(api, ui_path, electron_mode=False):
    # electron_mode: bridge injected by preload.js; serve raw html, no webbrowser.open
    handler = make_handler(api, ui_path, browser_mode=not electron_mode)
    try:
        server = LocalServer(("127.0.0.1", HOME_PORT), handler)   # same address every time, so other systems can find it
    except OSError:
        server = LocalServer(("127.0.0.1", 0), handler)           # taken (second copy running): any free port
    port   = server.server_port
    url    = f"http://127.0.0.1:{port}/"
    # Expose port so get_logo_path() can build an HTTP URL for the logo
    os.environ['ICT_HTTP_PORT'] = str(port)
    if electron_mode:
        print(f"ELECTRON_PORT={port}", flush=True)
    else:
        print("DericBI Inventory Control is running locally.")
        print("Open this address if the browser does not appear:")
        print(url)
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nClosing DericBI Inventory Control.")
    finally:
        server.server_close()

def run_desktop(api, ui_path):
    try:
        import webview
    except Exception as e:
        print("Desktop mode needs pywebview. Use start.bat for browser mode, or install pywebview first.")
        print(e)
        sys.exit(1)

    # Desktop mode: pywebview loads ui.html directly via file:// and exposes
    # js_api as window.pywebview.api natively. No HTTP server or bridge injection needed.
    window = webview.create_window(
        title            = "DericBI Inventory Control v2",
        url              = ui_path,
        js_api           = api,
        width            = 1440,
        height           = 880,
        min_size         = (1024, 680),
        background_color = "#0d1117",
        confirm_close    = True,
        text_select      = True,
    )
    api.window = window
    webview.start(debug=False, private_mode=False)

# ════════════════════════════════════════════
# PROFESSIONAL PDF REPORT ENGINE  (2-page max)
# ════════════════════════════════════════════
def generate_pdf_report(engine, output_path: str,
                        dept_filter: str = "", type_filter: str = "",
                        logo_path: str = "", date_from: str = "", date_to: str = "") -> dict:
    """
    Generate a concise 2-page professional ICT Asset Report.
    Page 1 : Header + KPI strip + Executive Summary + Equipment table + Department table
    Page 2 : (only if needed) Attention list + Recent Activity + sign-off
    No blank pages. No personal names. No redundant sections.
    """
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
        from reportlab.platypus import (
            Paragraph, Spacer, Table, TableStyle,
            HRFlowable
        )
        from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
        from reportlab.lib.colors import HexColor
        import datetime, os, collections, re

        # ── PALETTE ────────────────────────────────────────
        C_BRAND  = HexColor('#0A1532')
        C_GOLD   = HexColor('#FDC819')
        C_GREY   = HexColor('#475569')
        C_GREY_LT= HexColor('#f9f6f4')
        C_GREY_MD= HexColor('#e8e0dc')
        C_WHITE  = colors.white
        C_BLACK  = HexColor('#1a0e0b')
        C_STRIPE = HexColor('#fdf9f7')

        W, H = A4
        ML, MR, MT, MB = 1.8*cm, 1.8*cm, 2.4*cm, 2.2*cm
        BW = W - ML - MR

        # ── DATA ───────────────────────────────────────────
        data       = engine.load_all()
        all_rows   = data['inventory']
        hist_items = data['history']

        def bucket(c):
            c = (c or '').strip()
            if c in ('Replaced','Disposed'): return c.lower()
            if 'repair' in c.lower(): return 'repair'
            if c == 'New': return 'new'
            return 'good'

        filtered      = [r for r in all_rows
                         if (not dept_filter or r.get('Department/Location','') == dept_filter)
                         and (not type_filter or r.get('Equipment Type','') == type_filter)]
        active        = [r for r in filtered if bucket(r.get('Condition/Status')) not in ('replaced','disposed')]
        replaced_rows = [r for r in filtered if bucket(r.get('Condition/Status')) == 'replaced']
        disposed_rows = [r for r in filtered if bucket(r.get('Condition/Status')) == 'disposed']

        total    = len(active)
        n_new    = sum(1 for r in active if bucket(r.get('Condition/Status')) == 'new')
        n_good   = sum(1 for r in active if bucket(r.get('Condition/Status')) == 'good')
        n_repair = sum(1 for r in active if bucket(r.get('Condition/Status')) == 'repair')
        n_ok     = n_new + n_good
        pct_ok   = round(n_ok / total * 100) if total else 0

        type_counts = collections.Counter(r.get('Equipment Type','Unknown') or 'Unknown' for r in active)
        dept_counts = collections.Counter(r.get('Department/Location','Unknown') or 'Unknown' for r in active)

        today_d = datetime.date.today()
        cutoff  = today_d - datetime.timedelta(days=90)
        def pdate(s):
            try: return datetime.datetime.strptime(str(s)[:10],'%Y-%m-%d').date()
            except: return None
        # Apply date range: if specified use it, else default to last 90 days
        def in_range(h):
            d = pdate(h.get('date',''))
            if d is None: return False
            if date_from:
                try:
                    if d < datetime.datetime.strptime(date_from,'%Y-%m-%d').date(): return False
                except: pass
            if date_to:
                try:
                    if d > datetime.datetime.strptime(date_to,'%Y-%m-%d').date(): return False
                except: pass
            if not date_from and not date_to:
                if d < cutoff: return False
            return True
        recent = [h for h in hist_items if in_range(h)]

        total_cost = 0.0
        for h in recent:
            m = re.search(r'\[Cost: KES ([0-9,.]+)\]', h.get('note',''))
            if m:
                try: total_cost += float(m.group(1).replace(',',''))
                except: pass

        now_str  = datetime.datetime.now().strftime('%d %B %Y  %H:%M')
        date_str = datetime.datetime.now().strftime('%d %B %Y')
        scope    = dept_filter or type_filter or 'All Assets'
        top_type = type_counts.most_common(1)[0][0] if type_counts else '—'
        top_dept = dept_counts.most_common(1)[0][0] if dept_counts else '—'

        # ── STYLES ─────────────────────────────────────────
        def S(name, **kw): return ParagraphStyle(name, **kw)
        s_h2   = S('H2', fontSize=9.5, fontName='Helvetica-Bold', textColor=C_BRAND,
                   spaceAfter=2, spaceBefore=5, leading=12)
        s_exec = S('Exec', fontSize=8.5, fontName='Helvetica', textColor=C_BLACK,
                   leading=13, alignment=TA_JUSTIFY, spaceAfter=3)

        # ── PAGE TEMPLATE ──────────────────────────────────
        def make_doc(path):
            class MyDoc(BaseDocTemplate):
                def __init__(self, fn, **kw):
                    super().__init__(fn, **kw)
                    frame = Frame(ML, MB, BW, H - MT - MB,
                                  leftPadding=0, rightPadding=0,
                                  topPadding=0, bottomPadding=0, id='main')
                    self.addPageTemplates([
                        PageTemplate(id='main', frames=[frame], onPage=self._hf)])

                def _hf(self, cv, doc):
                    cv.saveState()
                    # ── Header band ──
                    cv.setFillColor(C_BRAND)
                    cv.rect(0, H - 1.4*cm, W, 1.4*cm, fill=1, stroke=0)
                    cv.setFillColor(C_GOLD)
                    cv.rect(0, H - 1.42*cm, W, 0.04*cm, fill=1, stroke=0)

                    lx = ML
                    if logo_path and os.path.exists(logo_path):
                        try:
                            from reportlab.lib.utils import ImageReader
                            img = ImageReader(logo_path)
                            iw, ih = img.getSize()
                            lh = 1.0*cm
                            lw = lh * iw / ih
                            cv.drawImage(logo_path, lx, H - 1.2*cm,
                                         width=lw, height=lh,
                                         preserveAspectRatio=True, mask='auto')
                            lx += lw + 6
                        except: pass

                    cv.setFillColor(C_WHITE)
                    cv.setFont('Helvetica-Bold', 8)
                    cv.drawString(lx, H - 0.75*cm, 'DericBI Inventory Control')
                    cv.setFont('Helvetica', 7)
                    cv.setFillColor(HexColor('#b8c4e0'))
                    cv.drawString(lx, H - 1.15*cm, f'ICT Asset Report  ·  {scope}  ·  Cut Through Noise')
                    cv.setFillColor(C_WHITE)
                    cv.setFont('Helvetica-Bold', 7)
                    cv.drawRightString(W - MR, H - 0.75*cm, date_str)
                    cv.setFont('Helvetica', 6.5)
                    cv.setFillColor(HexColor('#b8c4e0'))
                    cv.drawRightString(W - MR, H - 1.15*cm, f'Page {doc.page}  ·  CONFIDENTIAL')

                    # ── Footer ──
                    cv.setStrokeColor(C_GREY_MD)
                    cv.setLineWidth(0.5)
                    cv.line(ML, MB - 0.3*cm, W - MR, MB - 0.3*cm)
                    cv.setFillColor(C_GREY)
                    cv.setFont('Helvetica', 6.5)
                    cv.drawString(ML, MB - 0.55*cm,
                        f'DericBI Inventory Control  ·  Asset Report  ·  Generated {now_str}')
                    cv.drawRightString(W - MR, MB - 0.55*cm, 'CONFIDENTIAL — ICT USE ONLY')
                    cv.restoreState()

            return MyDoc(path, pagesize=A4,
                         leftMargin=ML, rightMargin=MR,
                         topMargin=MT, bottomMargin=MB,
                         title='ICT Asset Report',
                         author='DericBI Inventory Control')

        # ── HELPERS ────────────────────────────────────────
        def sp(h=4): return Spacer(1, h)
        def hr():    return HRFlowable(width='100%', thickness=0.6, color=C_GOLD,
                                       spaceAfter=3, spaceBefore=1)

        def sh(text):
            return [sp(4), Paragraph(text, s_h2), hr()]

        def dtable(headers, rows_data, col_widths, hbg=C_BRAND, fs=7.5):
            _th = ParagraphStyle('th', fontSize=fs-0.5, fontName='Helvetica-Bold',
                                 textColor=C_WHITE, leading=fs+1, alignment=TA_CENTER)
            _td = ParagraphStyle('td', fontSize=fs, fontName='Helvetica',
                                 textColor=C_BLACK, leading=fs+1.5, alignment=TA_LEFT)
            _tc = ParagraphStyle('tc', fontSize=fs, fontName='Helvetica',
                                 textColor=C_BLACK, leading=fs+1.5, alignment=TA_CENTER)

            tdata = [[Paragraph(h, _th) for h in headers]]
            for row in rows_data:
                tdata.append([Paragraph(str(c or '—'), _tc if i == 0 else _td)
                               for i, c in enumerate(row)])

            t = Table(tdata, colWidths=col_widths, repeatRows=1)
            style = [
                ('BACKGROUND',    (0,0), (-1,0), hbg),
                ('TOPPADDING',    (0,0), (-1,-1), 2),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                ('LEFTPADDING',   (0,0), (-1,-1), 4),
                ('RIGHTPADDING',  (0,0), (-1,-1), 4),
                ('GRID',          (0,0), (-1,-1), 0.3, C_GREY_MD),
                ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
                ('ROWBACKGROUNDS',(0,1), (-1,-1), [C_WHITE, C_STRIPE]),
            ]
            if rows_data and str(rows_data[-1][0]).upper().startswith('TOTAL'):
                style += [
                    ('FONTNAME',   (0,-1), (-1,-1), 'Helvetica-Bold'),
                    ('BACKGROUND', (0,-1), (-1,-1), C_GREY_LT),
                    ('LINEABOVE',  (0,-1), (-1,-1), 0.6, C_GREY),
                ]
            t.setStyle(TableStyle(style))
            return t

        # ── KPI STRIP ──────────────────────────────────────
        def kpi_strip():
            items = [
                ('Total Active Assets', str(total),            '#0A1532'),
                ('Good / New',          f'{n_ok}  ({pct_ok}%)','#1a7a3f'),
                ('Need Attention',      str(n_repair),          '#b45309'),
                ('Replaced / Retired',  str(len(replaced_rows)+len(disposed_rows)), '#C49A1A'),
                ('Departments',         str(len(dept_counts)),  '#475569'),
            ]
            n   = len(items)
            cw  = BW / n
            vals = [Paragraph(
                        f'<font size="18"><b><font color="{c}">{v}</font></b></font>',
                        ParagraphStyle('kv', alignment=TA_CENTER, leading=22))
                    for _, v, c in items]
            lbls = [Paragraph(
                        f'<font color="#475569"><font size="6.5">{l}</font></font>',
                        ParagraphStyle('kl', alignment=TA_CENTER, leading=9))
                    for l, _, _ in items]
            t = Table([vals, lbls], colWidths=[cw]*n, rowHeights=[26, 13])
            t.setStyle(TableStyle([
                ('BACKGROUND',    (0,0), (-1,-1), C_GREY_LT),
                ('BOX',           (0,0), (-1,-1), 0.4, C_GREY_MD),
                ('INNERGRID',     (0,0), (-1,-1), 0.4, C_GREY_MD),
                ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
                ('TOPPADDING',    (0,0), (-1,0),  4),
                ('BOTTOMPADDING', (0,1), (-1,1),  4),
            ]))
            return t

        # ════════════════════════════════════════════════
        # BUILD STORY
        # ════════════════════════════════════════════════
        story = []

        # ── KPI strip ──────────────────────────────────
        story.append(kpi_strip())
        story.append(sp(6))

        # ── Executive Summary ───────────────────────────
        story += sh('Executive Summary')
        repair_note = (f'{n_repair} asset{"s" if n_repair!=1 else ""} currently require maintenance.' if n_repair
                       else 'No assets are currently flagged for repair.')
        cost_note   = (f' Maintenance costs (last 90 days): <b>KES {total_cost:,.0f}</b>.' if total_cost else '')
        lifecycle_note = (f' {len(replaced_rows)} unit{"s" if len(replaced_rows)!=1 else ""} retired under lifecycle management.' if replaced_rows else '')
        story.append(Paragraph(
            f'The ICT Department manages <b>{total}</b> active asset{"s" if total!=1 else ""} across '
            f'<b>{len(dept_counts)}</b> department{"s" if len(dept_counts)!=1 else ""}. '
            f'<b>{n_ok}</b> ({pct_ok}%) are in good or new condition. {repair_note}'
            f'{cost_note}{lifecycle_note} '
            f'Leading category: <b>{top_type}</b>. Highest asset count: <b>{top_dept}</b>.',
            s_exec))

        # ── Equipment by Type — full width with row number column ──
        story += sh('Asset Breakdown by Equipment Type')

        # Row number | Equipment Type | Total | Good | Repair | %
        # Full page width; narrow No. col on the left margin
        NO_W  = BW * 0.055   # "#" column
        TY_W  = BW * 0.385   # Equipment Type
        TO_W  = BW * 0.12    # Total
        GD_W  = BW * 0.13    # Good
        RP_W  = BW * 0.13    # Repair
        PC_W  = BW * 0.18    # %
        type_cw_full = [NO_W, TY_W, TO_W, GD_W, RP_W, PC_W]

        type_rows = []
        for idx, (etype, cnt) in enumerate(type_counts.most_common(), 1):
            tr  = [r for r in active if r.get('Equipment Type','') == etype]
            rep = sum(1 for r in tr if bucket(r.get('Condition/Status')) == 'repair')
            ok  = cnt - rep
            type_rows.append([str(idx), etype, str(cnt), str(ok), str(rep),
                               f'{round(cnt/total*100)}%' if total else '0%'])
        type_rows.append(['—', 'TOTAL', str(total), str(n_ok), str(n_repair), '100%'])

        story.append(dtable(
            ['No.', 'Equipment Type', 'Total', 'Good', 'Repair', '%'],
            type_rows, type_cw_full))

        # ── Assets Needing Attention ────────────────────
        repair_assets = [r for r in active if bucket(r.get('Condition/Status')) == 'repair']
        if repair_assets:
            story += sh('Assets Requiring Attention')
            r_rows = [[
                r.get('Equipment Type',''),
                r.get('Brand/Model',''),
                r.get('Serial No.',''),
                r.get('Department/Location',''),
                r.get('Condition/Status',''),
            ] for r in repair_assets]
            cw_rep = [BW*0.18, BW*0.22, BW*0.18, BW*0.26, BW*0.16]
            story.append(dtable(
                ['Type','Brand / Model','No.','Department','Status'],
                r_rows, cw_rep, hbg=HexColor('#0A1532')))

        # ── Recent Activity ─────────────────────────────
        if recent:
            story += sh('Recent Activity — Last 90 Days')
            ev_counts = collections.Counter(h.get('event','other') for h in recent)
            ev_rows   = [[ev.title(), str(cnt)] for ev, cnt in ev_counts.most_common()]
            ev_rows.append(['TOTAL', str(len(recent))])
            activity_note = (
                f'{len(recent)} event{"s" if len(recent)!=1 else ""} recorded in the past 90 days'
                f'{f" · KES {total_cost:,.0f} in logged costs" if total_cost else ""}.'
            )
            story.append(Paragraph(activity_note, s_exec))
            story.append(dtable(['Event Type','Count'], ev_rows,
                                 [BW*0.55, BW*0.20]))

        # ── Lifecycle (only if data exists) ─────────────
        if replaced_rows or disposed_rows:
            story += sh('Lifecycle Summary')
            rep_types = collections.Counter(r.get('Equipment Type','Unknown') or 'Unknown' for r in replaced_rows)
            dis_types = collections.Counter(r.get('Equipment Type','Unknown') or 'Unknown' for r in disposed_rows)
            all_et  = sorted(set(list(rep_types)+list(dis_types)))
            lc_rows = [[et, str(rep_types.get(et,0)), str(dis_types.get(et,0))] for et in all_et]
            lc_rows.append(['TOTAL', str(len(replaced_rows)), str(len(disposed_rows))])
            story.append(dtable(['Equipment Type','Replaced','Disposed'],
                                 lc_rows, [BW*0.60, BW*0.20, BW*0.20]))

        # ── Sign-off ─────────────────────────────────────
        story.append(sp(8))
        story.append(HRFlowable(width='100%', thickness=0.5, color=C_GREY_MD,
                                spaceAfter=4, spaceBefore=4))
        story.append(Paragraph(
            f'<font color="#0A1532"><b>DericBI Inventory Control</b></font>'
            f'  &nbsp;·&nbsp;  <i>Cut Through Noise</i>'
            f'  &nbsp;·&nbsp;  <font color="#94a3b8">Generated {now_str}  ·  CONFIDENTIAL</font>',
            ParagraphStyle('sf', fontSize=7, fontName='Helvetica',
                           alignment=TA_CENTER, textColor=C_GREY)))

        # ── BUILD ─────────────────────────────────────────
        doc = make_doc(output_path)
        doc.build(story)
        engine.log_session('REPORT', f'PDF report generated: {os.path.basename(output_path)}')
        return {"ok": True, "path": output_path}

    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e), "trace": traceback.format_exc()}


def main():
    try:
        seed = next((p for p in (os.path.join(BUNDLE_DIR, "ICT_MASTER.xlsx"),
                                 os.path.join(BUNDLE_DIR, "..", "assets", "ICT_MASTER.xlsx")) if os.path.exists(p)), None)
        engine = Storage(BASE_DIR, BACKUP_DIR, seed_xlsx=seed)
    except Exception as e:
        msg = str(e)
        try:
            with open(os.path.join(BASE_DIR, "startup_error.txt"), "w", encoding="utf-8") as f: f.write(msg)
        except OSError: pass
        sys.stderr.write("STARTUP_ERROR: " + msg + "\n"); sys.stderr.flush()
        sys.exit(2)
    cfg = ConfigEngine()
    api = Api(engine, cfg)

    ui_path = os.path.join(BUNDLE_DIR, 'ui.html')
    if not os.path.exists(ui_path):
        print("ERROR: ui.html not found in", BUNDLE_DIR)
        sys.exit(1)

    if "--desktop" in sys.argv:
        run_desktop(api, ui_path)
    elif "--electron" in sys.argv:
        run_browser(api, ui_path, electron_mode=True)
    else:
        run_browser(api, ui_path)

if __name__ == "__main__":
    main()
