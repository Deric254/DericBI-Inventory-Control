"""Shared test environment: ONE temp data dir, `main` imported once."""
import os, sys, shutil, tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC  = os.path.join(ROOT, "src")
TMP  = tempfile.mkdtemp(prefix="dericbi_test_")
SEED = os.path.join(ROOT, "assets", "ICT_MASTER.xlsx")
os.environ["ICT_DATA_DIR"]   = TMP
os.environ["ICT_BUNDLE_DIR"] = SRC
os.environ["ICT_API_TOKEN"]  = "test-token"
sys.path.insert(0, SRC)
import main as app          # noqa: E402
import storage              # noqa: E402
import licence_gate         # noqa: E402
from licensing.licensing import Licensing   # noqa: E402


class MemoryStore:
    def __init__(self): self.data = {}
    def get(self, key, default=None): return self.data.get(key, default)
    def set(self, key, value): self.data[key] = value


def make_licensing(today=None):
    """A Licensing over a THROWAWAY key (never the real one): (instance, owner key text, owner key seed)."""
    key_text, public = Licensing.make_owner_key(licence_gate.PREFIX)
    kw = {} if today is None else {"today": lambda: today}
    lic = Licensing(licence_gate.APP_ID, licence_gate.PREFIX, public, MemoryStore(), **kw)
    return lic, key_text, lic.parse_owner_key(key_text)


def use_owner_copy():
    """Every API call is refused unless the copy is licensed, so the suite runs as an owner copy made with a
    throwaway key. The gate itself stays fully active; test_licence.py swaps in locked copies to prove it refuses."""
    lic, key_text, _ = make_licensing(); lic.claim_owner(key_text); licence_gate.use_licensing(lic)


use_owner_copy()


def wipe(d=TMP):
    for f in os.listdir(d):
        p = os.path.join(d, f)
        shutil.rmtree(p, ignore_errors=True) if os.path.isdir(p) else os.remove(p)


def fresh(seed=True):
    """Empty data dir -> Storage (optionally seeded with the shipped sample workbook) + Api."""
    wipe()
    eng = storage.Storage(TMP, os.path.join(TMP, "backups"), seed_xlsx=SEED if seed else None)
    return eng, app.Api(eng, app.ConfigEngine())


def asset(**kw):
    d = {"Equipment Type": "Laptop", "Brand/Model": "Test", "Serial No.": "T-1",
         "Department/Location": "ICT", "Condition/Status": "Good"}
    d.update(kw)
    return d
