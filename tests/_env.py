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
