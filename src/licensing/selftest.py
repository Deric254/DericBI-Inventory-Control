#!/usr/bin/env python3
"""Proves the licensing module works end to end. Run it after scaffolding, and in CI.

  python selftest.py                       # tests the kit next to this script / in the skill
  python selftest.py --module-dir path     # tests the copy that was scaffolded into a project

Covers: the Ed25519 RFC 8032 test vectors, issue -> activate, wrong computer, wrong app, expired, grace, tampered
licence, wrong owner key, bad tokens, clock rolled back, owner claim, and that a build without a key fails closed.
"""
import argparse
import os
import sys
import tempfile
from datetime import date, timedelta

import _common

ap = argparse.ArgumentParser()
ap.add_argument("--module-dir")
args = ap.parse_args()
if args.module_dir:
    sys.path.insert(0, os.path.abspath(args.module_dir))
else:
    _common.use_modules()
import ed25519  # noqa: E402
import licensing  # noqa: E402
from licensing import Licensing, LicenceError, FileStore  # noqa: E402

PASSED = []


def check(name, cond):
    PASSED.append((name, bool(cond)))
    print(("ok   " if cond else "FAIL ") + name)


def raises(fn, text=""):
    try:
        fn()
    except (LicenceError, ValueError) as e:
        return text.lower() in str(e).lower()
    return False


# 1. the signature maths against the published vectors
for sk, pk, msg, sig in (
    ("9d61b19deffd5a60ba844af492ec2cc44449c5697b326919703bac031cae7f60", "d75a980182b10ab7d54bfed3c964073a0ee172f3daa62325af021a68f707511a", "",
     "e5564300c360ac729086e2cc806e828a84877f1eb8e5d974d873e065224901555fb8821590a33bacc61e39701cf9b46bd25bf5f0595bbe24655141438e7a100b"),
    ("4ccd089b28ff96da9db6c346ec114e0f5b8a319f35aba624da8cf6ed4fb8a6fb", "3d4017c3e843895a92b70aa74d1b7ebc9c982ccf2ec4968cc0cd55f12af4660c", "72",
     "92a009a9f0d4cab8720e820b5f642540a2b27b5416503f8fb3762223ebdb69da085ac1e43e15996e458f3613d0f11d8c387b2eaeb4302aeeb00d291612bb0c00"),
):
    sk, pk, msg, sig = map(bytes.fromhex, (sk, pk, msg, sig))
    check("RFC 8032 vector: key, signature, verify",
          ed25519.public_key(sk) == pk and ed25519.sign(sk, msg) == sig and ed25519.verify(pk, msg, sig)
          and not ed25519.verify(pk, msg + b"!", sig))


# 1b. the fixed conformance vector (the same one documented in references/token-spec.md; any language must agree)
_seed = bytes.fromhex("3cf49a98e3c3183b259755133b9edc39550b2083f6fde514d63ae67fac7ddf6b")
_pub = "7a25536be5bd7d9462fcbe457e004d298ecff116dfd4aafd4304364bb57c91cc"
_tok = ("CNF1.eyJhcHAiOiJDb25mb3JtYW5jZSBBcHAiLCJleHBpcmVzIjoiMjAyNi0wMS0zMSIsImZlYXR1cmVzIjpbInBybyJdLCJmb3IiOiJBQkNERUZHSEpLTE0iLCJpc3N1ZWQiOiIyMDI2LTAxLTAxIiwicm9sZSI6Im1lbWJlciIsInVzZXIiOiJUZXN0IFVzZXIiLCJ2IjoxfQ."
        "1jo45EXZ9tV8w-iPDEgyt5syxF5kC25qKn_yvckRIUf_ZTflW3pR5tqojK-45g79LB3J3pWOF6A-Abk1p2YiAg")
_cv = Licensing("Conformance App", "CNF", _pub, None, today=lambda: date(2026, 1, 1))
check("conformance vector: public key derives from the seed", ed25519.public_key(_seed).hex() == _pub)
check("conformance vector: the documented licence verifies", _cv.read(_tok, for_install="ABCD-EFGH-JKLM")["user"] == "Test User")
check("conformance vector: issuing reproduces the documented licence byte for byte",
      _cv.issue(_seed, "ABCD-EFGH-JKLM", "Test User", 30, ["pro"])["licence"] == _tok)

# 2. the licence lifecycle
tmp = tempfile.mkdtemp(prefix="lic-selftest-")
key_text, public = Licensing.make_owner_key("TST")
now = {"d": date(2026, 1, 1), "t": 1_800_000_000}


def make(name="a", app="Test App", prefix="TST", pub=None, **kw):
    return Licensing(app, prefix, public if pub is None else pub, FileStore(os.path.join(tmp, name)), today=lambda: now["d"], clock=lambda: now["t"], **kw)


owner = make("owner")
seed = owner.parse_owner_key(key_text)
cust = make("cust")
code = cust.install_id()
check("activation code looks right", len(licensing.norm_code(code)) == 12 and code.count("-") == 2)
check("activation code is stable", cust.install_id() == code)
check("unlicensed copy is refused", not cust.state()["ok"] and cust.state()["kind"] == "unlicensed")

issued = owner.issue(seed, code, "Jane Doe", 30)
check("expiry is today + days", issued["expires"] == "2026-01-31")
data = cust.activate(issued["licence"])
st = cust.state()
check("valid licence opens the app", st["ok"] and st["kind"] == "licensed" and st["days_left"] == 30 and st["user"] == "Jane Doe")
check("licence survives a restart", make("cust").state()["ok"])
check("licence tolerates pasted line breaks", make("c2").read(issued["licence"][:40] + "\n " + issued["licence"][40:], for_install=code)["user"] == "Jane Doe")

other = make("other")
check("licence is refused on another computer", raises(lambda: other.activate(issued["licence"]), "different computer"))
check("licence is refused in another app", raises(lambda: make("x", app="Other App").read(issued["licence"], for_install=code), "different app"))
check("renewal must be the same person", raises(lambda: cust.activate(owner.issue(seed, code, "Someone Else", 30)["licence"], expect_user="Jane Doe"), "different person"))

tag, payload, sig = issued["licence"].split(".")
forged_payload = licensing._b64(licensing._unb64(payload).replace(b"2026-01-31", b"2099-01-31"))
check("a licence with an edited expiry is refused", raises(lambda: cust.read(f"{tag}.{forged_payload}.{sig}"), "wasn't signed"))
other_key, other_pub = Licensing.make_owner_key("TST")
check("a licence signed by someone else's key is refused", raises(lambda: make("m", pub=other_pub).read(issued["licence"], for_install=code), "wasn't signed"))
check("the wrong owner key cannot issue", raises(lambda: owner.issue(owner.parse_owner_key(other_key), code, "X Y", 5), "doesn't belong"))
for junk in ("", "hello", "a.b", "TST1.a.b.c", "XXX1." + payload + "." + sig):
    check(f"junk token refused: {junk[:12]!r}", raises(lambda j=junk: cust.read(j), "doesn't look like"))
check("a bad owner key text is refused", raises(lambda: owner.parse_owner_key("not a key"), "isn't an owner key"))
check("activation code must be 12 characters", raises(lambda: owner.issue(seed, "ABC", "X Y", 5), "12 letters"))
check("days must be in range", raises(lambda: owner.issue(seed, code, "X Y", 0), "days") and raises(lambda: owner.issue(seed, code, "X Y", 366), "days"))
check("features are carried in the licence", make("f").read(owner.issue(seed, code, "F G", 5, ["pro", "api"])["licence"], for_install=code)["features"] == ["api", "pro"])

now["d"] = date(2026, 1, 31)
check("last day still works", cust.state()["ok"] and cust.state()["days_left"] == 0)
now["d"] = date(2026, 2, 1)
check("expired licence closes the app", not cust.state()["ok"] and "expired" in cust.state()["message"])
check("grace days keep it open", make("cust", grace_days=3).state()["ok"])
renewed = owner.issue(seed, code, "Jane Doe", 30)
cust.activate(renewed["licence"], expect_user="Jane Doe")
check("renewal reopens the app", cust.state()["ok"])

now["t"] += 10 * 86400
check("normal time passing is fine", cust.state()["ok"])
now["t"] -= 5 * 86400
check("clock set back more than a day is refused", not cust.state()["ok"] and "clock" in cust.state()["message"])
now["t"] += 5 * 86400

# 3. the owner's own copy
check("a copy cannot claim owner with a customer's key", raises(lambda: owner.claim_owner("TST-OWNER-" + "A" * 43), "doesn't belong"))
check("owner copy is not open before it is claimed", not make("own2").state()["ok"])
o2 = make("own2")
o2.claim_owner(key_text)
check("owner claim opens that copy for good", o2.state()["kind"] == "owner" and o2.state()["ok"])
now["d"] = date(2040, 1, 1)
check("owner copy never expires", o2.state()["ok"])
import json as _json
copied = os.path.join(tmp, "stolen")
os.makedirs(copied)
open(os.path.join(copied, "licensing.json"), "w").write(open(os.path.join(tmp, "own2", "licensing.json")).read().replace(o2.install_id(), "ZZZZ-ZZZZ-ZZZZ"))
check("owner proof copied to another install is useless", not make("stolen").state()["ok"])
forged = FileStore(os.path.join(tmp, "forge")); forged.set("owner_proof", "A" * 86)
check("a made-up owner proof is useless", not make("forge").state()["ok"])

# 4. fail closed
check("a build with no public key refuses everything", make("nokey", pub="").state()["kind"] == "unconfigured" and not make("nokey", pub="").state()["ok"])
check("a malformed public key also fails closed", not make("badkey", pub="zz" * 32).state()["ok"])

bad = [n for n, ok in PASSED if not ok]
print(f"\n{len(PASSED) - len(bad)}/{len(PASSED)} passed")
sys.exit(1 if bad else 0)
