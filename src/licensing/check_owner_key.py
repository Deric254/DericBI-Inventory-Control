#!/usr/bin/env python3
"""Build guard. Fails the build when (1) the app has no valid owner public key, or (2) a private owner key was committed.

  python check_owner_key.py --file backend/owner_key.py --var OWNER_PUBLIC_KEY

Run it in CI before building the installer, so an installer can never ship without licence protection.
"""
import argparse
import os
import re
import sys

import _common


def main():
    cfg = _common.load_config()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--file", default=cfg.get("public_key_file"))
    ap.add_argument("--var", default=cfg.get("public_key_var", "OWNER_PUBLIC_KEY"))
    ap.add_argument("--no-scan", action="store_true", help="skip the committed-private-key scan")
    a = ap.parse_args()
    if not a.file:
        ap.error("--file is required (the file that holds the owner public key)")
    bad = []

    key = _common.read_public_key(a.file, a.var)
    if key is None:
        bad.append(f"Can't find {a.var} in {a.file}.")
    else:
        try:
            ok = len(bytes.fromhex(key)) == 32
        except ValueError:
            ok = False
        if not ok:
            bad.append(f"{a.file} has no valid {a.var} (64 hex characters). Run make_owner_key.py and set it before building.")

    if not a.no_scan:
        files = _common.tracked_files()
        if files is not None:
            private = re.compile(r"\b[A-Z]{2,6}-OWNER-[A-Za-z0-9_\-]{43}\b")
            for f in files:
                if f.endswith("-owner.key"):
                    bad.append(f"PRIVATE KEY COMMITTED: {f}. Remove it from git history, then make a NEW key: this one is no longer secret.")
                    continue
                try:
                    if os.path.getsize(f) < 2_000_000 and private.search(open(f, encoding="utf-8", errors="ignore").read()):
                        bad.append(f"PRIVATE KEY TEXT FOUND in committed file {f}. Remove it and make a NEW key: it is no longer secret.")
                except OSError:
                    pass
    if bad:
        sys.exit("\n".join(bad))
    print("Owner public key is set; no private key is committed.")


if __name__ == "__main__":
    main()
