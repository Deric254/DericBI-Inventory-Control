#!/usr/bin/env python3
"""Look inside a licence and say whether it checks out. Handy when a customer says 'it doesn't work'.

  python verify_licence.py --token DBI1.xxxx.yyyy [--code ABCD-EFGH-JKLM]
"""
import argparse
import json
import re
import sys

import _common

_common.use_modules()
import licensing  # noqa: E402


class _Mem:
    def __init__(self):
        self.d = {}

    def get(self, k, default=None):
        return self.d.get(k, default)

    def set(self, k, v):
        self.d[k] = v


def main():
    cfg = _common.load_config()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--token", required=True)
    ap.add_argument("--code", help="the customer's activation code; if given, it must match")
    ap.add_argument("--prefix", default=cfg.get("prefix"))
    ap.add_argument("--app", default=cfg.get("app_id"))
    ap.add_argument("--public-key-file", default=cfg.get("public_key_file"))
    ap.add_argument("--var", default=cfg.get("public_key_var", "OWNER_PUBLIC_KEY"))
    ap.add_argument("--public-key")
    a = ap.parse_args()
    public = a.public_key or (_common.read_public_key(a.public_key_file, a.var) if a.public_key_file else None)
    if not (a.prefix and a.app and public):
        ap.error("need --prefix, --app and the public key (--public-key-file or --public-key)")
    try:
        body = json.loads(licensing._unb64(re.sub(r"\s+", "", a.token).split(".")[1]))
        print("Contents:", json.dumps(body, indent=2))
    except Exception:
        sys.exit("That is not a licence (can't read its contents).")
    lic = licensing.Licensing(a.app, a.prefix, public, _Mem())
    try:
        lic.read(a.token, for_install=a.code or body.get("for"))
        print("RESULT: valid" + ("" if a.code else " (computer not checked: pass --code to check it)"))
    except licensing.LicenceError as e:
        sys.exit(f"RESULT: not valid: {e}")


if __name__ == "__main__":
    main()
