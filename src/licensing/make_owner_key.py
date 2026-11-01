#!/usr/bin/env python3
"""Make the owner key for ONE app. Run it once, on your own computer.

  python make_owner_key.py --prefix ACM --set-in backend/owner_key.py --var OWNER_PUBLIC_KEY

Writes the PRIVATE key file (default <prefix>-owner.key), refuses to overwrite one that exists, adds it to .gitignore,
and puts the PUBLIC key into the app's key file, but only if that file's key is still empty. An existing public key is
never replaced (a new key would invalidate every licence already issued).
"""
import argparse
import os
import sys

import _common

_common.use_modules()
import licensing  # noqa: E402


def main():
    cfg = _common.load_config()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--prefix", default=cfg.get("prefix"), help="2-6 capital letters that tag this app's keys and licences")
    ap.add_argument("--key-file", help="where to write the private key (default <prefix>-owner.key)")
    ap.add_argument("--set-in", default=cfg.get("public_key_file"), help="the app's file that holds OWNER_PUBLIC_KEY")
    ap.add_argument("--var", default=cfg.get("public_key_var", "OWNER_PUBLIC_KEY"))
    a = ap.parse_args()
    if not a.prefix:
        ap.error("--prefix is required (e.g. --prefix ACM)")
    key_file = a.key_file or f"{a.prefix.lower()}-owner.key"

    # Every refusal happens BEFORE a key is generated, so a refusal never leaves a half-made key behind.
    if os.path.exists(key_file):
        sys.exit(f"{key_file} already exists. Move it somewhere safe first: a new key would stop every licence you've issued working.")
    if a.set_in:
        current = _common.read_public_key(a.set_in, a.var)
        if current is None:
            sys.exit(f"Can't find {a.var} in {a.set_in}. Check --set-in/--var, or omit --set-in and paste the key in by hand.")
        if current:
            sys.exit(f"{a.set_in} already has a public key. Not touching it. If this app already has an owner key, keep using it.")

    key_text, public = licensing.Licensing.make_owner_key(a.prefix)
    with open(key_file, "x", encoding="utf-8") as f:
        f.write(key_text + "\n")
    try:
        os.chmod(key_file, 0o600)
    except OSError:
        pass
    added = _common.ensure_gitignored(os.path.basename(key_file))
    if a.set_in:
        _common.set_public_key(a.set_in, a.var, public)
    print(f"""
Made {os.path.abspath(key_file)}
  KEEP THIS FILE SECRET AND BACK IT UP (USB stick or password manager). If it is lost, no new licences can be issued.
  It is {'now in' if added else 'already in'} .gitignore. Never commit it, email it or paste it in a chat.

Public key (safe to publish): {public}
{'Written into ' + a.set_in + ' (' + a.var + ').' if a.set_in else 'Paste it into the app as ' + a.var + ' = "' + public + '" before you build the installer.'}
""")


if __name__ == "__main__":
    main()
