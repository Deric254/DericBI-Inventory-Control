"""Shared helpers for the licensing tools. Standard library only."""
import json
import os
import re
import subprocess
import sys

sys.dont_write_bytecode = True   # the tools must not leave __pycache__ folders in someone's project

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_NAME = "licensing.json"


def use_modules():
    """Make licensing.py / ed25519.py importable: next to the scripts (after scaffold) or in the skill's assets."""
    for d in (HERE, os.path.join(HERE, "..", "assets", "python")):
        d = os.path.abspath(d)
        if os.path.exists(os.path.join(d, "licensing.py")) and os.path.exists(os.path.join(d, "ed25519.py")):
            if d not in sys.path:
                sys.path.insert(0, d)
            return d
    sys.exit("Can't find licensing.py and ed25519.py next to this script or in ../assets/python.")


def load_config(path=None) -> dict:
    p = path or os.path.join(HERE, CONFIG_NAME)
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def key_pattern(var: str):
    # NAME = "hex"  |  NAME: &str = "hex"  |  "NAME": "hex"  (Python, Rust, JS, C#, Go, JSON ...)
    return re.compile(r"(" + re.escape(var) + r"""["']?\s*[:=]\s*(?:[^"'\n]*?)["'])([0-9a-fA-F]*)(["'])""")


def read_public_key(file: str, var: str):
    """The hex string assigned to `var` in `file`, or None when the file/assignment is missing."""
    try:
        text = open(file, encoding="utf-8").read()
    except OSError:
        return None
    m = key_pattern(var).search(text)
    return m.group(2) if m else None


def set_public_key(file: str, var: str, hex_key: str) -> str:
    """Put the key in, but ONLY where it is currently empty. Returns 'set'. Raises if a key is already there."""
    text = open(file, encoding="utf-8").read()
    m = key_pattern(var).search(text)
    if not m:
        raise SystemExit(f"Can't find `{var} = \"\"` in {file}. Paste the key in by hand instead.")
    if m.group(2):
        raise SystemExit(f"{file} already has a public key. Not touching it: replacing it would stop every licence you've issued working.")
    open(file, "w", encoding="utf-8", newline="").write(text[:m.start(2)] + hex_key + text[m.end(2):])
    return "set"


def ensure_gitignored(name: str):
    """Append the private key file name to .gitignore (only ever appends)."""
    gi = ".gitignore"
    existing = open(gi, encoding="utf-8").read() if os.path.exists(gi) else ""
    if name not in existing.split():
        with open(gi, "a", encoding="utf-8") as f:
            f.write(("" if existing.endswith("\n") or not existing else "\n") + f"{name}\n")
        return True
    return False


def tracked_files():
    try:
        out = subprocess.run(["git", "ls-files"], capture_output=True, text=True, check=True).stdout
        return [l for l in out.splitlines() if l]
    except Exception:
        return None
