"""Offline licensing for a locally installed app: Ed25519 signed licences, one per computer. Standard library only.

How it works (the same model for every app you ship):
  * The OWNER holds a private owner key (a file, kept secret). Its PUBLIC half is built into every copy of the app.
  * Each installed copy has an ACTIVATION CODE (a random id kept in its data folder).
  * The owner signs a LICENCE for exactly one activation code, one person and an expiry date.
  * The app only ever VERIFIES, so a copy of the app can check a licence but can never make one.
  * The owner's own computer proves ownership once with the owner key (a signed proof stored locally) and never expires.

Nothing here talks to the internet. The host app supplies a `store` (anything with get(key, default) / set(key, value)),
so it fits any existing storage: a settings table, a JSON file (FileStore below), the registry, a config file.
"""
import base64
import hmac
import json
import os
import re
import secrets
import time
from datetime import date, timedelta

try:                       # works as a package module or dropped next to ed25519.py
    from . import ed25519
except ImportError:        # pragma: no cover
    import ed25519

_CODE_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"   # no 0/O/1/I so codes can be read out over the phone


class LicenceError(ValueError):
    """Anything wrong with a key or licence, worded for a non-technical person."""


def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode().rstrip("=")


def _unb64(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def norm_code(code: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (code or "").upper())


def default_data_dir(app_name: str) -> str:
    """Where an installed copy keeps its own files on this OS (the same folder the app should already use)."""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif os.uname().sysname == "Darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    return os.path.join(base, app_name)


class FileStore:
    """A tiny JSON-file store, for apps that have no settings storage of their own. Writes are atomic."""

    def __init__(self, directory: str, filename: str = "licensing.json"):
        self.path = os.path.join(directory, filename)
        os.makedirs(directory, exist_ok=True)

    def _load(self) -> dict:
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def get(self, key, default=None):
        return self._load().get(key, default)

    def set(self, key, value):
        data = self._load()
        data[key] = value
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f)
        os.replace(tmp, self.path)


class Licensing:
    def __init__(self, app_id: str, prefix: str, public_key_hex: str, store, *, version: int = 1, max_days: int = 365,
                 grace_days: int = 0, rollback_seconds: int = 86400, today=date.today, clock=time.time):
        if not re.fullmatch(r"[A-Z]{2,6}", prefix or ""):
            raise ValueError("prefix must be 2 to 6 capital letters, e.g. DBI")
        self.app_id, self.prefix, self.version = app_id, prefix, int(version)
        self.store, self.max_days, self.grace_days = store, max_days, grace_days
        self.rollback_seconds, self._today, self._clock = rollback_seconds, today, clock
        try:
            raw = bytes.fromhex((public_key_hex or "").strip())
        except ValueError:
            raw = b""
        self.public = raw if len(raw) == 32 else None

    # ------------------------------------------------------------------ build-time facts
    def configured(self) -> bool:
        """False when the build has no valid owner public key: nothing can be licensed, and the app must fail closed."""
        return self.public is not None

    @property
    def tag(self) -> str:
        return f"{self.prefix}{self.version}"

    # ------------------------------------------------------------------ owner side (runs on the owner's computer)
    @staticmethod
    def make_owner_key(prefix: str):
        """(owner key text, public key hex). The text goes in the owner's secret key file; the hex goes in the app."""
        seed = secrets.token_bytes(32)
        return f"{prefix}-OWNER-" + _b64(seed), ed25519.public_key(seed).hex()

    def parse_owner_key(self, text: str) -> bytes:
        m = re.fullmatch(re.escape(self.prefix) + r"-OWNER-([A-Za-z0-9_\-]{43})", re.sub(r"\s+", "", text or ""))
        if not m:
            raise LicenceError(f"That isn't an owner key. It starts with {self.prefix}-OWNER- and is in the owner key file.")
        return _unb64(m.group(1))

    def owner_key_matches(self, seed: bytes) -> bool:
        return bool(self.public) and hmac.compare_digest(ed25519.public_key(seed), self.public)

    def issue(self, seed: bytes, code: str, user: str, days: int, features=None) -> dict:
        if not self.owner_key_matches(seed):
            raise LicenceError("That owner key doesn't belong to this app. Use the key made for this app's installer.")
        norm = norm_code(code)
        if len(norm) != 12:
            raise ValueError("The activation code is 12 letters and numbers, like ABCD-EFGH-JKLM.")
        user = re.sub(r"\s+", " ", (user or "").strip())
        if not 1 <= len(user) <= 60:
            raise ValueError("Give the licence holder's name (1 to 60 characters).")
        if not 1 <= int(days) <= self.max_days:
            raise ValueError(f"A licence lasts 1 to {self.max_days} days.")
        today = self._today()
        expires = (today + timedelta(days=int(days))).isoformat()
        body = {"v": self.version, "app": self.app_id, "for": norm, "user": user, "role": "member",
                "issued": today.isoformat(), "expires": expires}
        if features:
            body["features"] = sorted({str(f) for f in features})
        payload = json.dumps(body, separators=(",", ":"), sort_keys=True).encode()
        token = f"{self.tag}.{_b64(payload)}.{_b64(ed25519.sign(seed, payload))}"
        return {"licence": token, "expires": expires, "user": user, "for": norm}

    # ------------------------------------------------------------------ this installed copy
    def install_id(self) -> str:
        """This copy's activation code. Created once and kept in the app's own data folder."""
        code = self.store.get("install_id")
        if not (isinstance(code, str) and len(norm_code(code)) == 12):
            raw = "".join(secrets.choice(_CODE_ALPHABET) for _ in range(12))
            code = "-".join(raw[i:i + 4] for i in range(0, 12, 4))
            self.store.set("install_id", code)
        return code

    def read(self, token: str, for_install: str = None) -> dict:
        """Check signature, app, computer and expiry. Returns the licence data or raises LicenceError."""
        if not self.configured():
            raise LicenceError("This copy of the app can't check licences yet (the owner key isn't set in the build).")
        try:
            tag, payload_b64, sig_b64 = re.sub(r"\s+", "", token or "").split(".")
            payload, sig = _unb64(payload_b64), _unb64(sig_b64)
            if tag != self.tag:
                raise ValueError
            data = json.loads(payload)
            if not isinstance(data, dict):
                raise ValueError
        except Exception:
            raise LicenceError("That doesn't look like a licence. Paste the whole text the owner sent you.")
        if not ed25519.verify(self.public, payload, sig):
            raise LicenceError("This licence wasn't signed by the owner, so it can't be used.")
        if data.get("app") != self.app_id:
            raise LicenceError("This licence is for a different app.")
        if data.get("for") != norm_code(for_install or self.install_id()):
            raise LicenceError("This licence was made for a different computer. Send the owner the activation code shown on this screen.")
        try:
            expires = date.fromisoformat(data["expires"])
        except (KeyError, ValueError, TypeError):
            raise LicenceError("That licence has no valid expiry date.")
        if expires + timedelta(days=self.grace_days) < self._today():
            raise LicenceError("This licence has expired. Ask the owner for a new one.")
        return data

    def activate(self, token: str, expect_user: str = None) -> dict:
        """Verify and keep a licence (first activation, or a renewal; pass expect_user to make a renewal the same person)."""
        data = self.read(token)
        if expect_user and str(data.get("user", "")).casefold() != expect_user.strip().casefold():
            raise LicenceError("This licence is for a different person than the one who uses this computer.")
        self.store.set("license", re.sub(r"\s+", "", token))
        return data

    def claim_owner(self, key_text: str) -> None:
        """The owner proves it is them, once, on their own computer. A signed proof bound to this install is kept."""
        seed = self.parse_owner_key(key_text)
        if not self.owner_key_matches(seed):
            raise LicenceError("That owner key doesn't belong to this app, so it can't claim the owner copy.")
        self.store.set("owner_proof", _b64(ed25519.sign(seed, self._owner_msg())))

    def _owner_msg(self) -> bytes:
        return f"owner:{self.app_id}:{norm_code(self.install_id())}".encode()

    def _owner_ok(self) -> bool:
        proof = self.store.get("owner_proof")
        if not (self.public and isinstance(proof, str)):
            return False
        try:
            return ed25519.verify(self.public, self._owner_msg(), _unb64(proof))
        except Exception:
            return False

    def clock_ok(self) -> bool:
        """A computer whose clock was wound back to stretch a licence is refused."""
        high = self.store.get("clock_high", 0) or 0
        t = int(self._clock())
        if t < high - self.rollback_seconds:
            return False
        if t > high + 60:
            self.store.set("clock_high", t)
        return True

    def state(self) -> dict:
        """The one question the app asks: may this copy be used right now?"""
        if not self.configured():
            return {"ok": False, "kind": "unconfigured", "message": "This build has no owner key, so it can't be used. Set the owner public key and rebuild."}
        if self._owner_ok():
            return {"ok": True, "kind": "owner", "message": ""}
        token = self.store.get("license")
        if not token:
            return {"ok": False, "kind": "unlicensed", "message": "This computer isn't licensed yet.", "install_id": self.install_id()}
        try:
            data = self.read(token)
        except LicenceError as e:
            return {"ok": False, "kind": "unlicensed", "message": str(e), "install_id": self.install_id()}
        if not self.clock_ok():
            return {"ok": False, "kind": "unlicensed", "message": "This computer's clock was set back. Fix the date and time to continue.",
                    "install_id": self.install_id()}
        left = (date.fromisoformat(data["expires"]) - self._today()).days
        return {"ok": True, "kind": "licensed", "message": "", "expires": data["expires"], "days_left": left,
                "user": data.get("user"), "features": data.get("features", [])}
