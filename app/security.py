"""
Key management for encrypting the local database at rest.

The encryption key is derived from TWO factors combined:
  1. A password the user sets (required every time the tool opens)
  2. A fingerprint tied to this specific machine (Windows MachineGuid)

Both factors are needed to reproduce the key. Copying the encrypted database
to another machine changes factor #2, so the same password no longer derives
the same key there — the file fails to decrpyt, by design. There is
deliberately no recovery/export path (per product decision): losing the
original machine means losing access to the data.
"""
import os
import sys
import json
import base64
import hashlib
import hmac
import platform
import uuid

from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.fernet import Fernet, InvalidToken

KDF_ITERATIONS = 200_000
VERIFY_TAG = b"gst-recon-tool-verify-v1"


class SecurityError(Exception):
    """Raised for any unlock failure. Deliberately does not distinguish
    'wrong password' from 'wrong machine' — both must fail identically so
    the error message can't be used to narrow down an attack."""
    pass


def get_machine_fingerprint():
    """A stable identifier tied to this specific machine's OS installation.
    Changes on OS reinstall; does not travel with a copied file."""
    if sys.platform == "win32":
        try:
            import winreg
            for flag in (winreg.KEY_READ | winreg.KEY_WOW64_64KEY, winreg.KEY_READ):
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                          r"SOFTWARE\Microsoft\Cryptography", 0, flag)
                    guid, _ = winreg.QueryValueEx(key, "MachineGuid")
                    winreg.CloseKey(key)
                    if guid:
                        return f"win:{guid}"
                except OSError:
                    continue
        except Exception:
            pass

    if sys.platform == "darwin":
        try:
            import subprocess
            out = subprocess.check_output(
                ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"], text=True, timeout=5
            )
            for line in out.splitlines():
                if "IOPlatformUUID" in line:
                    uuid_val = line.split('"')[-2]
                    if uuid_val:
                        return f"mac:{uuid_val}"
        except Exception:
            pass

    # Linux (also covers dev/test environments without the above).
    try:
        for path in ("/etc/machine-id", "/var/lib/dbus/machine-id"):
            if os.path.exists(path):
                with open(path) as f:
                    mid = f.read().strip()
                if mid:
                    return f"linux:{mid}"
    except Exception:
        pass

    # Last resort: hostname alone. Weaker (shared across identical VM images),
    # but at least stable across runs on the same machine, unlike a random MAC.
    return f"hostname:{platform.node()}"


def _data_dir():
    from . import db
    return db.get_data_dir()


def _security_meta_path():
    return os.path.join(_data_dir(), "security.json")


def is_initialized():
    return os.path.exists(_security_meta_path())


def _derive_key(password, salt_bytes, machine_fp):
    material = (password + "|" + machine_fp).encode("utf-8")
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt_bytes, iterations=KDF_ITERATIONS)
    raw_key = kdf.derive(material)
    return base64.urlsafe_b64encode(raw_key)


def setup_new_password(password):
    """First-run setup. Returns the Fernet key to use for the new (empty) database."""
    if not password:
        raise ValueError("Password cannot be empty.")
    salt = os.urandom(16)
    machine_fp = get_machine_fingerprint()
    key = _derive_key(password, salt, machine_fp)
    verifier = hmac.new(key, VERIFY_TAG, hashlib.sha256).hexdigest()
    meta = {"salt": base64.b64encode(salt).decode(), "verifier": verifier,
             "iterations": KDF_ITERATIONS, "version": 1}
    with open(_security_meta_path(), "w") as f:
        json.dump(meta, f)
    return key


def unlock(password):
    """Returns the Fernet key if the password is correct on this machine.
    Raises SecurityError otherwise (wrong password OR wrong machine — same error)."""
    if not is_initialized():
        raise SecurityError("Security has not been set up yet.")
    with open(_security_meta_path()) as f:
        meta = json.load(f)
    salt = base64.b64decode(meta["salt"])
    machine_fp = get_machine_fingerprint()
    key = _derive_key(password, salt, machine_fp)
    expected = hmac.new(key, VERIFY_TAG, hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected, meta["verifier"]):
        raise SecurityError("Incorrect password, or this data was created on a different computer.")
    return key


def change_password(old_password, new_password):
    """Verifies the old password, then rotates to a new one. Returns the new Fernet key.
    Caller is responsible for re-encrypting the database contents with this new key."""
    unlock(old_password)  # raises SecurityError if wrong
    return setup_new_password(new_password)


def encrypt_bytes(key, raw_bytes):
    return Fernet(key).encrypt(raw_bytes)


def decrypt_bytes(key, token):
    try:
        return Fernet(key).decrypt(token)
    except InvalidToken:
        raise SecurityError("Could not decrypt the stored data with this key.")


def reset_password_lock():
    """
    VENDOR-SIDE ONLY. Not exposed anywhere in the end-user app.

    Clears the local password lock so the tool treats its next launch as a
    fresh install (mandatory PIN setup screen). This does NOT recover the
    old data — that remains encrypted with the forgotten password and is not
    decryptable by anyone, including the vendor, which is the whole point of
    the security model. The old encrypted files are archived (renamed with a
    timestamp) rather than deleted outright, in case a data-preserving
    recovery mechanism is ever built later — but with today's design they are
    just inert ciphertext.

    Intended flow: customer forgets PIN -> contacts vendor -> vendor charges
    a consultation fee -> vendor (or the customer, guided by the vendor) runs
    this function on the customer's machine -> customer re-opens the app,
    sets a new PIN, and re-uploads their source files.
    """
    import shutil
    from . import db as _db  # local import to avoid a circular import at module load time

    ts = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
    data_dir = _db.get_data_dir()
    archived = []

    meta_path = _security_meta_path()
    if os.path.exists(meta_path):
        dest = os.path.join(data_dir, f"security.json.orphaned_{ts}")
        shutil.move(meta_path, dest)
        archived.append(dest)

    enc_path = _db.ENC_DB_PATH()
    if os.path.exists(enc_path):
        dest = os.path.join(data_dir, f"gst_recon.db.enc.orphaned_{ts}")
        shutil.move(enc_path, dest)
        archived.append(dest)

    return archived
