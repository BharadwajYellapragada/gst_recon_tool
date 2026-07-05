"""
Per-machine activation lock.

Ed25519 signature scheme: the vendor holds a PRIVATE key (vendor_tools/
generate_license.py, run only on the vendor's own machine -- never shipped).
This app only embeds the matching PUBLIC key below, which can verify a
signature but cannot be used to create one. An activation key is simply an
Ed25519 signature (by the vendor's private key) over this machine's
fingerprint string (see security.get_machine_fingerprint()).

Copying the installed app (or just the .exe) to a different machine changes
the fingerprint, so a signature issued for one machine will not verify on
another -- that machine has to go through activation again, which means
sending its (new) fingerprint to the vendor for a new key.

This is a per-machine speed bump for a small, vendor-distributed tool, not a
defense against a reverse engineer willing to patch the binary -- there is no
purely client-side scheme that can prevent that.
"""
import base64
import json
import os

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
from cryptography.exceptions import InvalidSignature

from . import security

VENDOR_PUBLIC_KEY_HEX = "e684386e090d697bae0b2dff9cbbbb37f1eea3b4a4e4f20efcf4093c6eb7be72"


class LicenseError(Exception):
    pass


def _license_path():
    from . import db
    return os.path.join(db.get_data_dir(), "license.json")


def _clean_key(activation_key):
    # NOTE: '-' and '_' are meaningful characters in base64 *urlsafe* encoding
    # (replacing '+' and '/'), not cosmetic separators -- only whitespace is
    # safe to strip here.
    return "".join(activation_key.split())


def _verify(activation_key, fingerprint):
    cleaned = _clean_key(activation_key)
    try:
        sig = base64.urlsafe_b64decode(cleaned + "=" * (-len(cleaned) % 4))
    except Exception:
        return False
    pub = Ed25519PublicKey.from_public_bytes(bytes.fromhex(VENDOR_PUBLIC_KEY_HEX))
    try:
        pub.verify(sig, fingerprint.encode("utf-8"))
        return True
    except InvalidSignature:
        return False


def current_fingerprint():
    return security.get_machine_fingerprint()


def is_activated():
    """Re-derives this machine's live fingerprint and checks the stored
    activation key against it -- a stored key only ever validates on the
    exact machine it was issued for."""
    path = _license_path()
    if not os.path.exists(path):
        return False
    try:
        with open(path) as f:
            data = json.load(f)
    except Exception:
        return False
    return _verify(data.get("activation_key", ""), current_fingerprint())


def activate(activation_key):
    """Raises LicenseError if the key doesn't verify against this machine's
    fingerprint. Otherwise stores it and returns True."""
    if not _verify(activation_key, current_fingerprint()):
        raise LicenseError(
            "This activation key is not valid for this computer. "
            "Double-check it was copied in full, or request a new key for this machine."
        )
    with open(_license_path(), "w") as f:
        json.dump({"activation_key": _clean_key(activation_key)}, f)
    return True
