"""
VENDOR-ONLY UTILITY. Do not bundle this or vendor_private_key.pem into the
end-user app/exe.

Run this on the vendor's own machine to issue an activation key for a
customer's machine fingerprint (shown on the app's Activation screen when it
isn't activated yet -- ask the customer to read/send you that string).

Usage:
    python vendor_tools/generate_license.py "win:XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
"""
import sys
import os
import base64

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

PRIVATE_KEY_PATH = os.path.join(os.path.dirname(__file__), "vendor_private_key.pem")


def load_private_key():
    if not os.path.exists(PRIVATE_KEY_PATH):
        sys.exit(f"Private key not found at {PRIVATE_KEY_PATH}. This key must never leave "
                 f"the vendor's machine -- if it's missing, it needs to be restored from backup, "
                 f"not regenerated (regenerating invalidates every activation key issued so far).")
    with open(PRIVATE_KEY_PATH, "rb") as f:
        return serialization.load_pem_private_key(f.read(), password=None)


def generate_key(fingerprint: str) -> str:
    priv: Ed25519PrivateKey = load_private_key()
    sig = priv.sign(fingerprint.encode("utf-8"))
    return base64.urlsafe_b64encode(sig).decode("ascii").rstrip("=")


def main():
    if len(sys.argv) != 2:
        print("Usage: python generate_license.py \"<machine fingerprint string>\"")
        sys.exit(1)
    fingerprint = sys.argv[1].strip()
    key = generate_key(fingerprint)
    print()
    print(f"Machine fingerprint : {fingerprint}")
    print(f"Activation key      : {key}")
    print()
    print("Send the activation key above to the customer -- it only works for the")
    print("fingerprint shown above.")


if __name__ == "__main__":
    main()
