"""
VENDOR-ONLY UTILITY. Do not bundle this into the end-user app or .exe.

Run this on a customer's machine (after they've paid the recovery/consultation
fee) to clear a forgotten PIN. It does NOT recover their old data — that stays
encrypted and unreadable by design. It clears the local lock so the app treats
the next launch as a fresh install: the customer sets a brand new PIN and
re-uploads their source GSTR-2A / Purchase Register files.

Usage (on the customer's machine, in a terminal):
    python reset_lock.py

It will ask for a final confirmation before touching anything, and it archives
(renames) the old files rather than deleting them outright.
"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from app import security, db


def main():
    print("GST Reconciliation Tool — PIN Reset (vendor use only)")
    print("=" * 60)
    data_dir = db.get_data_dir()
    print(f"Data folder on this machine: {data_dir}")

    if not security.is_initialized():
        print("No PIN is currently set on this machine — nothing to reset.")
        return

    print()
    print("This will lock the customer out of their CURRENT data permanently")
    print("(it was already unrecoverable without the forgotten PIN — this just")
    print("clears the lock so they can start fresh). Old files are archived,")
    print("not deleted, with a timestamp.")
    confirm = input("Type RESET to proceed: ").strip()
    if confirm != "RESET":
        print("Cancelled — nothing was changed.")
        return

    archived = security.reset_password_lock()
    print()
    print("Done. Archived files:")
    for path in archived:
        print(f"  - {path}")
    print()
    print("The customer can now reopen the app and will be prompted to set a new PIN.")


if __name__ == "__main__":
    main()
