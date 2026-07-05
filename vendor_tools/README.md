# Vendor Tools — internal use only

Everything in this folder is for the vendor's own machine. Nothing here should
ever be bundled into the installer/exe or handed to a customer.

## `vendor_private_key.pem`

The Ed25519 private key used to sign activation keys. **This file is the only
thing that can generate a valid activation key.** It is gitignored and must
never be committed, emailed, or copied off this machine.

- If it's lost: there's no recovery. You'd have to generate a brand new
  keypair (see below), update the public key baked into `app/licensing.py`,
  and rebuild the installer -- and every activation key issued under the old
  key stops working, including for customers already activated.
- Keep a backup somewhere safe and private (e.g. a password manager or
  encrypted drive) -- just not in this git repo.

If you ever need to generate a fresh keypair from scratch (e.g. initial
setup, or recovering from a lost key):

```
python -c "
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives import serialization
priv = Ed25519PrivateKey.generate()
pub = priv.public_key()
open('vendor_tools/vendor_private_key.pem', 'wb').write(priv.private_bytes(
    encoding=serialization.Encoding.PEM,
    format=serialization.PrivateFormat.PKCS8,
    encryption_algorithm=serialization.NoEncryption()))
print('PUBLIC_KEY_HEX =', pub.public_bytes(
    encoding=serialization.Encoding.Raw,
    format=serialization.PublicFormat.Raw).hex())
"
```

Then paste the printed `PUBLIC_KEY_HEX` value into `VENDOR_PUBLIC_KEY_HEX` in
`app/licensing.py` and rebuild the exe/installer.

## `generate_license.py` — issuing an activation key to a customer

When a customer installs the app on a new machine, the app's Activation
screen shows their Machine ID (a string like `win:XXXXXXXX-XXXX-...`) and
asks them to send it to you. Once you have it:

```
.venv\Scripts\python.exe vendor_tools\generate_license.py "win:XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
```

It prints an Activation Key. Send that key back to the customer -- it only
verifies against the exact fingerprint you gave it, so it won't work if they
(or anyone else) try it on a different machine.

There's no record-keeping built in -- if you want to track which keys went to
which customer/machine, keep your own log (e.g. a spreadsheet of machine ID +
key + customer name + date) when you run this.

## `reset_lock.py` — customer forgot their PIN

Standalone "forgot PIN" recovery tool, run on the customer's machine (see the
docstring at the top of the file for the full flow). Does NOT recover old
data -- it clears the lock so the customer can set a new PIN and re-upload
their source files. This is separate from activation: resetting the PIN does
not require a new activation key, since the machine itself doesn't change.

## Rebuilding the installer after a code change

```
# 1. Rebuild the onedir bundle
.venv\Scripts\python.exe -m PyInstaller --onedir --windowed --icon "assets\icon.ico" ^
    --name "GST-Reconciliation-Tool" --add-data "assets\icon.ico;assets" main.py --noconfirm

# 2. Compile the installer (adjust path if Inno Setup installed elsewhere)
"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" installer.iss
```

Output lands in `installer_output\GST-Reconciliation-Tool-Setup.exe` -- that's
the single file to send to a customer, alongside `installer_output\README.txt`.
