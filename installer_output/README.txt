GST Reconciliation Tool — Setup Instructions
=============================================

1. INSTALL
   Run GST-Reconciliation-Tool-Setup.exe and follow the wizard.

   - Windows will likely show a blue "Windows protected your PC" screen first.
     This is normal for a small unsigned tool -- click "More info", then
     "Run anyway".
   - You'll get a Windows admin permission prompt (UAC) -- accept it, the
     installer needs it to install into Program Files.
   - You can optionally check "Create a desktop shortcut" during setup.

2. ACTIVATE (first launch only)
   The first time you open the app, you'll see an "Activate" screen showing
   this computer's Machine ID (a long string starting with "win:").

   - Click "Copy" next to the Machine ID and send it to whoever gave you
     this software (email, chat, whatever's easiest).
   - They will send you back an "Activation Key".
   - Paste that key into the Activation Key box and click "Activate".

   This is a one-time step per computer. The activation key you receive only
   works on the machine whose Machine ID you sent -- it won't work if you
   move the install to a different computer (you'd need a new key for that
   machine instead).

3. FIRST-RUN PIN
   After activating, you'll be asked to set a PIN. This PIN is required every
   time you open the app, in addition to being locked to this computer. There
   is NO way to recover the PIN if it's forgotten -- forgetting it means
   starting over with a fresh, empty database (all your uploaded GSTR-2A and
   Purchase Register data would need to be re-uploaded). Choose something
   you'll remember, and consider writing it down somewhere safe.

4. USING THE APP
   - Add a client from the left panel.
   - Upload Data tab: upload the GSTR-2A file from the GST portal, and the
     Purchase Register export (monthly or a full year at once).
   - Reconciliation tab: pick the financial year and click "Run
     Reconciliation" -- it automatically uses all the GSTR-2A data you've
     uploaded for that year, no need to pick a specific upload.
   - Export the full report to Excel from the Reconciliation tab.

If anything goes wrong or you need a PIN reset, contact whoever gave you this
software -- do not try to delete or reinstall the app yourself, as that alone
will not recover a forgotten PIN (the data stays encrypted either way).
