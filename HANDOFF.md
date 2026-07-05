# GST Reconciliation Tool — Handoff Notes

Picking this up in Claude Code. Everything below reflects the actual state of
the code as of this handoff — verified by running it, not just written.

## What's built and verified (all logic, no GUI wiring yet)

- `app/security.py` — Password + machine-fingerprint key derivation
  (PBKDF2-HMAC-SHA256, 200k iterations). Windows uses registry `MachineGuid`,
  macOS uses `IOPlatformUUID` (via `ioreg`), Linux uses `/etc/machine-id`,
  with a hostname-only last resort. `reset_password_lock()` is the
  vendor-side "forgot PIN" reset — archives (doesn't delete) old files and
  forces a fresh first-run setup. Does NOT recover old data by design.

- `app/db.py` — SQLite database that lives ONLY in memory during a session
  (`sqlite3.connect(":memory:")` + `.serialize()`/`.deserialize()`). On disk
  it exists only as `gst_recon.db.enc`, a Fernet-encrypted blob, rewritten
  atomically after every write via `persist()`. Call `db.unlock_or_init(pin)`
  once at startup before any other db.* call — raises `security.SecurityError`
  on wrong PIN or wrong machine (deliberately the same error for both).
  `db.needs_setup()` tells you whether to show "set a new PIN" vs "enter PIN"
  on the login screen.

  Purchase Register conflict handling: re-uploading an identical row is
  auto-skipped. Re-uploading the same invoice with a DIFFERENT amount creates
  a `status='pending_conflict'` row, excluded from all reconciliation until
  `db.resolve_conflict(id, "overwrite"|"ignore")` is called. This is the
  **one real GUI gap** — there's no dialog yet for showing pending conflicts
  and collecting the user's choice. `db.list_pending_conflicts(client_id)`
  returns everything needed to build that screen (stored value vs new value,
  side by side).

- `app/parsers.py` — Reads real GSTR-2A portal exports (`B2B` sheet) and
  Tally-style Purchase Register exports (auto-detects header row, sums
  scattered `Cgst@9%`/`Cgst@6%`/etc. columns). Tags every purchase row with
  `entry_fy` / `entry_month` at parse time.

- `app/reconcile.py` — Matches GSTIN + normalized invoice number, aggregating
  multi-rate line-splits before comparing (verified: real multi-rate
  invoices reconcile correctly). Categories: matched_clean,
  value_tax_mismatches, probable_matches (fuzzy GSTIN+date+amount fallback
  when invoice numbers don't match), only_in_gstr2a, only_in_purchase_register,
  true_duplicates_gstr2a (should always be 0 — multi-rate splits are NOT
  duplicates), multi_rate_reference.

- `app/fy_utils.py` — Indian FY math (Apr–Mar). `fy_period_window(fy_label,
  extra_months=6)` returns the GSTR-2A period window to check against
  (FY + 6 months late-filing allowance, per client's stated business rule —
  NOT the actual Section 16(4) legal cutoff, flagged to the client once).

- `app/service.py` — The orchestration layer a GUI should call.
  `run_fy_reconciliation(client_id, fy_label)` and `export_fy_reconciliation(...)`
  are the main entry points. `export_purchase_month(...)` pulls one month
  back out of a full-year upload.

- `app/report.py` — Builds the formatted multi-sheet Excel output
  (Summary, Pending_Conflicts, Value_Tax_Mismatches, Probable_Matches,
  Only_In_GSTR2A, Only_In_PurchaseRegister, Duplicates_GSTR2A,
  MultiRate_Reference, Matched_Clean).

- `vendor_tools/reset_lock.py` — Standalone script, deliberately NOT bundled
  into the end-user app/exe, for the "customer forgot their PIN" paid support
  flow.

## The one thing that needs a real rebuild, not a patch

`app/gui.py` (~460 lines) was written BEFORE the encryption layer and the
conflict-resolution rework. It still calls old function signatures
(`db.init_db()`, old `log_recon_run()` arity) and has no PIN entry screen at
all. Treat it as a rough sketch of the intended layout (client list on the
left, tabs for Upload/History/Reconciliation/Reports on the right), not
working code. It needs:
1. A PIN entry/setup screen before the main window shows anything
   (`db.needs_setup()` branches to "set a new PIN" vs "enter PIN"; catch
   `security.SecurityError` on wrong PIN and let them retry).
2. A pending-conflicts resolution dialog wired to
   `db.list_pending_conflicts()` / `db.resolve_conflict()`.
3. FY selection wired to `service.list_available_fys()` /
   `service.run_fy_reconciliation()` (the current gui.py only knows about
   snapshot-level reconciliation, predating the FY-scoping work).
4. Month-wise purchase export wired to `service.export_purchase_month()`.

## Product decisions already made (don't re-litigate these)

- Data storage: local machine only, no shared/network folder.
- Purchase Register uploads: monthly incremental (full-year-at-once also
  supported, auto-split by month/FY on parse).
- PIN required every app open, in addition to the automatic machine-lock.
- No data-preserving password recovery (explicitly decided against, twice).
  Forgot-PIN = vendor-run reset = customer starts fresh, re-uploads sources.
- Late-filing window = FY + 6 months, exactly as the client described it
  (flagged once as a simplification vs actual GST law, not revisited since).
- Target platforms: Windows (primary, Tally-using accountants) AND macOS now
  in scope too. PyInstaller can't cross-compile — build natively on each OS
  (GitHub Actions `windows-latest` runner is the no-Windows-PC option for the
  .exe).

## Test data used throughout

`SRINIVASA_EXCLUSIVE-GSTR2A.xls` (5,909 B2B rows) and
`SRINIVASA_EXCLUSIVE-PURCHASES-INPUT-FY22-23.xls` (5,745 rows), FY 2022-23.
Known-correct reconciliation counts to regression-test against:
matched_clean=5333–5337 (varies slightly by conflict resolution state),
value_tax_mismatches=36–40, probable_matches=196, only_in_gstr2a=236,
only_in_purchase_register=162, true_duplicates_gstr2a=0, multi_rate_reference
rows=204, purchase register conflicts=8 (6 exact dupes auto-skip + 8 real
conflicts, distinct from the mismatch count above).
