# GST Reconciliation Tool — Handoff Notes

Picking this up in Claude Code. Read the session log first — it's the current
state. The "Baseline" section below it is the original handoff from before
this session and is mostly superseded (kept for product-decision history).

## Session Log — 2026-07-24 (v1.0.7 → v1.1.0: Credit/Debit Note reconciliation)

User provided real sample files (`sample_data/`, gitignored) showing two things
the tool didn't handle: GSTR-2B exports carry a second sheet `B2B-CDNR` (credit/
debit notes, alongside the already-handled `B2B` invoice sheet), and the client
also has separate Tally-exported **Credit Note Register** / **Debit Note
Register** files that were never uploaded anywhere. Built a full second
reconciliation pipeline for these, mirroring the existing invoice pipeline
end-to-end (parser → db tables → reconcile → report → GUI).

**Critical semantics, confirmed by user**: GSTR classifies a note from the
*supplier's* filing side; Tally classifies it from the *buyer's* books side —
the two are opposite. A Tally "Credit Note" register entry matches a supplier's
GSTR "Debit Note" filing, and vice versa. `parsers.TALLY_TO_GSTR_NOTE_TYPE` encodes
this flip; `reconcile.run_note_reconciliation()` applies it before matching
(builds a `type_key` = match_key + GSTR-equivalent-type on both sides, so a
credit note can never accidentally match against another credit note).

**Two latent bugs found and fixed while building this** (both in
`parse_purchase_register`, pre-existing, unrelated to the new feature but same
code path being touched):
1. IGST was only picked up from a column literally named `IGST` — never matched
   `IGST@18%` (rate-suffixed) or Tally's expanded ledger names like `IN: Intgrted
   GST OP 18.00 %`. CGST/SGST already summed every `CGST*`/`SGST*`-prefixed
   column but IGST didn't get the same treatment, so every interstate purchase
   silently reconciled with IGST=0. Fixed with a unified `_gst_bucket()`
   classifier (in `parsers.py`) used for all three tax heads, matching by
   normalized content (`central`/`state`/`integ`/`intgr` keywords) not just a
   fixed prefix — verified against real files, this recovered ~₹50K of
   previously-dropped IGST on one test month alone.
2. `db.unlock_or_init()` never re-ran the schema against an *existing* encrypted
   database — new tables (like the ones this session added) would silently
   never get created for a user with a pre-existing `gst_recon.db.enc`, since
   `conn.executescript(SCHEMA)` only ran on the fresh-DB branch. Fixed by
   re-running the (all `IF NOT EXISTS`) schema on every unlock, before
   `_ensure_columns()`.

### What's new

- **Parsers** (`app/parsers.py`): `parse_gstr2b_cdnr(path)` reads the `B2B-CDNR`
  sheet (same two-row merged-header pattern as `parse_gstr2a`, empty result if
  absent rather than an error — older exports won't have it). `parse_note_register(path,
  note_type)` reads a Tally Credit/Debit Note Register export (same
  letterhead+header-by-content pattern as `parse_purchase_register`, refactored
  the header-locating logic into `_locate_tally_register_header()` shared by both).
- **DB** (`app/db.py`): `gstr2b_cdnr_notes` table (FK to the *same*
  `gstr2a_snapshots` row as the B2B invoices from that upload — one file upload,
  one snapshot, two child row sets). `note_batches`/`note_entries` tables mirror
  `purchase_batches`/`purchase_entries` exactly, including the same
  conflict-detection rules (identical re-upload skipped; same voucher +
  different amount → `pending_conflict` requiring Overwrite/Ignore), scoped by
  `(client_id, note_type, match_key)` so a Credit Note and Debit Note sharing a
  voucher number are never confused.
- **Reconcile** (`app/reconcile.py`): `run_note_reconciliation(cdnr_df, notes_df)`
  — same category shape as invoice reconciliation (matched_clean,
  value_tax_mismatches, probable_matches) plus `only_in_gstr2b_cdnr`/
  `only_in_note_register`, with the Tally↔GSTR flip applied via a `type_key`
  built from each side before merging.
- **Service** (`app/service.py`): `run_fy_note_reconciliation()` mirrors
  `run_fy_reconciliation()` but treats "no note data" as normal (returns
  `meta['has_data']=False` instead of raising) since a client may simply have no
  credit/debit notes for a period. `export_fy_reconciliation()` now runs both
  and bundles everything into **one** Excel report per FY (user's explicit
  choice — one holistic file per FY, not two separate reports).
- **Report** (`app/report.py`): `build_report()` takes optional
  `note_results`/`note_meta`; when given, adds a "Credit / Debit Note
  Reconciliation" section to the Summary sheet plus `Notes_Value_Tax_Mismatches`/
  `Notes_Probable_Matches`/`Notes_Only_In_GSTR2B_CDNR`/`Notes_Only_In_NoteRegister`/
  `Notes_Matched_Clean`/`Notes_Pending_Conflicts` sheets.
- **GUI** (`app/gui.py`): Upload tab gained Step 3/4 buttons (Credit Note
  Register, Debit Note Register — separate buttons per user's explicit
  preference over one auto-detecting button, "explicit beats implicit"). GSTR-2A/2B
  upload (Step 1) now also parses+stores the `B2B-CDNR` sheet transparently, no
  new button needed there. History tab gained a combined "Credit/Debit Note
  Register Upload Batches" table (Type column distinguishes CN/DN rather than two
  separate tables). `ConflictResolutionDialog` generalized with an `entity=
  "purchase"|"note"` parameter rather than duplicating the class. New
  "Credit/Debit Notes" sub-tab under Reconciliation (own category filter/table,
  per user's explicit choice over netting into the existing ITC KPIs) — no
  Insights charts for notes in this pass (scope decision, can revisit).
  `ReportViewerWindow` (past-report viewer) does NOT show note sheets when
  reopening an old report — same acceptable-limitation pattern as the existing
  monthly-trend-chart omission, not fixed this session.

### Known limitations / things to revisit

- Past-report viewer doesn't surface Notes_* data (only reads
  `_REPORT_SHEET_ORDER`'s fixed invoice-sheet list).
- No Insights charts (KPI tiles/graphs) for credit/debit note reconciliation yet
  — only the Category Details-style table. Purely a scope call for this session,
  not a technical blocker.
- Fuzzy-match fallback for notes doesn't have multi-rate-split detection (invoice
  side's `find_multi_rate_invoices`) — not clearly applicable to notes and not
  requested, so intentionally omitted.

## Session Log — 2026-07-06 (v1.0.0 baseline → v1.0.7)

User is Sai Srinivasa Bharadwaj (BharadwajYellapragada on GitHub), the
developer/vendor of this tool, not an end client — he tests it himself and
relays what he sees on screen (often via screenshots) rather than describing
bugs abstractly. Repo: https://github.com/BharadwajYellapragada/gst_recon_tool.

### Standing workflow (established this session, keep using it)

Whenever a code change is made to this project: rebuild the PyInstaller app →
rebuild the installer → git commit + push → publish a GitHub Release with the
installer attached → give the user the download link. Do this automatically,
without asking for confirmation on each step — the user said so explicitly
("I don't want to repeat this all again whenever I want to make the change").

Exact commands:
```
".venv/Scripts/python.exe" -m PyInstaller --noconfirm GST-Reconciliation-Tool.spec
"C:\Users\bhara\AppData\Local\Programs\Inno Setup 6\ISCC.exe" installer.iss
git add <files> && git commit -m "..." && git push origin master
gh release create vX.Y.Z installer_output/GST-Reconciliation-Tool-Setup.exe installer_output/README.txt --title "..." --notes "..."
```
Gotchas:
- Calling `.venv/Scripts/pyinstaller.exe` directly fails silently (exit 1, no
  output) in this shell — always use `python -m PyInstaller` instead.
- Before rebuilding, check `tasklist | grep -i reconciliation` — if the app is
  running from `dist/`, the build fails (locked files). The user's own test
  runs are usually from `C:\Program Files\GST Reconciliation Tool\`, a
  different path that doesn't block the `dist/` rebuild.
- Inno Setup lives at `C:\Users\bhara\AppData\Local\Programs\Inno Setup 6\ISCC.exe`
  (installed via winget under the user profile, NOT in PATH, NOT in Program Files).
- `docs/index.html` (GitHub Pages) links to the stable URL
  `.../releases/latest/download/GST-Reconciliation-Tool-Setup.exe` — any new
  release just needs the asset named exactly `GST-Reconciliation-Tool-Setup.exe`
  and not marked prerelease; no docs update needed per release.
- Release tags are `vX.Y.Z`, bump patch per release (currently at v1.0.7).

### What happened, chronologically

1. **v1.0.1 — Purchase Register upload silently rejected every row.**
   `app/parsers.py::parse_purchase_register` filtered the Date column with
   `isinstance(date_val, (int, float))`. openpyxl returns `datetime.datetime`
   (not a numeric serial) for cells that are *date-formatted*, which Tally
   exports always are — so every row got skipped and the tool reported "No
   valid data rows were found." Fixed with a `_to_pydate()` helper that
   accepts both representations.

2. **v1.0.2 — Garbled/blank button text on high-DPI displays + multi-file upload.**
   The app never declared DPI awareness, so Windows bitmap-stretched the whole
   window on scaled displays. That specifically mangles `ttk.Button` text
   (rendered natively via the Windows theme engine) while leaving Tk-drawn
   labels crisp — explaining screenshots where every button showed blank/
   fragmented text but labels were fine. Fixed with `_enable_dpi_awareness()`
   (`SetProcessDpiAwareness`/`SetProcessDPIAware`, called once before any Tk
   root) + `_apply_dpi_scaling(root)` (sets `tk scaling` to match real DPI,
   called in each of the three Tk root windows: ActivationScreen, LoginScreen,
   App). Also added multi-file selection (`askopenfilenames`) to both the
   GSTR-2A/2B and Purchase Register upload buttons — continues past a single
   bad file in the batch instead of aborting the rest.

3. **v1.0.3 — GSTR-2B support + Download/Delete for uploads.**
   User uploaded a **GSTR-2B** export (not GSTR-2A) and got "Could not read
   file" with garbage in the Periods column and generation date = None.
   GSTR-2A and GSTR-2B both use a sheet named `B2B` but with *different
   column layouts* (2B has no leading Period column; adds IRN/Source columns).
   The old code assumed a fixed column position. Rewrote
   `parse_gstr2a` to locate every column by header text (scans for the header
   row anchored on "GSTIN of supplier", combines the two-row merged header,
   matches columns by keyword) instead of a hardcoded index — now handles
   both formats. Also found and fixed a second bug on the way: GSTR-2A spells
   the period `APR-22`, GSTR-2B spells it `Apr'25` — the FY-window
   reconciliation filter only recognized the first form, so GSTR-2B rows
   would silently be excluded from every reconciliation run even after a
   "successful" upload. Added `_normalize_period()` to canonicalize both to
   `MMM-YY`. Also added Download/Delete actions for GSTR-2A/2B snapshots and
   Purchase Register upload batches in the History tab (the original upload
   file itself was never stored, so Download re-exports the stored rows to a
   user-chosen path via a save dialog; Delete removes a mistaken upload and
   all its rows).

4. **v1.0.4 — Reconciliation preview showed NaN; History got row numbers.**
   `gui.py::_render_category` blindly showed `df.columns[:12]` of the raw
   outer-merge DataFrame. Since the GSTR-2A side's columns come first in the
   merge, categories dominated by Purchase-Register-only rows (e.g. "Only in
   books") showed nothing but `nan` on screen even though the underlying
   counts and the *exported Excel report* were correct all along (report.py
   already picked the right side's columns per category). Added
   `_CATEGORY_COLUMNS` mapping mirroring report.py's column choices per
   category, used by the live preview too. Also caught that
   `value_tax_mismatches` has 17 columns but the preview silently capped at
   12 — raised the cap and added a horizontal scrollbar. Added a
   Snapshot #/Batch # column to History's two tables so a row can be
   identified before Download/Delete.

5. **v1.0.5 — UI restyle.** User asked to make it "clean and beautiful" and
   stop text clipping. Added one shared style function (`_configure_style`,
   called by all three Tk roots) switching Arial → Segoe UI with consistent
   named styles (Title/Header/Section/Muted/Error/Summary), a reusable
   `_make_scrollable_tree()` (every table got vertical+horizontal scrollbars,
   replacing ad-hoc `ttk.Treeview` construction repeated 5x), and a
   content-aware `_col_width()` heuristic (supplier names/filenames/
   timestamps were clipped to a flat 120/150/180px regardless of content).

6. **v1.0.6 — Reconciliation Insights (charts) + Indian number formatting +
   maximized launch.** User asked about WinUI3/Fluent Design for a more
   "modern Windows app" look — **not achievable from Python** (WinUI3 is
   C#/.NET+XAML only); `qfluentwidgets` (PySide6/PyQt6) is the closest real
   alternative but means a full GUI rewrite. User chose to stay in
   Tkinter and just add an Analysis tab with charts (matplotlib) instead of
   attempting that migration — **this decision may be revisited later if the
   user wants to schedule the PySide6 rewrite as its own project.**
   Added an "Insights" **sub-tab inside Reconciliation** (user explicitly
   rejected a separate top-level "Analysis" tab — "having it separated is not
   suitable"): KPI tiles (Total Purchase Value, Matched Value, ITC at Risk,
   Mismatch Value) + 3 charts (category counts — categorical color, one per
   the dataviz skill's method; monthly Purchase Register totals and top
   suppliers not yet in GSTR-2A/2B — both single-hue sequential, since they're
   magnitude/ranking not identity). Charts live in a scrollable frame
   (`_make_scrollable_frame`) since 3 stacked charts don't fit one screen —
   first attempt crammed them side-by-side and the supplier chart came out
   "crunched" per user feedback; fixed by stacking full-width + scrolling.
   Also: all currency now uses **Indian digit grouping** (`₹54,25,630` not
   `₹5,425,630`) via `format_inr()`/`format_inr_short()` (lakh/crore, for
   compact chart labels/axes) — this also fixed matplotlib's ugly `1e7`
   scientific-notation axis and a value-label clipping bug on the supplier
   chart. App now launches maximized (`self.state("zoomed")`). Fixed a few
   labels (notably the reconciliation summary line) that wrapped to 2 lines
   even at full window width because of a hardcoded `wraplength` — replaced
   with `_bind_dynamic_wraplength()` which tracks the container's real width
   on `<Configure>`.

7. **v1.0.7 — Report viewer.** User wanted to reopen a past exported report
   (double-click in Past Reports) and see the same Insights/Details view, in
   a new window. Built `ReportViewerWindow(tk.Toplevel)` that reads the
   **saved .xlsx back**, not a live re-run — because Purchase Register data
   can change after a report was exported (new uploads, deletions), a live
   re-run would show different numbers than what was actually reported;
   reading the frozen file guarantees it matches exactly. Real bug caught
   here (user: "column names are not in sync check once"): several report
   sheets have a note sentence written above the real header —
   `Probable_Matches` always, `Only_In_PurchaseRegister`/`Pending_Conflicts`
   only *conditionally* (report.py calls `ws.insert_rows(1)` only when a
   `past_6_month_window` column happens to be present) — so a naive
   `pd.read_excel` (header=0 default) sometimes grabbed the note text as
   column headers and miscounted rows by one. Fixed with `_load_report_sheets()`,
   which scans the first 10 rows per sheet for the real header by content (a
   header is several short text cells; a note is one long sentence alone in
   column A) instead of assuming a fixed row. Wired to both a double-click on
   the Past Reports tree and a new "View Insights & Details..." button; the
   original "Open Selected Report" (opens the raw file in Excel) still exists
   alongside it, unchanged.

### Current known state / things to watch

- No known open bugs as of v1.0.7.
- `ReportViewerWindow`'s Insights only has 2 of the 3 live charts (category
  counts + top suppliers) — the monthly-trend chart isn't reconstructable
  from the saved file (no month-level sheet is exported), so it's correctly
  omitted rather than faked.
- Matplotlib (`matplotlib>=3.8`) is now a project dependency
  (`requirements.txt`) — bundles into the PyInstaller build fine with no
  extra hidden-imports/hooks needed.
- Test files used this session (still in `~/Downloads` as of this writing):
  `APR25.xlsx`/`MAY25.xlsx`/`JUNE25.xlsx` (Purchase Register, DEEPIKA client,
  FY2025-26), `042025_36ACQFS1575G1ZP_GSTR2B_06072026.xlsx` (a real GSTR-2B
  export — the one that surfaced the v1.0.3 bugs), and the original
  `SRINIVASA EXCLUSIVE-GSTR2A.xls`/`SRINIVASA EXCLUSIVE-GSTR2A (1).xls`
  (real GSTR-2A, 5909 B2B rows, FY2022-23 era, portal generation date
  12/12/2023) referenced in the original baseline below.

## Baseline (pre-session handoff, superseded but kept for history)

### What's built and verified

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
  `db.resolve_conflict(id, "overwrite"|"ignore")` is called. This gap is now
  closed — `ConflictResolutionDialog` in `gui.py` is the resolution screen.

- `app/parsers.py` — Reads real GSTR-2A/2B portal exports (`B2B` sheet,
  columns located by header text — see Session Log #3) and Tally-style
  Purchase Register exports (auto-detects header row, sums scattered
  `Cgst@9%`/`Cgst@6%`/etc. columns). Tags every purchase row with
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

- `app/service.py` — The orchestration layer the GUI calls.
  `run_fy_reconciliation(client_id, fy_label)` and `export_fy_reconciliation(...)`
  are the main entry points. `export_purchase_month(...)` pulls one month
  back out of a full-year upload.

- `app/report.py` — Builds the formatted multi-sheet Excel output
  (Summary, Pending_Conflicts, Value_Tax_Mismatches, Probable_Matches,
  Only_In_GSTR2A, Only_In_PurchaseRegister, Duplicates_GSTR2A,
  MultiRate_Reference, Matched_Clean). Several sheets have a note row above
  the real header — see Session Log #7 if writing anything that reads these
  sheets back.

- `app/gui.py` — Full Tkinter GUI, built out over this session (see log
  above): PIN login, client list, Upload/History/Reconciliation(+Insights
  sub-tab)/Past Reports tabs, conflict resolution dialog, report viewer.

- `vendor_tools/reset_lock.py` — Standalone script, deliberately NOT bundled
  into the end-user app/exe, for the "customer forgot their PIN" paid support
  flow.

## Product decisions already made (don't re-litigate these)

- Data storage: local machine only, no shared/network folder.
- Purchase Register uploads: monthly incremental (full-year-at-once also
  supported, auto-split by month/FY on parse). Multi-file select supported
  for both GSTR-2A/2B and Purchase Register uploads.
- PIN required every app open, in addition to the automatic machine-lock.
- No data-preserving password recovery (explicitly decided against, twice).
  Forgot-PIN = vendor-run reset = customer starts fresh, re-uploads sources.
- Late-filing window = FY + 6 months, exactly as the client described it
  (flagged once as a simplification vs actual GST law, not revisited since).
- Target platforms: Windows (primary, Tally-using accountants) AND macOS now
  in scope too. PyInstaller can't cross-compile — build natively on each OS
  (GitHub Actions `windows-latest` runner is the no-Windows-PC option for the
  .exe).
- Both GSTR-2A and GSTR-2B portal exports must be accepted (GSTR-2B is
  arguably the more commonly used one now; column layout differs from 2A).
- The uploaded source files themselves are never stored, only the parsed
  rows — Download re-exports an equivalent file from stored data, not the
  original bytes.
- Insights/Analysis belongs *inside* the Reconciliation tab (a sub-tab), not
  as its own top-level tab — explicit user preference.
- Past reports are viewed by reading the frozen exported file back, not by
  re-running reconciliation against current (possibly since-changed) data.
- Staying on Tkinter, not migrating to PySide6/qfluentwidgets for a
  WinUI3/Fluent look — a real option if revisited, but declined for now in
  favor of shipping features faster in the current stack.

## Test data used throughout

`SRINIVASA_EXCLUSIVE-GSTR2A.xls` (5,909 B2B rows) and
`SRINIVASA_EXCLUSIVE-PURCHASES-INPUT-FY22-23.xls` (5,745 rows), FY 2022-23.
Known-correct reconciliation counts to regression-test against:
matched_clean=5333–5337 (varies slightly by conflict resolution state),
value_tax_mismatches=36–40, probable_matches=196, only_in_gstr2a=236,
only_in_purchase_register=162, true_duplicates_gstr2a=0, multi_rate_reference
rows=204, purchase register conflicts=8 (6 exact dupes auto-skip + 8 real
conflicts, distinct from the mismatch count above).

Also (this session, DEEPIKA client, FY2025-26): `APR25.xlsx`/`MAY25.xlsx`/
`JUNE25.xlsx` Purchase Register + `042025_36ACQFS1575G1ZP_GSTR2B_06072026.xlsx`
GSTR-2B. Known-correct counts as of v1.0.7: matched_clean=559,
value_tax_mismatches=0, probable_matches=19, only_in_gstr2a=20,
only_in_purchase_register=1282, pending_conflicts=1.
