"""
GST Reconciliation Tool — desktop GUI (Tkinter, stdlib only).

Layout:
  Login screen: PIN entry (existing install) or PIN setup (first run).
  Main window left panel  : client list (add / search / delete)
  Main window right panel : tabs for Upload, History, Reconciliation, Past Reports
"""
import os
import sys
import json
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from . import db, parsers, reconcile, report, service, security, licensing

APP_TITLE = "GST Reconciliation Tool"


def icon_path():
    """Location of the app icon, whether running from source or from a
    PyInstaller-frozen exe (whose bundled data lands in sys._MEIPASS)."""
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.join(os.path.dirname(__file__), "..")
    path = os.path.join(base, "assets", "icon.ico")
    return path if os.path.exists(path) else None


def set_app_icon(window):
    path = icon_path()
    if path:
        try:
            window.iconbitmap(path)
        except tk.TclError:
            pass


def open_file(path):
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Could not open file", str(e))


class ActivationScreen(tk.Tk):
    """Shown before anything else, on every machine that hasn't been activated
    yet. Displays this machine's fingerprint (to send to the vendor) and takes
    the activation key the vendor issues for it -- see app/licensing.py."""

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("560x400")
        self.resizable(False, False)
        set_app_icon(self)
        self.activated = False

        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Activate " + APP_TITLE, font=("Arial", 14, "bold")).pack(anchor="w")
        ttk.Label(frame, text="This copy of the app hasn't been activated on this computer yet. "
                              "Send the machine ID below to whoever gave you this software; they'll "
                              "send back an activation key to paste in below.",
                  foreground="#595959", wraplength=500, justify="left").pack(anchor="w", pady=(6, 16))

        ttk.Label(frame, text="This computer's Machine ID").pack(anchor="w")
        fp_row = ttk.Frame(frame)
        fp_row.pack(fill="x", pady=(2, 16))
        self.fp_var = tk.StringVar(value=licensing.current_fingerprint())
        fp_entry = ttk.Entry(fp_row, textvariable=self.fp_var, state="readonly")
        fp_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(fp_row, text="Copy", command=self._copy_fingerprint).pack(side="left", padx=(6, 0))

        ttk.Label(frame, text="Activation Key").pack(anchor="w")
        self.key_var = tk.StringVar()
        key_entry = ttk.Entry(frame, textvariable=self.key_var)
        key_entry.pack(fill="x", pady=(2, 10))
        key_entry.bind("<Return>", lambda e: self._submit())
        key_entry.focus_set()

        self.error_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.error_var, foreground="#9C0006", wraplength=500,
                  justify="left").pack(anchor="w", pady=(0, 8))

        ttk.Button(frame, text="Activate", command=self._submit).pack(anchor="e")

    def _copy_fingerprint(self):
        self.clipboard_clear()
        self.clipboard_append(self.fp_var.get())

    def _submit(self):
        key = self.key_var.get().strip()
        if not key:
            self.error_var.set("Paste the activation key you were given.")
            return
        try:
            licensing.activate(key)
        except licensing.LicenseError as e:
            self.error_var.set(str(e))
            return
        self.activated = True
        self.destroy()


class LoginScreen(tk.Tk):
    """Shown before anything else. Branches to PIN setup or PIN entry
    depending on db.needs_setup(), and only unblocks the main App once
    db.unlock_or_init() succeeds."""

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("420x260")
        self.resizable(False, False)
        set_app_icon(self)
        self.unlocked = False

        self.is_first_run = db.needs_setup()

        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=APP_TITLE, font=("Arial", 14, "bold")).pack(anchor="w")

        if self.is_first_run:
            ttk.Label(frame, text="First run on this machine — set a PIN to protect your data.\n"
                                  "This PIN is required every time the app opens, and cannot be "
                                  "recovered if forgotten (see vendor support for a reset).",
                      foreground="#595959", wraplength=370, justify="left").pack(anchor="w", pady=(6, 16))
            ttk.Label(frame, text="New PIN").pack(anchor="w")
            self.pin_var = tk.StringVar()
            ttk.Entry(frame, textvariable=self.pin_var, show="*").pack(fill="x", pady=(0, 10))
            ttk.Label(frame, text="Confirm PIN").pack(anchor="w")
            self.pin2_var = tk.StringVar()
            entry2 = ttk.Entry(frame, textvariable=self.pin2_var, show="*")
            entry2.pack(fill="x", pady=(0, 10))
            entry2.bind("<Return>", lambda e: self._submit())
            btn_text = "Set PIN and continue"
        else:
            ttk.Label(frame, text="Enter your PIN to unlock this client's data.",
                      foreground="#595959", wraplength=370, justify="left").pack(anchor="w", pady=(6, 16))
            ttk.Label(frame, text="PIN").pack(anchor="w")
            self.pin_var = tk.StringVar()
            entry = ttk.Entry(frame, textvariable=self.pin_var, show="*")
            entry.pack(fill="x", pady=(0, 10))
            entry.bind("<Return>", lambda e: self._submit())
            entry.focus_set()
            btn_text = "Unlock"

        self.error_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.error_var, foreground="#9C0006").pack(anchor="w", pady=(0, 8))

        ttk.Button(frame, text=btn_text, command=self._submit).pack(anchor="e")

    def _submit(self):
        pin = self.pin_var.get()
        if not pin:
            self.error_var.set("PIN cannot be empty.")
            return
        if self.is_first_run:
            if pin != self.pin2_var.get():
                self.error_var.set("PINs do not match.")
                return
            if len(pin) < 4:
                self.error_var.set("Use at least 4 characters.")
                return
        try:
            db.unlock_or_init(pin)
        except security.SecurityError:
            self.error_var.set("Incorrect PIN, or this data was created on a different computer.")
            return
        except Exception as e:
            self.error_var.set(f"Could not unlock: {e}")
            return
        self.unlocked = True
        self.destroy()


class ConflictResolutionDialog(tk.Toplevel):
    """Shows every Purchase Register row awaiting an Overwrite/Ignore decision,
    stored value vs newly-uploaded value side by side, per HANDOFF's 'one real
    GUI gap'. Wired to db.list_pending_conflicts() / db.resolve_conflict()."""

    def __init__(self, parent, client_id, on_close=None):
        super().__init__(parent)
        self.client_id = client_id
        self.on_close = on_close
        self.title("Pending Purchase Register Conflicts")
        self.geometry("980x460")
        self.transient(parent)
        set_app_icon(self)

        ttk.Label(self, text="These invoices were re-uploaded with a DIFFERENT amount than what's "
                             "already stored. They are excluded from reconciliation until you resolve "
                             "each one — Overwrite uses the new upload, Ignore keeps the stored value.",
                  foreground="#595959", wraplength=940, justify="left", padding=10).pack(anchor="w")

        cols = ("Invoice No", "GSTIN", "Stored Gross", "New Gross", "Stored CGST", "New CGST",
                "Stored SGST", "New SGST", "Stored IGST", "New IGST")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=90, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10)

        btns = ttk.Frame(self, padding=10)
        btns.pack(fill="x")
        ttk.Button(btns, text="Overwrite Selected", command=lambda: self._resolve_selected("overwrite")).pack(side="left")
        ttk.Button(btns, text="Ignore Selected", command=lambda: self._resolve_selected("ignore")).pack(side="left", padx=6)
        ttk.Separator(btns, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(btns, text="Overwrite All", command=lambda: self._resolve_all("overwrite")).pack(side="left")
        ttk.Button(btns, text="Ignore All", command=lambda: self._resolve_all("ignore")).pack(side="left", padx=6)
        ttk.Button(btns, text="Close", command=self._close).pack(side="right")

        self._rows = []
        self._refresh()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.grab_set()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        self._rows = db.list_pending_conflicts(self.client_id)
        for item in self._rows:
            p, s = item["pending"], item["stored"]
            self.tree.insert("", "end", iid=str(p["id"]), values=(
                p["invoice_no"], p["gstin"],
                s["gross_total"] if s else "n/a", p["gross_total"],
                s["cgst"] if s else "n/a", p["cgst"],
                s["sgst"] if s else "n/a", p["sgst"],
                s["igst"] if s else "n/a", p["igst"],
            ))
        if not self._rows:
            messagebox.showinfo("No conflicts", "No pending conflicts remain.", parent=self)
            self._close()

    def _resolve_selected(self, action):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Nothing selected", "Select one or more rows first.", parent=self)
            return
        for iid in sel:
            db.resolve_conflict(int(iid), action)
        self._refresh()

    def _resolve_all(self, action):
        if not self._rows:
            return
        if not messagebox.askyesno("Confirm", f"{action.capitalize()} all {len(self._rows)} pending conflicts?", parent=self):
            return
        db.resolve_all_conflicts(self.client_id, action)
        self._refresh()

    def _close(self):
        self.grab_release()
        self.destroy()
        if self.on_close:
            self.on_close()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x720")
        self.minsize(1000, 620)
        set_app_icon(self)

        self.current_client_id = None
        self._results = None
        self._current_fy_meta = None

        self._build_layout()
        self._refresh_client_list()

    # ---------------- layout ----------------

    def _build_layout(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", rowheight=24)
        style.configure("TNotebook.Tab", padding=(14, 8))

        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)

        # ---- left: client panel ----
        left = ttk.Frame(root, width=280)
        left.pack(side="left", fill="y", padx=(0, 8))

        ttk.Label(left, text="Clients", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 4))

        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(left, textvariable=self.search_var)
        search_entry.pack(fill="x", pady=(0, 4))
        search_entry.bind("<KeyRelease>", lambda e: self._refresh_client_list())

        self.client_listbox = tk.Listbox(left, height=25, exportselection=False)
        self.client_listbox.pack(fill="both", expand=True)
        self.client_listbox.bind("<<ListboxSelect>>", self._on_client_selected)

        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="+ Add Client", command=self._add_client_dialog).pack(side="left", expand=True, fill="x", padx=(0, 4))
        ttk.Button(btns, text="Delete", command=self._delete_client).pack(side="left", expand=True, fill="x")

        # ---- right: client detail ----
        right = ttk.Frame(root)
        right.pack(side="left", fill="both", expand=True)

        header_row = ttk.Frame(right)
        header_row.pack(fill="x")
        self.header_var = tk.StringVar(value="Select or add a client to begin")
        ttk.Label(header_row, textvariable=self.header_var, font=("Arial", 13, "bold")).pack(side="left", pady=(0, 8))
        self.conflicts_btn = ttk.Button(header_row, text="Resolve Pending Conflicts",
                                         command=self._open_conflicts_dialog, state="disabled")
        self.conflicts_btn.pack(side="right")

        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill="both", expand=True)

        self.tab_upload = ttk.Frame(self.notebook, padding=12)
        self.tab_history = ttk.Frame(self.notebook, padding=12)
        self.tab_recon = ttk.Frame(self.notebook, padding=12)
        self.tab_reports = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.tab_upload, text="Upload Data")
        self.notebook.add(self.tab_history, text="History")
        self.notebook.add(self.tab_recon, text="Reconciliation")
        self.notebook.add(self.tab_reports, text="Past Reports")

        self._build_upload_tab()
        self._build_history_tab()
        self._build_recon_tab()
        self._build_reports_tab()

        self._set_tabs_enabled(False)

    def _set_tabs_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for i in range(1, 4):
            self.notebook.tab(i, state=state if enabled else "disabled")
        for w in (self.upload_g2a_btn, self.upload_pr_btn):
            w.configure(state=state)
        self.conflicts_btn.configure(state=state)

    # ---------------- client list ----------------

    def _refresh_client_list(self):
        self.client_listbox.delete(0, "end")
        query = self.search_var.get().strip().lower()
        self._clients_cache = [c for c in db.list_clients() if query in c["name"].lower()]
        for c in self._clients_cache:
            self.client_listbox.insert("end", f"{c['name']}" + (f"  ({c['gstin']})" if c["gstin"] else ""))

    def _on_client_selected(self, event):
        sel = self.client_listbox.curselection()
        if not sel:
            return
        client = self._clients_cache[sel[0]]
        self.current_client_id = client["client_id"]
        self.header_var.set(f"{client['name']}" + (f"   |   GSTIN {client['gstin']}" if client["gstin"] else ""))
        self._set_tabs_enabled(True)
        self._refresh_history_tab()
        self._refresh_recon_tab()
        self._refresh_reports_tab()

    def _add_client_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Add Client")
        dlg.geometry("360x160")
        dlg.transient(self)
        set_app_icon(dlg)
        ttk.Label(dlg, text="Client / Business Name*").pack(anchor="w", padx=12, pady=(12, 0))
        name_var = tk.StringVar()
        ttk.Entry(dlg, textvariable=name_var).pack(fill="x", padx=12)
        ttk.Label(dlg, text="GSTIN (optional)").pack(anchor="w", padx=12, pady=(8, 0))
        gstin_var = tk.StringVar()
        ttk.Entry(dlg, textvariable=gstin_var).pack(fill="x", padx=12)

        def save():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Name required", "Please enter a client name.")
                return
            try:
                db.add_client(name, gstin_var.get().strip())
            except Exception as e:
                messagebox.showerror("Could not add client", f"A client with this name may already exist.\n\n{e}")
                return
            dlg.destroy()
            self._refresh_client_list()

        ttk.Button(dlg, text="Save", command=save).pack(pady=14)
        dlg.grab_set()

    def _delete_client(self):
        sel = self.client_listbox.curselection()
        if not sel:
            return
        client = self._clients_cache[sel[0]]
        if messagebox.askyesno(
            "Delete client",
            f"This permanently deletes ALL stored data for '{client['name']}' "
            f"(GSTR-2A snapshots, purchase entries, and reconciliation history).\n\nContinue?",
        ):
            db.delete_client(client["client_id"])
            self.current_client_id = None
            self.header_var.set("Select or add a client to begin")
            self._set_tabs_enabled(False)
            self._refresh_client_list()

    # ---------------- Pending conflicts dialog ----------------

    def _open_conflicts_dialog(self):
        if not self.current_client_id:
            return
        pending = db.list_pending_conflicts(self.current_client_id)
        if not pending:
            messagebox.showinfo("No conflicts", "There are no pending Purchase Register conflicts for this client.")
            return
        ConflictResolutionDialog(self, self.current_client_id, on_close=self._on_conflicts_closed)

    def _on_conflicts_closed(self):
        self._refresh_history_tab()
        self._refresh_recon_tab()

    # ---------------- Upload tab ----------------

    def _build_upload_tab(self):
        t = self.tab_upload
        ttk.Label(t, text="Step 1 — Upload the GSTR-2A file downloaded from the GST portal",
                  font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Label(t, text="Each upload is kept as a dated snapshot — nothing is overwritten, "
                          "so you can track suppliers filing late over the following months.",
                  foreground="#595959").pack(anchor="w", pady=(0, 6))
        self.upload_g2a_btn = ttk.Button(t, text="Select GSTR-2A file (.xls/.xlsx)...",
                                          command=self._upload_gstr2a, state="disabled")
        self.upload_g2a_btn.pack(anchor="w", pady=(0, 16))

        ttk.Label(t, text="Step 2 — Upload the Purchase Register (monthly file, or a full year at once)",
                  font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Label(t, text="Upload each month's file as it becomes available, or a full-year file — rows "
                          "are auto-tagged by month/FY. Anything already stored is automatically skipped, "
                          "so it's safe to re-upload the same file by mistake. Re-uploading the same "
                          "invoice with a DIFFERENT amount is held as a pending conflict for review "
                          "(see the 'Resolve Pending Conflicts' button above).",
                  foreground="#595959", wraplength=900, justify="left").pack(anchor="w", pady=(0, 6))
        self.upload_pr_btn = ttk.Button(t, text="Select Purchase Register file (.xls/.xlsx)...",
                                         command=self._upload_purchase, state="disabled")
        self.upload_pr_btn.pack(anchor="w", pady=(0, 16))

        ttk.Separator(t).pack(fill="x", pady=8)
        self.upload_log = tk.Text(t, height=14, wrap="word", state="disabled",
                                   bg="#F7F7F7", relief="flat")
        self.upload_log.pack(fill="both", expand=True)

    def _log(self, widget, msg):
        widget.configure(state="normal")
        widget.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        widget.see("end")
        widget.configure(state="disabled")

    def _upload_gstr2a(self):
        path = filedialog.askopenfilename(
            title="Select GSTR-2A file",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            df, gen_date = parsers.parse_gstr2a(path)
            snap_id = db.add_gstr2a_snapshot(self.current_client_id, os.path.basename(path), gen_date, df)
            periods = sorted(df["Period"].unique())
            self._log(self.upload_log,
                       f"GSTR-2A loaded: {len(df)} invoice rows, periods {periods[0]}–{periods[-1]}, "
                       f"portal generation date {gen_date or 'n/a'}. Saved as snapshot #{snap_id}.")
            messagebox.showinfo("GSTR-2A uploaded",
                                 f"{len(df)} invoice rows saved as a new snapshot.\n\n"
                                 f"Periods covered: {periods[0]} to {periods[-1]}")
        except parsers.ParseError as e:
            messagebox.showerror("Could not read file", str(e))
        except Exception as e:
            messagebox.showerror("Unexpected error", f"{e}\n\n{traceback.format_exc()[-800:]}")
            return
        self._refresh_history_tab()
        self._refresh_recon_tab()

    def _upload_purchase(self):
        path = filedialog.askopenfilename(
            title="Select Purchase Register file",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            entries = parsers.parse_purchase_register(path)
            result = db.add_purchase_batch(self.current_client_id, os.path.basename(path), entries)
            self._log(self.upload_log,
                       f"Purchase Register file processed: {result['total_in_file']} rows read → "
                       f"{result['new_rows']} new, {result['duplicate_rows_skipped']} already stored (skipped), "
                       f"{result['pending_conflicts']} flagged as amount conflicts for review.")
            msg = (f"{result['new_rows']} new entries added.\n"
                   f"{result['duplicate_rows_skipped']} rows were already in the system and skipped.\n")
            if result["pending_conflicts"]:
                msg += (f"\n{result['pending_conflicts']} entries have the same invoice number as one already "
                        f"stored but a DIFFERENT amount. Use 'Resolve Pending Conflicts' above before "
                        f"relying on totals — these are excluded from reconciliation until resolved.")
            messagebox.showinfo("Purchase Register uploaded", msg)
        except parsers.ParseError as e:
            messagebox.showerror("Could not read file", str(e))
        except Exception as e:
            messagebox.showerror("Unexpected error", f"{e}\n\n{traceback.format_exc()[-800:]}")
            return
        self._refresh_history_tab()
        self._refresh_recon_tab()

    # ---------------- History tab ----------------

    def _build_history_tab(self):
        t = self.tab_history
        ttk.Label(t, text="GSTR-2A Snapshots", font=("Arial", 10, "bold")).pack(anchor="w")
        cols = ("Uploaded", "File", "Portal Generation Date", "Periods", "Invoices")
        self.snap_tree = ttk.Treeview(t, columns=cols, show="headings", height=8)
        for c in cols:
            self.snap_tree.heading(c, text=c)
            self.snap_tree.column(c, width=180 if c != "Invoices" else 80, anchor="w")
        self.snap_tree.pack(fill="x", pady=(4, 16))

        ttk.Label(t, text="Purchase Register Upload Batches", font=("Arial", 10, "bold")).pack(anchor="w")
        cols2 = ("Uploaded", "File", "Rows in File", "New", "Skipped (dup)", "Pending Conflicts")
        self.batch_tree = ttk.Treeview(t, columns=cols2, show="headings", height=10)
        for c in cols2:
            self.batch_tree.heading(c, text=c)
            self.batch_tree.column(c, width=150, anchor="w")
        self.batch_tree.pack(fill="both", expand=True, pady=(4, 0))

    def _refresh_history_tab(self):
        if not self.current_client_id:
            return
        self.snap_tree.delete(*self.snap_tree.get_children())
        for s in db.list_snapshots(self.current_client_id):
            self.snap_tree.insert("", "end", values=(
                s["uploaded_at"], s["source_filename"], s["generation_date"] or "n/a",
                f"{s['period_min']}–{s['period_max']}", s["row_count"],
            ))
        self.batch_tree.delete(*self.batch_tree.get_children())
        for b in db.list_purchase_batches(self.current_client_id):
            self.batch_tree.insert("", "end", values=(
                b["uploaded_at"], b["source_filename"], b["row_count"], b["new_rows"],
                b["duplicate_rows_skipped"], b["conflict_rows"],
            ))

    # ---------------- Reconciliation tab ----------------

    def _build_recon_tab(self):
        t = self.tab_recon
        top = ttk.Frame(t)
        top.pack(fill="x", pady=(0, 4))
        ttk.Label(top, text="Financial Year:").pack(side="left")
        self.fy_var = tk.StringVar()
        self.fy_combo = ttk.Combobox(top, textvariable=self.fy_var, width=14, state="readonly")
        self.fy_combo.pack(side="left", padx=8)
        ttk.Button(top, text="Run Reconciliation", command=self._run_reconciliation).pack(side="left", padx=8)
        ttk.Label(top, text="GSTR-2A data used is automatically merged across every snapshot "
                            "uploaded for this client — no manual selection needed.",
                  foreground="#595959").pack(side="left", padx=(16, 0))

        month_row = ttk.Frame(t)
        month_row.pack(fill="x", pady=(0, 10))
        ttk.Label(month_row, text="Export one month of Purchase Register data:").pack(side="left")
        self.month_var = tk.StringVar()
        self.month_combo = ttk.Combobox(month_row, textvariable=self.month_var, width=14, state="readonly")
        self.month_combo.pack(side="left", padx=8)
        ttk.Button(month_row, text="Export Month to Excel...", command=self._export_month).pack(side="left")

        self.recon_summary_var = tk.StringVar(value="")
        ttk.Label(t, textvariable=self.recon_summary_var, foreground="#1F4E78",
                  font=("Arial", 10, "bold"), wraplength=1000, justify="left").pack(anchor="w", pady=(0, 2))
        self.recon_cutoff_var = tk.StringVar(value="")
        ttk.Label(t, textvariable=self.recon_cutoff_var, font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 6))

        filt = ttk.Frame(t)
        filt.pack(fill="x", pady=(0, 4))
        ttk.Label(filt, text="View category:").pack(side="left")
        self.category_var = tk.StringVar(value="value_tax_mismatches")
        self.category_combo = ttk.Combobox(filt, textvariable=self.category_var, state="readonly", width=45, values=[
            "value_tax_mismatches", "probable_matches", "only_in_gstr2a",
            "only_in_purchase_register", "matched_clean", "multi_rate_reference",
        ])
        self.category_combo.pack(side="left", padx=8)
        self.category_combo.bind("<<ComboboxSelected>>", lambda e: self._render_category())

        self.result_tree = ttk.Treeview(t, show="headings")
        self.result_tree.pack(fill="both", expand=True, pady=(4, 8))

        btns = ttk.Frame(t)
        btns.pack(fill="x")
        ttk.Button(btns, text="Export Full Report to Excel...", command=self._export_report).pack(side="left")

        self._results = None

    def _refresh_recon_tab(self):
        if not self.current_client_id:
            return
        fys = service.list_available_fys(self.current_client_id)
        self.fy_combo.configure(values=fys)
        if fys:
            self.fy_combo.current(0)
        else:
            self.fy_var.set("")

        months = db.list_available_months(self.current_client_id)
        self.month_combo.configure(values=months)
        if months:
            self.month_combo.current(0)
        else:
            self.month_var.set("")

        self._results = None
        self._current_fy_meta = None
        self.result_tree.delete(*self.result_tree.get_children())
        self.recon_summary_var.set("")
        self.recon_cutoff_var.set("")

    def _run_reconciliation(self):
        if not self.fy_var.get():
            messagebox.showwarning("No financial year", "Upload at least one Purchase Register file first.")
            return
        try:
            outcome = service.run_fy_reconciliation(self.current_client_id, self.fy_var.get())
        except ValueError as e:
            messagebox.showwarning("Cannot reconcile", str(e))
            return
        except Exception as e:
            messagebox.showerror("Reconciliation failed", f"{e}\n\n{traceback.format_exc()[-800:]}")
            return

        self._results = outcome["results"]
        self._current_fy_meta = outcome["meta"]
        r = self._results
        self.recon_summary_var.set(
            f"FY {self._current_fy_meta['fy_label']}  |  Matched clean: {len(r['matched_clean'])}   |   "
            f"Mismatches: {len(r['value_tax_mismatches'])}   |   Probable matches: {len(r['probable_matches'])}   |   "
            f"Only in GSTR-2A: {len(r['only_in_gstr2a'])}   |   Only in books: {len(r['only_in_purchase_register'])}   |   "
            f"Pending conflicts (excluded): {self._current_fy_meta['pending_conflicts_count']}"
        )
        days = self._current_fy_meta["days_until_cutoff"]
        snap_word = "snapshot" if self._current_fy_meta["gstr2a_snapshot_count"] == 1 else "snapshots"
        cutoff_text = (f"({days} days remaining)" if days >= 0 else f"({abs(days)} days PAST cutoff)")
        self.recon_cutoff_var.set(
            f"Late-filing cutoff: {self._current_fy_meta['cutoff_date']}  {cutoff_text}   |   "
            f"GSTR-2A data merged from {self._current_fy_meta['gstr2a_snapshot_count']} {snap_word} "
            f"(most recent uploaded {self._current_fy_meta['gstr2a_latest_uploaded_at']})"
        )
        self._render_category()

    def _render_category(self):
        self.result_tree.delete(*self.result_tree.get_children())
        self.result_tree["columns"] = ()
        if not self._results:
            return
        cat = self.category_var.get()
        df = self._results.get(cat)
        if df is None or len(df) == 0:
            self.result_tree["columns"] = ("info",)
            self.result_tree.heading("info", text="Result")
            self.result_tree.insert("", "end", values=("No rows in this category.",))
            return
        display_cols = list(df.columns)[:12]
        self.result_tree["columns"] = display_cols
        for c in display_cols:
            self.result_tree.heading(c, text=c)
            self.result_tree.column(c, width=120, anchor="w")
        for _, row in df.head(500).iterrows():
            self.result_tree.insert("", "end", values=[row[c] for c in display_cols])

    def _export_report(self):
        if not self._results:
            messagebox.showwarning("Run reconciliation first", "Click 'Run Reconciliation' before exporting.")
            return
        client = db.get_client(self.current_client_id)
        fy_label = self._current_fy_meta["fy_label"]
        default_name = f"{client['name'].replace(' ', '_')}_GST_Reconciliation_FY{fy_label}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save reconciliation report", initialfile=default_name,
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            service.export_fy_reconciliation(self.current_client_id, fy_label, path)
        except Exception as e:
            messagebox.showerror("Export failed", f"{e}\n\n{traceback.format_exc()[-800:]}")
            return
        self._refresh_reports_tab()
        if messagebox.askyesno("Report saved", f"Report saved to:\n{path}\n\nOpen it now?"):
            open_file(path)

    def _export_month(self):
        if not self.month_var.get():
            messagebox.showwarning("No month available", "Upload a Purchase Register file first.")
            return
        client = db.get_client(self.current_client_id)
        year_month = self.month_var.get()
        default_name = f"{client['name'].replace(' ', '_')}_Purchases_{year_month}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save month export", initialfile=default_name,
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            service.export_purchase_month(self.current_client_id, year_month, path)
        except Exception as e:
            messagebox.showerror("Export failed", f"{e}\n\n{traceback.format_exc()[-800:]}")
            return
        if messagebox.askyesno("Export saved", f"Saved to:\n{path}\n\nOpen it now?"):
            open_file(path)

    # ---------------- Reports tab ----------------

    def _build_reports_tab(self):
        t = self.tab_reports
        ttk.Label(t, text="Past reconciliation reports for this client", font=("Arial", 10, "bold")).pack(anchor="w")
        cols = ("Run Date", "FY", "GSTR-2A Snapshot", "Purchase Data As Of", "File")
        self.reports_tree = ttk.Treeview(t, columns=cols, show="headings")
        for c in cols:
            self.reports_tree.heading(c, text=c)
            self.reports_tree.column(c, width=180, anchor="w")
        self.reports_tree.pack(fill="both", expand=True, pady=(4, 8))
        ttk.Button(t, text="Open Selected Report", command=self._open_selected_report).pack(anchor="w")

    def _refresh_reports_tab(self):
        if not self.current_client_id:
            return
        self.reports_tree.delete(*self.reports_tree.get_children())
        for r in db.list_recon_runs(self.current_client_id):
            self.reports_tree.insert("", "end", values=(
                r["run_at"], r["fy_label"] or "n/a", f"#{r['snapshot_id']}", r["purchase_asof"], r["report_path"],
            ))

    def _open_selected_report(self):
        sel = self.reports_tree.selection()
        if not sel:
            return
        path = self.reports_tree.item(sel[0])["values"][4]
        if os.path.exists(path):
            open_file(path)
        else:
            messagebox.showerror("File not found", f"This report file no longer exists at:\n{path}")


def main():
    if not licensing.is_activated():
        activation = ActivationScreen()
        activation.mainloop()
        if not activation.activated:
            return

    login = LoginScreen()
    login.mainloop()
    if not login.unlocked:
        return
    app = App()
    app.mainloop()
