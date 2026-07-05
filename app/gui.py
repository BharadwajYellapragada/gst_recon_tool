"""
GST Reconciliation Tool — desktop GUI (Tkinter, stdlib only).

Layout:
  Left panel  : client list (add / search / delete)
  Right panel : tabs for Upload, History, Reconciliation, Past Reports
"""
import os
import sys
import json
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from . import db, parsers, reconcile, report

APP_TITLE = "GST Reconciliation Tool"


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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x720")
        self.minsize(1000, 620)

        db.init_db()
        self.current_client_id = None

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

        self.header_var = tk.StringVar(value="Select or add a client to begin")
        ttk.Label(right, textvariable=self.header_var, font=("Arial", 13, "bold")).pack(anchor="w", pady=(0, 8))

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

        ttk.Label(t, text="Step 2 — Upload the Purchase Register (monthly file)",
                  font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Label(t, text="Upload each month's file as it becomes available. New invoices are "
                          "added; anything already stored is automatically skipped, so it's safe "
                          "to re-upload the same file by mistake.",
                  foreground="#595959").pack(anchor="w", pady=(0, 6))
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
                       f"GSTR-2A loaded: {len(df)} invoice rows, periods {periods[0]}\u2013{periods[-1]}, "
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
                       f"Purchase Register file processed: {result['total_in_file']} rows read \u2192 "
                       f"{result['new_rows']} new, {result['duplicate_rows_skipped']} already stored (skipped), "
                       f"{result['conflict_rows']} flagged as amount conflicts for review.")
            msg = (f"{result['new_rows']} new entries added.\n"
                   f"{result['duplicate_rows_skipped']} rows were already in the system and skipped.\n")
            if result["conflict_rows"]:
                msg += (f"\n{result['conflict_rows']} entries have the same invoice number as one already "
                        f"stored but a DIFFERENT amount — both were kept. Review these in the "
                        f"Reconciliation tab's Purchase Register conflicts before relying on totals.")
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
        cols2 = ("Uploaded", "File", "Rows in File", "New", "Skipped (dup)", "Conflicts")
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
                f"{s['period_min']}\u2013{s['period_max']}", s["row_count"],
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
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="GSTR-2A snapshot to reconcile against:").pack(side="left")
        self.snapshot_var = tk.StringVar()
        self.snapshot_combo = ttk.Combobox(top, textvariable=self.snapshot_var, width=60, state="readonly")
        self.snapshot_combo.pack(side="left", padx=8)
        ttk.Button(top, text="Run Reconciliation", command=self._run_reconciliation).pack(side="left", padx=8)

        self.recon_summary_var = tk.StringVar(value="")
        ttk.Label(t, textvariable=self.recon_summary_var, foreground="#1F4E78",
                  font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 6))

        filt = ttk.Frame(t)
        filt.pack(fill="x", pady=(0, 4))
        ttk.Label(filt, text="View category:").pack(side="left")
        self.category_var = tk.StringVar(value="value_tax_mismatches")
        self.category_combo = ttk.Combobox(filt, textvariable=self.category_var, state="readonly", width=45, values=[
            "value_tax_mismatches", "probable_matches", "only_in_gstr2a",
            "only_in_purchase_register", "purchase_register_conflicts", "matched_clean",
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
        snaps = db.list_snapshots(self.current_client_id)
        self._snap_lookup = {}
        values = []
        for s in snaps:
            label = f"#{s['snapshot_id']}  uploaded {s['uploaded_at']}  ({s['period_min']}\u2013{s['period_max']}, {s['row_count']} rows)"
            values.append(label)
            self._snap_lookup[label] = s["snapshot_id"]
        self.snapshot_combo.configure(values=values)
        if values:
            self.snapshot_combo.current(0)
        self._results = None
        self.result_tree.delete(*self.result_tree.get_children())
        self.recon_summary_var.set("")

    def _run_reconciliation(self):
        if not self.snapshot_var.get():
            messagebox.showwarning("No snapshot", "Upload a GSTR-2A file first (Upload Data tab).")
            return
        snap_id = self._snap_lookup[self.snapshot_var.get()]
        g2a_df = db.get_snapshot_invoices(snap_id)
        pr_df = db.get_all_purchase_entries(self.current_client_id)
        if len(pr_df) == 0:
            messagebox.showwarning("No purchase data", "Upload at least one Purchase Register file first.")
            return
        try:
            self._results = reconcile.run_reconciliation(g2a_df, pr_df)
            self._results["purchase_register_conflicts"] = db.get_conflict_entries(self.current_client_id)
            self._current_snapshot_id = snap_id
        except Exception as e:
            messagebox.showerror("Reconciliation failed", f"{e}\n\n{traceback.format_exc()[-800:]}")
            return

        r = self._results
        self.recon_summary_var.set(
            f"Matched clean: {len(r['matched_clean'])}   |   Mismatches: {len(r['value_tax_mismatches'])}   |   "
            f"Probable matches: {len(r['probable_matches'])}   |   Only in GSTR-2A: {len(r['only_in_gstr2a'])}   |   "
            f"Only in books: {len(r['only_in_purchase_register'])}   |   PR conflicts: {len(r['purchase_register_conflicts'])}"
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
        default_name = f"{client['name'].replace(' ', '_')}_GST_Reconciliation_{datetime.now().strftime('%Y%m%d')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save reconciliation report", initialfile=default_name,
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        snap = next(s for s in db.list_snapshots(self.current_client_id) if s["snapshot_id"] == self._current_snapshot_id)
        pr_df = db.get_all_purchase_entries(self.current_client_id)
        purchase_asof = pr_df["entry_date"].max() if len(pr_df) else ""
        try:
            report.build_report(client["name"], client["gstin"], dict(snap), purchase_asof, self._results, path)
            summary = {k: len(v) for k, v in self._results.items()}
            db.log_recon_run(self.current_client_id, self._current_snapshot_id, purchase_asof, path, json.dumps(summary))
        except Exception as e:
            messagebox.showerror("Export failed", f"{e}\n\n{traceback.format_exc()[-800:]}")
            return
        self._refresh_reports_tab()
        if messagebox.askyesno("Report saved", f"Report saved to:\n{path}\n\nOpen it now?"):
            open_file(path)

    # ---------------- Reports tab ----------------

    def _build_reports_tab(self):
        t = self.tab_reports
        ttk.Label(t, text="Past reconciliation reports for this client", font=("Arial", 10, "bold")).pack(anchor="w")
        cols = ("Run Date", "GSTR-2A Snapshot", "Purchase Data As Of", "File")
        self.reports_tree = ttk.Treeview(t, columns=cols, show="headings")
        for c in cols:
            self.reports_tree.heading(c, text=c)
            self.reports_tree.column(c, width=200, anchor="w")
        self.reports_tree.pack(fill="both", expand=True, pady=(4, 8))
        ttk.Button(t, text="Open Selected Report", command=self._open_selected_report).pack(anchor="w")

    def _refresh_reports_tab(self):
        if not self.current_client_id:
            return
        self.reports_tree.delete(*self.reports_tree.get_children())
        for r in db.list_recon_runs(self.current_client_id):
            self.reports_tree.insert("", "end", values=(r["run_at"], f"#{r['snapshot_id']}", r["purchase_asof"], r["report_path"]))

    def _open_selected_report(self):
        sel = self.reports_tree.selection()
        if not sel:
            return
        path = self.reports_tree.item(sel[0])["values"][3]
        if os.path.exists(path):
            open_file(path)
        else:
            messagebox.showerror("File not found", f"This report file no longer exists at:\n{path}")


def main():
    app = App()
    app.mainloop()
