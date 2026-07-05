"""
Orchestration layer sitting above db / parsers / reconcile / report.
This is what a GUI (or CLI, or tests) should call — it encodes the
FY-scoping and 6-month late-filing rules described by the client.
"""
import datetime
import pandas as pd

from . import db, reconcile, report, fy_utils

DEFAULT_LATE_FILING_MONTHS = 6


def list_available_fys(client_id):
    """FYs the user can select, driven by what Purchase Register data has been uploaded."""
    return db.list_available_fys(client_id)


def get_latest_snapshot(client_id):
    snaps = db.list_snapshots(client_id)
    return snaps[0] if snaps else None


def run_fy_reconciliation(client_id, fy_label, snapshot_id=None, late_filing_months=DEFAULT_LATE_FILING_MONTHS):
    """
    Reconciles the selected financial year's Purchase Register against GSTR-2A
    data for that FY plus `late_filing_months` beyond its end.

    Returns a dict:
      results        -> the category DataFrames from reconcile.run_reconciliation(),
                         with 'only_in_purchase_register' additionally carrying a
                         'past_6_month_window' boolean column
      meta            -> dict with fy_label, period_window, cutoff_date, days_until_cutoff,
                          snapshot info, purchase row count, generated_at
    """
    if snapshot_id is None:
        snap = get_latest_snapshot(client_id)
        if snap is None:
            raise ValueError("No GSTR-2A snapshot has been uploaded for this client yet.")
        snapshot_id = snap["snapshot_id"]
    else:
        snap = next((s for s in db.list_snapshots(client_id) if s["snapshot_id"] == snapshot_id), None)
        if snap is None:
            raise ValueError(f"Snapshot #{snapshot_id} not found for this client.")

    period_window = fy_utils.fy_period_window(fy_label, late_filing_months)
    cutoff_date = fy_utils.fy_cutoff_date(fy_label, late_filing_months)

    g2a_full = db.get_snapshot_invoices(snapshot_id)
    g2a_df = g2a_full[g2a_full["period"].isin(period_window)].copy()

    pr_df = db.get_purchase_entries_by_fy(client_id, fy_label)
    if len(pr_df) == 0:
        raise ValueError(f"No Purchase Register entries stored for FY {fy_label}.")

    results = reconcile.run_reconciliation(g2a_df, pr_df)
    pending = db.list_pending_conflicts(client_id)
    results["pending_conflicts"] = pending

    # Flag ITC-at-risk invoices by whether we're already past the late-filing cutoff.
    only_pr = results["only_in_purchase_register"]
    today = datetime.date.today()
    if len(only_pr):
        only_pr = only_pr.copy()
        only_pr["past_6_month_window"] = today > cutoff_date
        results["only_in_purchase_register"] = only_pr

    meta = {
        "fy_label": fy_label,
        "period_window": period_window,
        "period_window_display": f"{period_window[0]} to {period_window[-1]}",
        "cutoff_date": cutoff_date.strftime("%d/%m/%Y"),
        "days_until_cutoff": (cutoff_date - today).days,
        "snapshot_id": snapshot_id,
        "snapshot_uploaded_at": snap["uploaded_at"],
        "snapshot_generation_date": snap["generation_date"],
        "purchase_row_count": len(pr_df),
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "late_filing_months": late_filing_months,
        "pending_conflicts_count": len(pending),
    }
    return {"results": results, "meta": meta}


def export_fy_reconciliation(client_id, fy_label, output_path, snapshot_id=None,
                              late_filing_months=DEFAULT_LATE_FILING_MONTHS):
    """Runs the FY reconciliation and writes the formatted Excel report, logging the run."""
    outcome = run_fy_reconciliation(client_id, fy_label, snapshot_id, late_filing_months)
    results, meta = outcome["results"], outcome["meta"]

    client = db.get_client(client_id)
    snap = next(s for s in db.list_snapshots(client_id) if s["snapshot_id"] == meta["snapshot_id"])
    pr_df = db.get_purchase_entries_by_fy(client_id, fy_label)
    purchase_asof = fy_utils.latest_ddmmyyyy(pr_df["entry_date"]) if len(pr_df) else ""

    report.build_report(
        client["name"], client["gstin"], dict(snap), purchase_asof, results, output_path,
        fy_meta=meta,
    )

    summary = {k: len(v) for k, v in results.items()}
    db.log_recon_run(client_id, meta["snapshot_id"], fy_label, purchase_asof, output_path,
                      __import__("json").dumps(summary))
    return outcome


def export_purchase_month(client_id, year_month, output_path):
    """Exports one month's stored Purchase Register entries as a plain Excel sheet."""
    df = db.get_purchase_entries_by_month(client_id, year_month)
    if len(df) == 0:
        raise ValueError(f"No Purchase Register entries stored for {year_month}.")
    cols = ["entry_date", "particulars", "invoice_no", "invoice_date", "gstin",
            "gross_total", "cgst", "sgst", "igst"]
    out = df[cols].rename(columns={
        "entry_date": "Date", "particulars": "Particulars", "invoice_no": "Supplier Invoice No.",
        "invoice_date": "Supplier Invoice Date", "gstin": "GSTIN/UIN", "gross_total": "Gross Total",
        "cgst": "CGST", "sgst": "SGST", "igst": "IGST",
    })
    out.to_excel(output_path, index=False)
    return output_path
