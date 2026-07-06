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


def run_fy_reconciliation(client_id, fy_label, late_filing_months=DEFAULT_LATE_FILING_MONTHS):
    """
    Reconciles the selected financial year's Purchase Register against a merged,
    de-duplicated view of every GSTR-2A snapshot uploaded for this client (see
    db.get_merged_gstr2a_invoices) for that FY plus `late_filing_months` beyond
    its end -- no manual snapshot selection needed.

    Returns a dict:
      results        -> the category DataFrames from reconcile.run_reconciliation(),
                         with 'only_in_purchase_register' additionally carrying a
                         'past_6_month_window' boolean column
      meta            -> dict with fy_label, period_window, cutoff_date, days_until_cutoff,
                          contributing-snapshot info, purchase row count, generated_at
    """
    period_window = fy_utils.fy_period_window(fy_label, late_filing_months)
    cutoff_date = fy_utils.fy_cutoff_date(fy_label, late_filing_months)

    g2a_df = db.get_merged_gstr2a_invoices(client_id, period_window)
    if len(g2a_df) == 0:
        raise ValueError(
            f"No GSTR-2A data covers the period window for FY {fy_label} "
            f"({period_window[0]} to {period_window[-1]}). Upload a GSTR-2A file first."
        )

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

    contributing_ids = set(g2a_df["snapshot_id"].unique().tolist())
    contributing_snaps = [s for s in db.list_snapshots(client_id) if s["snapshot_id"] in contributing_ids]
    latest_snap = max(contributing_snaps, key=lambda s: s["uploaded_at"])

    meta = {
        "fy_label": fy_label,
        "period_window": period_window,
        "period_window_display": f"{period_window[0]} to {period_window[-1]}",
        "cutoff_date": cutoff_date.strftime("%d/%m/%Y"),
        "days_until_cutoff": (cutoff_date - today).days,
        "gstr2a_snapshot_count": len(contributing_snaps),
        "gstr2a_latest_snapshot_id": latest_snap["snapshot_id"],
        "gstr2a_latest_uploaded_at": latest_snap["uploaded_at"],
        "gstr2a_latest_generation_date": latest_snap["generation_date"],
        "purchase_row_count": len(pr_df),
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "late_filing_months": late_filing_months,
        "pending_conflicts_count": len(pending),
    }
    return {"results": results, "meta": meta}


def export_fy_reconciliation(client_id, fy_label, output_path, late_filing_months=DEFAULT_LATE_FILING_MONTHS):
    """Runs the FY reconciliation and writes the formatted Excel report, logging the run."""
    outcome = run_fy_reconciliation(client_id, fy_label, late_filing_months)
    results, meta = outcome["results"], outcome["meta"]

    client = db.get_client(client_id)
    pr_df = db.get_purchase_entries_by_fy(client_id, fy_label)
    purchase_asof = fy_utils.latest_ddmmyyyy(pr_df["entry_date"]) if len(pr_df) else ""

    report.build_report(
        client["name"], client["gstin"], {}, purchase_asof, results, output_path,
        fy_meta=meta,
    )

    summary = {k: len(v) for k, v in results.items() if k != "pending_conflicts"}
    db.log_recon_run(client_id, meta["gstr2a_latest_snapshot_id"], fy_label, purchase_asof, output_path,
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


def export_purchase_batch(batch_id, output_path):
    """Exports every row a single Purchase Register upload added, as a plain Excel sheet
    — since the originally uploaded file itself isn't kept, this re-derives an equivalent
    file from what was stored, for a user who wants a copy of what an upload batch contains."""
    df = db.get_purchase_entries_by_batch(batch_id)
    if len(df) == 0:
        raise ValueError("No entries found for this upload batch.")
    cols = ["entry_date", "particulars", "invoice_no", "invoice_date", "gstin",
            "gross_total", "cgst", "sgst", "igst"]
    out = df[cols].rename(columns={
        "entry_date": "Date", "particulars": "Particulars", "invoice_no": "Supplier Invoice No.",
        "invoice_date": "Supplier Invoice Date", "gstin": "GSTIN/UIN", "gross_total": "Gross Total",
        "cgst": "CGST", "sgst": "SGST", "igst": "IGST",
    })
    out.to_excel(output_path, index=False)
    return output_path


def export_gstr2a_snapshot(snapshot_id, output_path):
    """Exports a GSTR-2A/2B snapshot's stored invoices as a plain Excel sheet — a
    re-derived copy (the originally uploaded portal file itself isn't kept)."""
    df = db.get_snapshot_invoices(snapshot_id)
    if len(df) == 0:
        raise ValueError("No invoices found for this snapshot.")
    cols = ["period", "gstin", "supplier_name", "invoice_no", "invoice_date", "invoice_value",
            "rate", "taxable_value", "igst", "cgst", "sgst", "cess", "filing_date",
            "itc_available", "reason"]
    out = df[cols].rename(columns={
        "period": "Period", "gstin": "GSTIN", "supplier_name": "Supplier Name",
        "invoice_no": "Invoice No", "invoice_date": "Invoice Date", "invoice_value": "Invoice Value",
        "rate": "Rate", "taxable_value": "Taxable Value", "igst": "IGST", "cgst": "CGST",
        "sgst": "SGST", "cess": "Cess", "filing_date": "Filing Date",
        "itc_available": "ITC Available", "reason": "Reason",
    })
    out.to_excel(output_path, index=False)
    return output_path
