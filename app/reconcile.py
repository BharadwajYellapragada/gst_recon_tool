"""
Reconciliation engine: matches GSTR-2A snapshot invoices against stored
Purchase Register entries for a client.

Categories produced:
  - matched_clean        : ties out within tolerance
  - value_tax_mismatches  : matched by GSTIN+invoice, but amounts differ
  - probable_matches      : GSTIN+date+amount agree but invoice numbers differ
  - only_in_gstr2a         : truly unmatched, present only in GSTR-2A
  - only_in_purchase_register : truly unmatched, present only in books
  - pr_duplicates          : flagged conflict rows already recorded during import
  - multi_rate_reference   : legitimate multi-rate line splits within GSTR-2A (not an error)
"""
import pandas as pd

TOLERANCE = 2.0


def _agg_g2a(df):
    return df.groupby("match_key").agg(
        gstin=("gstin", "first"),
        supplier_name=("supplier_name", "first"),
        invoice_no=("invoice_no", "first"),
        invoice_date=("invoice_date", "first"),
        period=("period", lambda x: ", ".join(sorted(set(x)))),
        invoice_value=("invoice_value", "max"),
        taxable_value=("taxable_value", "sum"),
        igst=("igst", "sum"),
        cgst=("cgst", "sum"),
        sgst=("sgst", "sum"),
        row_count=("invoice_no", "count"),
        itc_available=("itc_available", lambda x: ", ".join(sorted(set(x)))),
    ).reset_index()


def _agg_pr(df):
    return df.groupby("match_key").agg(
        gstin=("gstin", "first"),
        particulars=("particulars", "first"),
        invoice_no=("invoice_no", "first"),
        invoice_date=("invoice_date", "first"),
        gross_total=("gross_total", "sum"),
        igst=("igst", "sum"),
        cgst=("cgst", "sum"),
        sgst=("sgst", "sum"),
        row_count=("invoice_no", "count"),
    ).reset_index()


def find_multi_rate_invoices(g2a_df):
    counts = g2a_df.groupby("match_key").size()
    dup_keys = counts[counts > 1].index
    if len(dup_keys) == 0:
        return g2a_df.iloc[0:0], g2a_df.iloc[0:0]
    dupdf = g2a_df[g2a_df["match_key"].isin(dup_keys)]

    def classify(grp):
        combos = grp[["rate", "taxable_value"]].drop_duplicates()
        return "MULTI_RATE_SPLIT" if len(combos) == len(grp) else "TRUE_DUPLICATE"

    cls = dupdf.groupby("match_key").apply(classify)
    true_dup_keys = cls[cls == "TRUE_DUPLICATE"].index
    multi_rate_keys = cls[cls == "MULTI_RATE_SPLIT"].index
    true_dups = g2a_df[g2a_df["match_key"].isin(true_dup_keys)]
    multi_rate = g2a_df[g2a_df["match_key"].isin(multi_rate_keys)]
    return true_dups, multi_rate


def run_reconciliation(g2a_df, pr_df):
    """
    g2a_df: DataFrame from db.get_snapshot_invoices() (columns: gstin, supplier_name,
            invoice_no, invoice_date, invoice_value, rate, taxable_value, igst, cgst, sgst,
            period, itc_available, match_key, ...)
    pr_df:  DataFrame from db.get_all_purchase_entries() (columns: gstin, particulars,
            invoice_no, invoice_date, gross_total, cgst, sgst, igst, match_key,
            is_conflict, conflict_note, ...)

    Returns a dict of DataFrames, one per category.
    """
    true_dups, multi_rate = find_multi_rate_invoices(g2a_df)

    g2a_agg = _agg_g2a(g2a_df)
    pr_agg = _agg_pr(pr_df)

    merged = pd.merge(g2a_agg, pr_agg, on="match_key", how="outer", indicator=True,
                       suffixes=("_2a", "_pr"))

    both = merged[merged["_merge"] == "both"].copy()
    left_only = merged[merged["_merge"] == "left_only"].copy()
    right_only = merged[merged["_merge"] == "right_only"].copy()

    both["diff_value"] = (both["invoice_value"] - both["gross_total"]).round(2)
    both["diff_igst"] = (both["igst_2a"] - both["igst_pr"]).round(2)
    both["diff_cgst"] = (both["cgst_2a"] - both["cgst_pr"]).round(2)
    both["diff_sgst"] = (both["sgst_2a"] - both["sgst_pr"]).round(2)
    both["mismatch"] = (
        (both["diff_value"].abs() > TOLERANCE) | (both["diff_igst"].abs() > TOLERANCE) |
        (both["diff_cgst"].abs() > TOLERANCE) | (both["diff_sgst"].abs() > TOLERANCE)
    )
    matched_clean = both[~both["mismatch"]].copy()
    value_tax_mismatches = both[both["mismatch"]].copy()

    # Fuzzy fallback pass: same GSTIN + invoice date + amount, invoice number differs.
    # left_only/right_only share the SAME merged column set (both _2a and _pr suffixed
    # columns exist in every row; the non-matching side is just NaN), so we must select
    # columns by which side is populated, not by generic name lookup.
    probable_matches, true_left, true_right = _fuzzy_fallback(left_only, right_only)

    return {
        "matched_clean": matched_clean,
        "value_tax_mismatches": value_tax_mismatches,
        "probable_matches": probable_matches,
        "only_in_gstr2a": true_left,
        "only_in_purchase_register": true_right,
        "true_duplicates_gstr2a": true_dups,
        "multi_rate_reference": multi_rate,
    }


def _fuzzy_fallback(left_only, right_only):
    # left_only rows: the _2a-suffixed columns are populated, _pr columns are NaN.
    # right_only rows: the reverse. invoice_value/gross_total are never suffixed
    # (they don't collide across the two source frames), so they're always safe as-is.
    left_only = left_only.copy()
    right_only = right_only.copy()
    left_only["fuzzy_key"] = (
        left_only["gstin_2a"].astype(str) + "|" + left_only["invoice_date_2a"].astype(str)
        + "|" + left_only["invoice_value"].round(0).astype(str)
    )
    right_only["fuzzy_key"] = (
        right_only["gstin_pr"].astype(str) + "|" + right_only["invoice_date_pr"].astype(str)
        + "|" + right_only["gross_total"].round(0).astype(str)
    )

    l_counts = left_only["fuzzy_key"].value_counts()
    r_counts = right_only["fuzzy_key"].value_counts()
    common = set(l_counts[l_counts == 1].index) & set(r_counts[r_counts == 1].index)

    # Select + rename to a clean, flat schema BEFORE merging, so the merge
    # (whose two sides already share no overlapping names) can't double-suffix.
    fuzzy_left = left_only[left_only["fuzzy_key"].isin(common)][
        ["fuzzy_key", "gstin_2a", "supplier_name", "invoice_no_2a", "invoice_date_2a", "invoice_value"]
    ].rename(columns={
        "gstin_2a": "gstin_2a", "invoice_no_2a": "invoice_no_2a", "invoice_date_2a": "invoice_date_2a",
    })
    fuzzy_right = right_only[right_only["fuzzy_key"].isin(common)][
        ["fuzzy_key", "gstin_pr", "particulars", "invoice_no_pr", "invoice_date_pr", "gross_total"]
    ]
    probable = pd.merge(fuzzy_left, fuzzy_right, on="fuzzy_key")

    true_left = left_only[~left_only["fuzzy_key"].isin(common)]
    true_right = right_only[~right_only["fuzzy_key"].isin(common)]
    return probable, true_left, true_right
