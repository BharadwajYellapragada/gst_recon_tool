"""
Reconciliation engine: matches GSTR-2A snapshot invoices against stored
Purchase Register entries for a client, and separately matches GSTR-2B
credit/debit notes against stored Credit/Debit Note Register entries.

Invoice categories produced:
  - matched_clean        : ties out within tolerance
  - value_tax_mismatches  : matched by GSTIN+invoice, but amounts differ
  - probable_matches      : GSTIN+date+amount agree but invoice numbers differ
  - only_in_gstr2a         : truly unmatched, present only in GSTR-2A
  - only_in_purchase_register : truly unmatched, present only in books
  - pr_duplicates          : flagged conflict rows already recorded during import
  - multi_rate_reference   : legitimate multi-rate line splits within GSTR-2A (not an error)

Note categories produced by run_note_reconciliation (see its docstring for the
Tally<->GSTR type-flip this applies before matching):
  - matched_clean, value_tax_mismatches, probable_matches : same meaning as above
  - only_in_gstr2b_cdnr    : truly unmatched, present only in the GSTR-2B B2B-CDNR sheet
  - only_in_note_register  : truly unmatched, present only in the Tally CN/DN register
"""
import pandas as pd

from . import parsers

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


# ---------------- Credit/Debit Note reconciliation ----------------

def _agg_cdnr(df):
    return df.groupby("type_key").agg(
        gstin=("gstin", "first"),
        supplier_name=("supplier_name", "first"),
        note_no=("note_no", "first"),
        note_type=("note_type", "first"),
        note_date=("note_date", "first"),
        period=("period", lambda x: ", ".join(sorted(set(x)))),
        note_value=("note_value", "max"),
        taxable_value=("taxable_value", "sum"),
        igst=("igst", "sum"),
        cgst=("cgst", "sum"),
        sgst=("sgst", "sum"),
        row_count=("note_no", "count"),
        itc_available=("itc_available", lambda x: ", ".join(sorted(set(x)))),
    ).reset_index()


def _agg_notes(df):
    return df.groupby("type_key").agg(
        gstin=("gstin", "first"),
        particulars=("particulars", "first"),
        voucher_no=("voucher_no", "first"),
        voucher_date=("voucher_date", "first"),
        book_note_type=("note_type", "first"),
        gstr_equiv_type=("gstr_equiv_type", "first"),
        gross_total=("gross_total", "sum"),
        igst=("igst", "sum"),
        cgst=("cgst", "sum"),
        sgst=("sgst", "sum"),
        row_count=("voucher_no", "count"),
    ).reset_index()


def run_note_reconciliation(cdnr_df, notes_df):
    """
    cdnr_df:  DataFrame from db.get_merged_gstr2b_cdnr_notes() -- GSTR-2B B2B-CDNR rows.
              note_type here is the SUPPLIER's own filing classification: 'Credit
              Note' or 'Debit Note'.
    notes_df: DataFrame from db.get_all_note_entries() / get_note_entries_by_fy() --
              Tally Credit/Debit Note Register rows. note_type here is 'credit' or
              'debit' -- the BUYER's own books classification.

    A Tally 'credit' note is the buyer-side mirror of a supplier's GSTR 'Debit Note'
    filing, and vice versa (see parsers.TALLY_TO_GSTR_NOTE_TYPE): when the buyer
    returns goods, the buyer's books record it as a Credit Note (crediting the
    supplier's account back), while the supplier's GSTR-1 filing records the same
    event as a Debit Note reducing the buyer's ITC... the two sides use opposite
    labels for the identical transaction. So matching flips the Tally side's type
    to its GSTR equivalent before comparing, instead of comparing types directly.

    Returns a dict of DataFrames: matched_clean, value_tax_mismatches,
    probable_matches, only_in_gstr2b_cdnr, only_in_note_register.
    """
    cdnr_df = cdnr_df.copy()
    notes_df = notes_df.copy()
    cdnr_df["type_key"] = cdnr_df["match_key"] + "|" + cdnr_df["note_type"]
    notes_df["gstr_equiv_type"] = notes_df["note_type"].map(parsers.TALLY_TO_GSTR_NOTE_TYPE)
    notes_df["type_key"] = notes_df["match_key"] + "|" + notes_df["gstr_equiv_type"]

    cdnr_agg = _agg_cdnr(cdnr_df)
    notes_agg = _agg_notes(notes_df)

    merged = pd.merge(cdnr_agg, notes_agg, on="type_key", how="outer", indicator=True,
                       suffixes=("_gstr", "_books"))

    both = merged[merged["_merge"] == "both"].copy()
    left_only = merged[merged["_merge"] == "left_only"].copy()
    right_only = merged[merged["_merge"] == "right_only"].copy()

    both["diff_value"] = (both["note_value"] - both["gross_total"]).round(2)
    both["diff_igst"] = (both["igst_gstr"] - both["igst_books"]).round(2)
    both["diff_cgst"] = (both["cgst_gstr"] - both["cgst_books"]).round(2)
    both["diff_sgst"] = (both["sgst_gstr"] - both["sgst_books"]).round(2)
    both["mismatch"] = (
        (both["diff_value"].abs() > TOLERANCE) | (both["diff_igst"].abs() > TOLERANCE) |
        (both["diff_cgst"].abs() > TOLERANCE) | (both["diff_sgst"].abs() > TOLERANCE)
    )
    matched_clean = both[~both["mismatch"]].copy()
    value_tax_mismatches = both[both["mismatch"]].copy()

    probable_matches, true_left, true_right = _fuzzy_note_fallback(left_only, right_only)

    return {
        "matched_clean": matched_clean,
        "value_tax_mismatches": value_tax_mismatches,
        "probable_matches": probable_matches,
        "only_in_gstr2b_cdnr": true_left,
        "only_in_note_register": true_right,
    }


def _fuzzy_note_fallback(left_only, right_only):
    left_only = left_only.copy()
    right_only = right_only.copy()
    left_only["fuzzy_key"] = (
        left_only["gstin_gstr"].astype(str) + "|" + left_only["note_date"].astype(str)
        + "|" + left_only["note_value"].round(0).astype(str) + "|" + left_only["note_type"].astype(str)
    )
    right_only["fuzzy_key"] = (
        right_only["gstin_books"].astype(str) + "|" + right_only["voucher_date"].astype(str)
        + "|" + right_only["gross_total"].round(0).astype(str) + "|" + right_only["gstr_equiv_type"].astype(str)
    )

    l_counts = left_only["fuzzy_key"].value_counts()
    r_counts = right_only["fuzzy_key"].value_counts()
    common = set(l_counts[l_counts == 1].index) & set(r_counts[r_counts == 1].index)

    fuzzy_left = left_only[left_only["fuzzy_key"].isin(common)][
        ["fuzzy_key", "gstin_gstr", "supplier_name", "note_no", "note_date", "note_value", "note_type"]
    ]
    fuzzy_right = right_only[right_only["fuzzy_key"].isin(common)][
        ["fuzzy_key", "gstin_books", "particulars", "voucher_no", "voucher_date", "gross_total", "book_note_type"]
    ]
    probable = pd.merge(fuzzy_left, fuzzy_right, on="fuzzy_key")

    true_left = left_only[~left_only["fuzzy_key"].isin(common)]
    true_right = right_only[~right_only["fuzzy_key"].isin(common)]
    return probable, true_left, true_right
