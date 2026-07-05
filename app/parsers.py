"""
Parsers for the two source file types:
  - GSTR-2A export from the GST portal (sheet 'B2B')
  - Internal Purchase Register export (e.g. from Tally)

Both .xls (legacy, via xlrd) and .xlsx (via openpyxl) are supported.
"""
import re
import datetime
import pandas as pd

from . import fy_utils

G2A_COLS = ['Period', 'GSTIN', 'SupplierName', 'InvoiceNo', 'InvoiceType', 'InvoiceDate',
            'InvoiceValue', 'PlaceOfSupply', 'RCM', 'Rate', 'TaxableValue', 'IGST', 'CGST',
            'SGST', 'Cess', 'FilingDate', 'ITCAvailable', 'Reason']


class ParseError(Exception):
    pass


def _norm_inv(x):
    x = str(x).upper().strip()
    return re.sub(r'[^A-Z0-9]', '', x)


def _num(x):
    return x if isinstance(x, (int, float)) else 0


def _open_book(path):
    if path.lower().endswith(".xls"):
        import xlrd
        return ("xlrd", xlrd.open_workbook(path))
    else:
        from openpyxl import load_workbook
        return ("openpyxl", load_workbook(path, read_only=True, data_only=True))


def parse_gstr2a(path):
    """
    Returns (invoices_df, generation_date_str).
    invoices_df has columns G2A_COLS + MatchKey, GSTIN uppercased.
    Raises ParseError with a user-facing message on bad file.
    """
    kind, wb = _open_book(path)
    sheet_names = wb.sheet_names() if kind == "xlrd" else wb.sheetnames
    if "B2B" not in sheet_names:
        raise ParseError(
            "This doesn't look like a GSTR-2A export — no 'B2B' sheet found. "
            "Please upload the .xls/.xlsx file downloaded from the GST portal's "
            "'Download GSTR-2A' option."
        )

    generation_date = ""
    if "Read me" in sheet_names:
        try:
            if kind == "xlrd":
                rm = wb.sheet_by_name("Read me")
                for r in range(min(10, rm.nrows)):
                    row = rm.row_values(r)
                    if len(row) > 3 and "generation" in str(row[3]).lower():
                        generation_date = str(row[4])
            else:
                rm = wb["Read me"]
                for row in rm.iter_rows(min_row=1, max_row=10, values_only=True):
                    for i, cell in enumerate(row):
                        if cell and "generation" in str(cell).lower() and i + 1 < len(row):
                            generation_date = str(row[i + 1])
        except Exception:
            pass

    if kind == "xlrd":
        ws = wb.sheet_by_name("B2B")
        rows = [ws.row_values(r) for r in range(6, ws.nrows)]
    else:
        ws = wb["B2B"]
        rows = [list(r) for r in ws.iter_rows(min_row=7, values_only=True)]

    rows = [r[:18] for r in rows if r and str(r[1] if len(r) > 1 else "").strip()]
    if not rows:
        raise ParseError("No invoice rows found in the B2B sheet. The file may be empty or in an unexpected format.")

    df = pd.DataFrame(rows, columns=G2A_COLS)
    df = df[df["GSTIN"].astype(str).str.strip() != ""].copy()
    for c in ["InvoiceValue", "Rate", "TaxableValue", "IGST", "CGST", "SGST", "Cess"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["GSTIN"] = df["GSTIN"].astype(str).str.strip().str.upper()
    df["InvoiceNo"] = df["InvoiceNo"].astype(str).str.strip()
    df["MatchKey"] = df["GSTIN"] + "|" + df["InvoiceNo"].apply(_norm_inv)
    return df, generation_date


def _excel_serial_to_ddmmyyyy(serial):
    try:
        d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(serial))
        return d.strftime("%d/%m/%Y")
    except Exception:
        return str(serial)


def parse_purchase_register(path):
    """
    Returns a list of dicts ready for db.add_purchase_batch(), by locating the header
    row dynamically (looks for 'Date' + 'Gross Total' + a GSTIN-like column), then summing
    all CGST/SGST-rate-split columns into single CGST/SGST totals.
    """
    kind, wb = _open_book(path)
    sheet_names = wb.sheet_names() if kind == "xlrd" else wb.sheetnames
    target_sheet = sheet_names[0]
    for s in sheet_names:
        if "purchase" in s.lower():
            target_sheet = s
            break

    if kind == "xlrd":
        ws = wb.sheet_by_name(target_sheet)
        all_rows = [ws.row_values(r) for r in range(ws.nrows)]
    else:
        ws = wb[target_sheet]
        all_rows = [list(r) for r in ws.iter_rows(values_only=True)]

    header_row_idx = None
    for i, row in enumerate(all_rows[:40]):
        cells = [str(c).strip().lower() for c in row if c is not None]
        if any("gross total" in c for c in cells) and any(c == "date" for c in cells):
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ParseError(
            "Could not find the header row (expected columns like 'Date' and 'Gross Total'). "
            "Please upload the Purchase Register export as-is, without removing header rows."
        )

    header = [str(c).strip() if c is not None else "" for c in all_rows[header_row_idx]]
    data_rows = all_rows[header_row_idx + 1:]

    def col_idx(*names):
        for n in names:
            for i, h in enumerate(header):
                if h.strip().lower() == n.lower():
                    return i
        return None

    i_date = col_idx("Date")
    i_particulars = col_idx("Particulars")
    i_invno = col_idx("Supplier Invoice No.", "Supplier Invoice No")
    i_invdate = col_idx("Supplier Invoice Date")
    i_gstin = col_idx("GSTIN/UIN", "GSTIN")
    i_gross = col_idx("Gross Total")
    i_igst = col_idx("IGST")

    if None in (i_date, i_invno, i_gstin, i_gross):
        raise ParseError(
            "Some required columns (Date, Supplier Invoice No., GSTIN/UIN, Gross Total) "
            "were not found by name. The file's column headers may differ from the expected export format."
        )

    cgst_idx = [i for i, h in enumerate(header) if h.lower().startswith("cgst")]
    sgst_idx = [i for i, h in enumerate(header) if h.lower().startswith("sgst")]

    entries = []
    for row in data_rows:
        if row is None:
            continue
        if i_date >= len(row):
            continue
        date_val = row[i_date]
        if not isinstance(date_val, (int, float)) or date_val <= 0:
            continue
        gstin = str(row[i_gstin]).strip().upper() if i_gstin < len(row) and row[i_gstin] else ""
        invno = str(row[i_invno]).strip() if i_invno < len(row) and row[i_invno] else ""
        if not gstin or not invno:
            continue

        cgst = sum(_num(row[i]) for i in cgst_idx if i < len(row))
        sgst = sum(_num(row[i]) for i in sgst_idx if i < len(row))
        igst = _num(row[i_igst]) if i_igst is not None and i_igst < len(row) else 0
        gross = _num(row[i_gross]) if i_gross < len(row) else 0
        inv_date_raw = row[i_invdate] if i_invdate is not None and i_invdate < len(row) else date_val
        inv_date_fmt = _excel_serial_to_ddmmyyyy(inv_date_raw) if isinstance(inv_date_raw, (int, float)) else str(inv_date_raw)
        entry_date_fmt = _excel_serial_to_ddmmyyyy(date_val)
        entry_date_obj = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(date_val))
        entry_fy = fy_utils.fy_label_for_date(entry_date_obj.date())
        entry_month = entry_date_obj.strftime("%Y-%m")

        entries.append({
            "entry_date": entry_date_fmt,
            "entry_fy": entry_fy,
            "entry_month": entry_month,
            "particulars": str(row[i_particulars]).strip() if i_particulars is not None and i_particulars < len(row) and row[i_particulars] else "",
            "invoice_no": invno,
            "invoice_date": inv_date_fmt,
            "gstin": gstin,
            "gross_total": round(gross, 2),
            "cgst": round(cgst, 2),
            "sgst": round(sgst, 2),
            "igst": round(igst, 2),
            "match_key": gstin + "|" + _norm_inv(invno),
        })

    if not entries:
        raise ParseError("No valid data rows were found in the Purchase Register file.")
    return entries
