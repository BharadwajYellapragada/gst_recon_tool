"""
SQLite data layer for the GST Reconciliation Tool.

Security model: the working database lives ONLY in memory during a session.
On disk, it exists only as a Fernet-encrypted blob (gst_recon.db.enc). The
encryption key is derived from a password + this machine's identity (see
security.py) — see unlock_or_init(). Every write is followed by persist(),
which re-serializes the in-memory DB and overwrites the encrypted file.
There is no plaintext database file at any point on disk.
"""
import sqlite3
import hashlib
import os
import sys
import tempfile
from datetime import datetime

from . import security


def get_data_dir():
    """Local per-user app data folder. Created on first run."""
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    else:
        base = os.path.expanduser("~/.local/share")
    path = os.path.join(base, "GSTReconTool")
    os.makedirs(path, exist_ok=True)
    os.makedirs(os.path.join(path, "reports"), exist_ok=True)
    return path


def ENC_DB_PATH():
    return os.path.join(get_data_dir(), "gst_recon.db.enc")


SCHEMA = """
CREATE TABLE IF NOT EXISTS clients (
    client_id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    gstin TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS gstr2a_snapshots (
    snapshot_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    uploaded_at TEXT NOT NULL,
    source_filename TEXT,
    generation_date TEXT,
    period_min TEXT,
    period_max TEXT,
    row_count INTEGER,
    FOREIGN KEY(client_id) REFERENCES clients(client_id)
);

CREATE TABLE IF NOT EXISTS gstr2a_invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    snapshot_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    period TEXT,
    gstin TEXT,
    supplier_name TEXT,
    invoice_no TEXT,
    invoice_date TEXT,
    invoice_value REAL,
    rate REAL,
    taxable_value REAL,
    igst REAL,
    cgst REAL,
    sgst REAL,
    cess REAL,
    filing_date TEXT,
    itc_available TEXT,
    reason TEXT,
    match_key TEXT,
    FOREIGN KEY(snapshot_id) REFERENCES gstr2a_snapshots(snapshot_id)
);
CREATE INDEX IF NOT EXISTS idx_g2a_matchkey ON gstr2a_invoices(snapshot_id, match_key);

CREATE TABLE IF NOT EXISTS purchase_batches (
    batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    uploaded_at TEXT NOT NULL,
    source_filename TEXT,
    row_count INTEGER,
    new_rows INTEGER,
    duplicate_rows_skipped INTEGER,
    conflict_rows INTEGER,
    FOREIGN KEY(client_id) REFERENCES clients(client_id)
);

CREATE TABLE IF NOT EXISTS purchase_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    batch_id INTEGER NOT NULL,
    entry_date TEXT,
    entry_fy TEXT,
    entry_month TEXT,
    particulars TEXT,
    invoice_no TEXT,
    invoice_date TEXT,
    gstin TEXT,
    gross_total REAL,
    cgst REAL,
    sgst REAL,
    igst REAL,
    match_key TEXT,
    row_hash TEXT UNIQUE,
    is_conflict INTEGER DEFAULT 0,
    conflict_note TEXT,
    status TEXT DEFAULT 'active',
    FOREIGN KEY(client_id) REFERENCES clients(client_id),
    FOREIGN KEY(batch_id) REFERENCES purchase_batches(batch_id)
);
CREATE INDEX IF NOT EXISTS idx_pr_matchkey ON purchase_entries(client_id, match_key);
CREATE INDEX IF NOT EXISTS idx_pr_fy ON purchase_entries(client_id, entry_fy);
CREATE INDEX IF NOT EXISTS idx_pr_month ON purchase_entries(client_id, entry_month);
CREATE INDEX IF NOT EXISTS idx_pr_status ON purchase_entries(client_id, status);

CREATE TABLE IF NOT EXISTS recon_runs (
    run_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    snapshot_id INTEGER NOT NULL,
    fy_label TEXT,
    run_at TEXT NOT NULL,
    purchase_asof TEXT,
    report_path TEXT,
    summary_json TEXT,
    FOREIGN KEY(client_id) REFERENCES clients(client_id),
    FOREIGN KEY(snapshot_id) REFERENCES gstr2a_snapshots(snapshot_id)
);
"""


def _ensure_columns(conn):
    """Lightweight migration: add columns that may be missing from a database
    created by an earlier version of this tool, without disturbing existing data."""
    def cols(table):
        return {r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}

    pe_cols = cols("purchase_entries")
    if "entry_fy" not in pe_cols:
        conn.execute("ALTER TABLE purchase_entries ADD COLUMN entry_fy TEXT")
    if "entry_month" not in pe_cols:
        conn.execute("ALTER TABLE purchase_entries ADD COLUMN entry_month TEXT")
    if "status" not in pe_cols:
        conn.execute("ALTER TABLE purchase_entries ADD COLUMN status TEXT DEFAULT 'active'")
        conn.execute("UPDATE purchase_entries SET status='active' WHERE status IS NULL")

    rr_cols = cols("recon_runs")
    if "fy_label" not in rr_cols:
        conn.execute("ALTER TABLE recon_runs ADD COLUMN fy_label TEXT")
    conn.commit()


# ---------------- session state: in-memory connection + encryption key ----------------

_conn = None
_key = None


class NotUnlockedError(RuntimeError):
    pass


def is_unlocked():
    return _conn is not None


def needs_setup():
    """True if this is the first run on this machine (no password set yet)."""
    return not security.is_initialized()


def unlock_or_init(password):
    """
    Call once at app startup with the password the user typed.
    - First run: sets up a new password + a fresh empty database.
    - Later runs: verifies the password (and implicitly, the machine) and
      loads the existing encrypted database into memory.
    Raises security.SecurityError on wrong password / wrong machine.
    """
    global _conn, _key

    if security.is_initialized():
        key = security.unlock(password)  # raises SecurityError if wrong
        conn = sqlite3.connect(":memory:")
        conn.row_factory = sqlite3.Row
        enc_path = ENC_DB_PATH()
        if os.path.exists(enc_path):
            with open(enc_path, "rb") as f:
                token = f.read()
            raw = security.decrypt_bytes(key, token)
            conn.deserialize(raw)
        else:
            conn.executescript(SCHEMA)
            conn.commit()
        conn.execute("PRAGMA foreign_keys = ON")
        _ensure_columns(conn)
    else:
        key = security.setup_new_password(password)
        conn = sqlite3.connect(":memory:")
        conn.row_factory = sqlite3.Row
        conn.executescript(SCHEMA)
        conn.commit()
        conn.execute("PRAGMA foreign_keys = ON")

    _conn = conn
    _key = key
    persist()  # ensure an encrypted file exists on disk from the very first run


def lock():
    """Drops the in-memory database from this process. Does not touch the encrypted file."""
    global _conn, _key
    _conn = None
    _key = None


def get_conn():
    if _conn is None:
        raise NotUnlockedError("Database is locked. Call db.unlock_or_init(password) first.")
    return _conn


def persist():
    """Serializes the in-memory database and overwrites the encrypted file on disk.
    Called after every write. Write is atomic (temp file + rename) to avoid a
    corrupted file if the app is killed mid-write; the temp file holds only
    already-encrypted ciphertext, never plaintext."""
    if _conn is None or _key is None:
        return
    raw = _conn.serialize()
    token = security.encrypt_bytes(_key, bytes(raw))
    enc_path = ENC_DB_PATH()
    fd, tmp_path = tempfile.mkstemp(dir=os.path.dirname(enc_path), prefix=".tmp_gstrecon_")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(token)
        os.replace(tmp_path, enc_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def row_hash(client_id, match_key, invoice_date, gross_total, cgst, sgst, igst):
    s = f"{client_id}|{match_key}|{invoice_date}|{gross_total:.2f}|{cgst:.2f}|{sgst:.2f}|{igst:.2f}"
    return hashlib.sha256(s.encode()).hexdigest()


# ---------------- Client operations ----------------

def add_client(name, gstin=""):
    conn = get_conn()
    conn.execute(
        "INSERT INTO clients (name, gstin, created_at) VALUES (?, ?, ?)",
        (name.strip(), gstin.strip(), now()),
    )
    conn.commit()
    client_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    persist()
    return client_id


def list_clients():
    return get_conn().execute("SELECT * FROM clients ORDER BY name COLLATE NOCASE").fetchall()


def get_client(client_id):
    return get_conn().execute("SELECT * FROM clients WHERE client_id=?", (client_id,)).fetchone()


def delete_client(client_id):
    conn = get_conn()
    conn.execute("DELETE FROM recon_runs WHERE client_id=?", (client_id,))
    conn.execute("DELETE FROM purchase_entries WHERE client_id=?", (client_id,))
    conn.execute("DELETE FROM purchase_batches WHERE client_id=?", (client_id,))
    conn.execute("DELETE FROM gstr2a_invoices WHERE client_id=?", (client_id,))
    conn.execute("DELETE FROM gstr2a_snapshots WHERE client_id=?", (client_id,))
    conn.execute("DELETE FROM clients WHERE client_id=?", (client_id,))
    conn.commit()
    persist()


# ---------------- GSTR-2A snapshot operations ----------------

def add_gstr2a_snapshot(client_id, source_filename, generation_date, invoices_df):
    conn = get_conn()
    periods = sorted(invoices_df["Period"].unique()) if len(invoices_df) else []
    cur = conn.execute(
        """INSERT INTO gstr2a_snapshots
           (client_id, uploaded_at, source_filename, generation_date, period_min, period_max, row_count)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (client_id, now(), source_filename, generation_date,
         periods[0] if periods else "", periods[-1] if periods else "", len(invoices_df)),
    )
    snapshot_id = cur.lastrowid
    records = [
        (snapshot_id, client_id, r.Period, r.GSTIN, r.SupplierName, r.InvoiceNo, r.InvoiceDate,
         r.InvoiceValue, r.Rate, r.TaxableValue, r.IGST, r.CGST, r.SGST, r.Cess,
         r.FilingDate, r.ITCAvailable, r.Reason, r.MatchKey)
        for r in invoices_df.itertuples()
    ]
    conn.executemany(
        """INSERT INTO gstr2a_invoices
           (snapshot_id, client_id, period, gstin, supplier_name, invoice_no, invoice_date,
            invoice_value, rate, taxable_value, igst, cgst, sgst, cess,
            filing_date, itc_available, reason, match_key)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        records,
    )
    conn.commit()
    persist()
    return snapshot_id


def list_snapshots(client_id):
    return get_conn().execute(
        "SELECT * FROM gstr2a_snapshots WHERE client_id=? ORDER BY uploaded_at DESC", (client_id,)
    ).fetchall()


def get_snapshot_invoices(snapshot_id):
    import pandas as pd
    return pd.read_sql_query(
        "SELECT * FROM gstr2a_invoices WHERE snapshot_id=?", get_conn(), params=(snapshot_id,)
    )


# ---------------- Purchase Register operations ----------------

def add_purchase_batch(client_id, source_filename, entries):
    """
    entries: list of dicts with keys entry_date, entry_fy, entry_month, particulars,
             invoice_no, invoice_date, gstin, gross_total, cgst, sgst, igst, match_key

    Behaviour on a re-uploaded record:
      - Identical to what's already stored (same invoice + same amounts)  -> silently
        skipped. Nothing to decide; the data already matches.
      - Same invoice, DIFFERENT amount than what's stored                -> NOT inserted
        as active data. Stored with status='pending_conflict' and excluded from every
        reconciliation/query until resolve_conflict() is called with the user's choice
        of 'overwrite' or 'ignore'. See list_pending_conflicts().

    Returns dict with counts: new_rows, duplicate_rows_skipped, pending_conflicts.
    """
    conn = get_conn()
    cur = conn.execute(
        """INSERT INTO purchase_batches
           (client_id, uploaded_at, source_filename, row_count, new_rows, duplicate_rows_skipped, conflict_rows)
           VALUES (?, ?, ?, ?, 0, 0, 0)""",
        (client_id, now(), source_filename, len(entries)),
    )
    batch_id = cur.lastrowid

    new_rows = 0
    dup_rows = 0
    pending_conflicts = 0

    existing = conn.execute(
        "SELECT match_key, invoice_date, gross_total, cgst, sgst, igst FROM purchase_entries "
        "WHERE client_id=? AND status='active'",
        (client_id,),
    ).fetchall()
    existing_by_key = {}
    for e in existing:
        existing_by_key.setdefault(e["match_key"], []).append(e)

    for e in entries:
        rh = row_hash(client_id, e["match_key"], e["invoice_date"], e["gross_total"], e["cgst"], e["sgst"], e["igst"])
        already = conn.execute("SELECT 1 FROM purchase_entries WHERE row_hash=?", (rh,)).fetchone()
        if already:
            dup_rows += 1
            continue

        prior = existing_by_key.get(e["match_key"], [])
        is_conflict = 1 if prior else 0
        status = "pending_conflict" if prior else "active"
        conflict_note = None
        if prior:
            p = prior[0]
            conflict_note = (
                f"Same invoice already stored with different amount: "
                f"Gross {p['gross_total']:.2f} (stored) vs {e['gross_total']:.2f} (this upload)"
            )
            pending_conflicts += 1

        conn.execute(
            """INSERT INTO purchase_entries
               (client_id, batch_id, entry_date, entry_fy, entry_month, particulars, invoice_no,
                invoice_date, gstin, gross_total, cgst, sgst, igst, match_key, row_hash,
                is_conflict, conflict_note, status)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (client_id, batch_id, e["entry_date"], e.get("entry_fy", ""), e.get("entry_month", ""),
             e["particulars"], e["invoice_no"], e["invoice_date"], e["gstin"], e["gross_total"],
             e["cgst"], e["sgst"], e["igst"], e["match_key"], rh, is_conflict, conflict_note, status),
        )
        if status == "active":
            new_rows += 1
            existing_by_key.setdefault(e["match_key"], []).append(
                {"match_key": e["match_key"], "invoice_date": e["invoice_date"],
                 "gross_total": e["gross_total"], "cgst": e["cgst"], "sgst": e["sgst"], "igst": e["igst"]}
            )

    conn.execute(
        "UPDATE purchase_batches SET new_rows=?, duplicate_rows_skipped=?, conflict_rows=? WHERE batch_id=?",
        (new_rows, dup_rows, pending_conflicts, batch_id),
    )
    conn.commit()
    persist()
    return {"batch_id": batch_id, "new_rows": new_rows, "duplicate_rows_skipped": dup_rows,
            "pending_conflicts": pending_conflicts, "total_in_file": len(entries)}


def list_pending_conflicts(client_id):
    """
    Every stored row still awaiting a user decision, paired with the currently-active
    row it collides with, so a GUI can show 'stored value' vs 'new value' side by side.
    """
    conn = get_conn()
    pending = conn.execute(
        "SELECT * FROM purchase_entries WHERE client_id=? AND status='pending_conflict' ORDER BY id",
        (client_id,),
    ).fetchall()
    result = []
    for p in pending:
        active = conn.execute(
            "SELECT * FROM purchase_entries WHERE client_id=? AND match_key=? AND status='active'",
            (client_id, p["match_key"]),
        ).fetchone()
        result.append({"pending": dict(p), "stored": dict(active) if active else None})
    return result


def resolve_conflict(pending_entry_id, action):
    """
    action: 'overwrite' (the new upload replaces what's stored) or
            'ignore' (discard the new upload, keep what's stored).
    Nothing is ever hard-deleted — the losing side is marked 'superseded'/'ignored'
    for audit purposes, and excluded from all reconciliation queries.
    """
    if action not in ("overwrite", "ignore"):
        raise ValueError("action must be 'overwrite' or 'ignore'")
    conn = get_conn()
    pending = conn.execute(
        "SELECT * FROM purchase_entries WHERE id=? AND status='pending_conflict'", (pending_entry_id,)
    ).fetchone()
    if pending is None:
        raise ValueError(f"No pending conflict found with id={pending_entry_id}")

    if action == "ignore":
        conn.execute("UPDATE purchase_entries SET status='ignored' WHERE id=?", (pending_entry_id,))
    else:  # overwrite
        conn.execute(
            "UPDATE purchase_entries SET status='superseded' WHERE client_id=? AND match_key=? AND status='active'",
            (pending["client_id"], pending["match_key"]),
        )
        conn.execute("UPDATE purchase_entries SET status='active' WHERE id=?", (pending_entry_id,))
    conn.commit()
    persist()


def resolve_all_conflicts(client_id, action):
    """Bulk-apply the same decision to every pending conflict for this client."""
    for item in list_pending_conflicts(client_id):
        resolve_conflict(item["pending"]["id"], action)


def list_purchase_batches(client_id):
    return get_conn().execute(
        "SELECT * FROM purchase_batches WHERE client_id=? ORDER BY uploaded_at DESC", (client_id,)
    ).fetchall()


def get_all_purchase_entries(client_id):
    import pandas as pd
    return pd.read_sql_query(
        "SELECT * FROM purchase_entries WHERE client_id=? AND status='active'", get_conn(), params=(client_id,)
    )


def list_available_fys(client_id):
    """Distinct financial years present in this client's stored (active) Purchase
    Register data, most recent first."""
    rows = get_conn().execute(
        "SELECT DISTINCT entry_fy FROM purchase_entries WHERE client_id=? AND status='active' "
        "AND entry_fy IS NOT NULL AND entry_fy != '' ORDER BY entry_fy DESC",
        (client_id,),
    ).fetchall()
    return [r["entry_fy"] for r in rows]


def list_available_months(client_id):
    """Distinct YYYY-MM months present among active entries, most recent first —
    used for month-wise download."""
    rows = get_conn().execute(
        "SELECT DISTINCT entry_month FROM purchase_entries WHERE client_id=? AND status='active' "
        "AND entry_month IS NOT NULL AND entry_month != '' ORDER BY entry_month DESC",
        (client_id,),
    ).fetchall()
    return [r["entry_month"] for r in rows]


def get_purchase_entries_by_fy(client_id, fy_label):
    import pandas as pd
    return pd.read_sql_query(
        "SELECT * FROM purchase_entries WHERE client_id=? AND entry_fy=? AND status='active'",
        get_conn(), params=(client_id, fy_label)
    )


def get_purchase_entries_by_month(client_id, year_month):
    import pandas as pd
    return pd.read_sql_query(
        "SELECT * FROM purchase_entries WHERE client_id=? AND entry_month=? AND status='active'",
        get_conn(), params=(client_id, year_month)
    )


# ---------------- Reconciliation run log ----------------

def log_recon_run(client_id, snapshot_id, fy_label, purchase_asof, report_path, summary_json):
    conn = get_conn()
    conn.execute(
        """INSERT INTO recon_runs (client_id, snapshot_id, fy_label, run_at, purchase_asof, report_path, summary_json)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (client_id, snapshot_id, fy_label, now(), purchase_asof, report_path, summary_json),
    )
    conn.commit()
    persist()


def list_recon_runs(client_id):
    return get_conn().execute(
        "SELECT * FROM recon_runs WHERE client_id=? ORDER BY run_at DESC", (client_id,)
    ).fetchall()
