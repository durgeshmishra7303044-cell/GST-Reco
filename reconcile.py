import pandas as pd
import re
from datetime import datetime


# ── helpers ──────────────────────────────────────────────────────────────────

def normalize_inv(s):
    """Remove spaces, slashes, dashes, uppercase → loose match."""
    return re.sub(r'[\s/\-]', '', str(s)).upper().strip()

def normalize_date(val):
    """Try common date formats → YYYY-MM-DD string, or original string."""
    if pd.isna(val) or str(val).strip() in ('', 'nan'):
        return ''
    s = str(val).strip()
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%d-%b-%Y', '%d %b %Y',
                '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%y', '%d-%m-%y'):
        try:
            return datetime.strptime(s[:11], fmt).strftime('%Y-%m-%d')
        except Exception:
            pass
    # pandas timestamp
    try:
        return pd.to_datetime(val).strftime('%Y-%m-%d')
    except Exception:
        pass
    return s

def make_key_exact(gstin, inv_no, date):
    return f"{str(gstin).strip()}||{str(inv_no).strip()}||{normalize_date(date)}"

def make_key_loose(gstin, inv_no, date):
    return f"{str(gstin).strip()}||{normalize_inv(inv_no)}||{normalize_date(date)}"

def parse_num(val):
    try:
        return float(str(val).replace(',', '').replace('₹', '').strip())
    except Exception:
        return 0.0


# ── column detection ──────────────────────────────────────────────────────────

GSTR2B_COL_HINTS = {
    'gstin':     ['gstin of supplier', 'supplier gstin', 'gstin'],
    'inv_no':    ['invoice number', 'inv no', 'invoice no', 'document number'],
    'inv_date':  ['invoice date', 'inv date', 'document date'],
    'trade_name':['trade/legal name', 'trade name', 'supplier name', 'party name'],
    'taxable':   ['taxable value', 'taxable val'],
    'igst':      ['integrated tax', 'igst'],
    'cgst':      ['central tax', 'cgst'],
    'sgst':      ['state/ut tax', 'sgst', 'utgst'],
    'cess':      ['cess'],
    'inv_type':  ['invoice type', 'inv type'],
    'itc_avail': ['itc availability', 'itc avail'],
}

BOOKS_COL_HINTS = {
    'gstin':         ['gstin', 'supplier gstin', 'party gstin'],
    'party_inv_no':  ['party invoice no', 'party inv no', 'vendor invoice', 'supplier invoice', 'invoice no', 'inv no'],
    'party_inv_date':['party invoice date', 'party inv date', 'invoice date', 'inv date'],
    'party':         ['party name', 'supplier name', 'vendor name', 'party'],
    'doc_no':        ['document no', 'doc no', 'voucher no', 'entry no'],
    'total':         ['total amount', 'gross amount', 'total'],
    'net':           ['net amount', 'taxable', 'net'],
    'cgst':          ['cgst'],
    'sgst':          ['sgst', 'utgst'],
    'igst':          ['igst'],
    'division':      ['division', 'branch', 'location', 'cost center'],
}

def detect_col(df, hints):
    """Return {logical_name: actual_col} by fuzzy header matching."""
    headers = {c.lower().strip(): c for c in df.columns}
    mapping = {}
    for logical, candidates in hints.items():
        for cand in candidates:
            if cand in headers:
                mapping[logical] = headers[cand]
                break
    return mapping


# ── loaders ───────────────────────────────────────────────────────────────────

def load_gstr2b(path):
    """Load GSTR-2B Excel. Handles multi-row headers."""
    raw = pd.read_excel(path, header=None, dtype=str)

    # Find the header row — look for a row containing 'GSTIN'
    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v).lower().strip() for v in row if not pd.isna(v)]
        if any('gstin' in v for v in vals):
            header_row = i
            break
    if header_row is None:
        header_row = 0

    df = pd.read_excel(path, header=header_row, dtype=str)
    df.columns = [str(c).lower().strip() for c in df.columns]
    df = df.dropna(how='all')
    col_map = detect_col(df, GSTR2B_COL_HINTS)
    return df, col_map


def load_books(path):
    """Load Purchase Register Excel."""
    raw = pd.read_excel(path, header=None, dtype=str)

    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v).lower().strip() for v in row if not pd.isna(v)]
        if any(k in ' '.join(vals) for k in ['party', 'invoice', 'gstin', 'voucher']):
            header_row = i
            break
    if header_row is None:
        header_row = 0

    df = pd.read_excel(path, header=header_row, dtype=str)
    df.columns = [str(c).lower().strip() for c in df.columns]
    df = df.dropna(how='all')
    col_map = detect_col(df, BOOKS_COL_HINTS)
    return df, col_map


# ── main reconciliation ───────────────────────────────────────────────────────

def reconcile(gstr2b_path, books_path):
    df2b, cm2b = load_gstr2b(gstr2b_path)
    dfbk, cmbk = load_books(books_path)

    # ── Build 2B working frame ────────────────────────────────────────────────
    g = df2b.copy()
    g['_gstin']    = g[cm2b.get('gstin', df2b.columns[0])].fillna('').str.strip()
    g['_inv_no']   = g[cm2b.get('inv_no', df2b.columns[1])].fillna('').str.strip()
    g['_inv_date'] = g[cm2b.get('inv_date', df2b.columns[2])].fillna('').apply(normalize_date)
    g['_name']     = g[cm2b.get('trade_name', df2b.columns[3])].fillna('').str.strip() if 'trade_name' in cm2b else ''
    g['_taxable']  = g[cm2b.get('taxable', '')].apply(parse_num) if 'taxable' in cm2b else 0
    g['_igst']     = g[cm2b.get('igst', '')].apply(parse_num) if 'igst' in cm2b else 0
    g['_cgst']     = g[cm2b.get('cgst', '')].apply(parse_num) if 'cgst' in cm2b else 0
    g['_sgst']     = g[cm2b.get('sgst', '')].apply(parse_num) if 'sgst' in cm2b else 0
    g['_total_itc'] = g['_igst'] + g['_cgst'] + g['_sgst']
    g['_key_exact'] = g.apply(lambda r: make_key_exact(r['_gstin'], r['_inv_no'], r['_inv_date']), axis=1)
    g['_key_loose'] = g.apply(lambda r: make_key_loose(r['_gstin'], r['_inv_no'], r['_inv_date']), axis=1)

    # ── Build Books working frame ─────────────────────────────────────────────
    b = dfbk.copy()
    b['_gstin']    = b[cmbk.get('gstin', dfbk.columns[0])].fillna('').str.strip() if 'gstin' in cmbk else ''
    b['_inv_no']   = b[cmbk.get('party_inv_no', dfbk.columns[1])].fillna('').str.strip() if 'party_inv_no' in cmbk else ''
    b['_inv_date'] = b[cmbk.get('party_inv_date', dfbk.columns[2])].fillna('').apply(normalize_date) if 'party_inv_date' in cmbk else ''
    b['_party']    = b[cmbk.get('party', '')].fillna('').str.strip() if 'party' in cmbk else ''
    b['_doc_no']   = b[cmbk.get('doc_no', '')].fillna('').str.strip() if 'doc_no' in cmbk else ''
    b['_division'] = b[cmbk.get('division', '')].fillna('').str.strip() if 'division' in cmbk else ''
    b['_igst']     = b[cmbk.get('igst', '')].apply(parse_num) if 'igst' in cmbk else 0
    b['_cgst']     = b[cmbk.get('cgst', '')].apply(parse_num) if 'cgst' in cmbk else 0
    b['_sgst']     = b[cmbk.get('sgst', '')].apply(parse_num) if 'sgst' in cmbk else 0
    b['_total_itc'] = b['_igst'] + b['_cgst'] + b['_sgst']
    b['_key_exact'] = b.apply(lambda r: make_key_exact(r['_gstin'], r['_inv_no'], r['_inv_date']), axis=1)
    b['_key_loose'] = b.apply(lambda r: make_key_loose(r['_gstin'], r['_inv_no'], r['_inv_date']), axis=1)

    books_exact = set(b['_key_exact'])
    books_loose = set(b['_key_loose'])
    gstr_exact  = set(g['_key_exact'])
    gstr_loose  = set(g['_key_loose'])

    # ── Classify 2B rows ──────────────────────────────────────────────────────
    def classify_2b(row):
        if row['_key_exact'] in books_exact:
            return 'Matched'
        elif row['_key_loose'] in books_loose:
            return 'Format Mismatch'
        else:
            return '2B Not In Books'

    g['_status'] = g.apply(classify_2b, axis=1)

    # ── Classify Books rows ───────────────────────────────────────────────────
    def classify_books(row):
        if row['_key_exact'] in gstr_exact:
            return 'Matched'
        elif row['_key_loose'] in gstr_loose:
            return 'Format Mismatch'
        else:
            return 'Books Not In 2B'

    b['_status'] = b.apply(classify_books, axis=1)

    # ── Split into result frames ───────────────────────────────────────────────
    matched_2b        = g[g['_status'] == 'Matched']
    fmt_mismatch_2b   = g[g['_status'] == 'Format Mismatch']
    not_in_books      = g[g['_status'] == '2B Not In Books']
    not_in_2b         = b[b['_status'] == 'Books Not In 2B']

    # ── Summary stats ─────────────────────────────────────────────────────────
    summary = {
        'total_2b':              len(g),
        'total_books':           len(b),
        'matched':               len(matched_2b),
        'format_mismatch':       len(fmt_mismatch_2b),
        'not_in_books_count':    len(not_in_books),
        'not_in_books_itc':      not_in_books['_total_itc'].sum(),
        'not_in_2b_count':       len(not_in_2b),
        'not_in_2b_itc':         not_in_2b['_total_itc'].sum(),
    }

    return {
        'summary':          summary,
        'matched':          matched_2b,
        'fmt_mismatch':     fmt_mismatch_2b,
        'not_in_books':     not_in_books,
        'not_in_2b':        not_in_2b,
        'all_2b':           g,
        'all_books':        b,
    }
