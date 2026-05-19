"""
Inventory Reconciliation Engine
Handles QBO (consolidated + single-company) vs TicketVault comparison.
"""

import re
import pandas as pd
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

COMPANY_MAPPING = {
    "Damona & Crew":      ["Damon and Crew"],
    "The Ticket Guy LLC": ["The Ticket Guy"],
    "Y&S Tickets":        ["YS Tickets", "YS-SeatGeek2", "YS-Seatgeek", "YS Tickets Spec"],
    "YourTickets":        ["YourTickets"],
    "YS Asher Tickets":   ["YSA", "YSA 2", "YSA 3"],
    "YS Chase Tickets":   ["Jacks YS"],
    "YS Katz Tickets":    ["YS Katz"],
    "YS Levine Tickets":  ["Yoni Levine"],
    "YS Levovitz Tickets":["Levovitz"],
    "YS Needle Tickets":  ["Needle Tickets LLC"],
    "YS TL Tickets":      ["YS TL"],
    "YSKG Tickets":       ["GK LLC"],
    "YSM Tickets":        ["YSM Tickets"],
    "YSP Tickets":        ["Pollak Tickets"],
    "YSS Tickets":        ["YSS Tickets"],
    "YSW Tickets":        ["YSW"],
}

TV_TO_QBO = {}
for _qbo, _tv_list in COMPANY_MAPPING.items():
    for _tv in _tv_list:
        TV_TO_QBO[_tv.strip().lower()] = _qbo

DESC_TO_QBO = {}
for _qbo, _tv_list in COMPANY_MAPPING.items():
    for _tv in _tv_list:
        DESC_TO_QBO[_tv.strip().lower()] = _qbo
    DESC_TO_QBO[_qbo.strip().lower()] = _qbo


def _extract_desc_company(description):
    if not description or pd.isna(description):
        return None
    m = re.search(r'\(([^)]+)\)\s*$', str(description).strip())
    return m.group(1).strip() if m else None

def _map_tv_to_qbo(tv_company):
    return TV_TO_QBO.get(str(tv_company).strip().lower())

def _map_desc_to_qbo(desc_company):
    return DESC_TO_QBO.get(str(desc_company).strip().lower())

def _clean_num(val):
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return 'N/A' if s.lower() in ('nan', 'none', '') else s

def _fmt_date(val):
    try:
        ts = pd.Timestamp(val)
        if pd.isna(ts):
            return ''
        return ts.strftime('%m/%d/%Y')
    except:
        return ''


# ---------------------------------------------------------------------------
# QBO Parsing
# ---------------------------------------------------------------------------

def parse_qbo_consolidated(filepath_or_buffer):
    raw = pd.read_excel(filepath_or_buffer, header=None)
    header_row = None
    for i, row in raw.iterrows():
        if any(str(v).strip().lower() == 'transaction date' for v in row if pd.notna(v)):
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find header row in consolidated QBO file.")

    df = pd.read_excel(filepath_or_buffer, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'company' in cl and 'transaction' not in cl: col_map['qbo_company'] = c
        elif 'transaction date' in cl: col_map['date'] = c
        elif 'transaction type' in cl: col_map['transaction_type'] = c
        elif cl == 'num': col_map['num'] = c
        elif cl == 'name': col_map['name'] = c
        elif 'description' in cl and 'full' not in cl: col_map['description'] = c
        elif 'amount' in cl and 'balance' not in cl: col_map['amount'] = c

    missing = [r for r in ['qbo_company','date','transaction_type','num','description','amount'] if r not in col_map]
    if missing:
        raise ValueError(f"Consolidated QBO missing columns: {missing}")

    df = df.rename(columns={v: k for k, v in col_map.items()})
    final_cols = ['qbo_company','transaction_type','date','num','name','description','amount']
    for c in final_cols:
        if c not in df.columns: df[c] = None
    df = df[final_cols].copy()

    df = df[df['qbo_company'].notna() & (df['qbo_company'].astype(str).str.strip() != '')]
    df = df[df['date'].notna()]
    df = df[~df['qbo_company'].astype(str).str.strip().str.lower().isin(['company','beginning balance',''])]
    # Drop rows with no transaction type (#4)
    df = df[df['transaction_type'].notna()]
    df = df[~df['transaction_type'].astype(str).str.strip().str.lower().isin(['nan','none',''])]

    df['date'] = pd.to_datetime(df['date'], errors='coerce').dt.normalize()
    df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0)
    df['num'] = df['num'].astype(str).str.strip()
    df['qbo_company'] = df['qbo_company'].astype(str).str.strip()
    df['transaction_type'] = df['transaction_type'].astype(str).str.strip()
    df['source_file'] = 'consolidated'
    return df.dropna(subset=['date'])


def parse_qbo_single(filepath_or_buffer, filename=''):
    raw = pd.read_excel(filepath_or_buffer, header=None)

    company_name = None
    for i in range(5):
        val = str(raw.iloc[i, 0]).strip()
        if val and val.lower() not in ['nan','none','','daily inventory summary']:
            company_name = val
            break
    if not company_name:
        import os
        company_name = os.path.splitext(os.path.basename(filename))[0]

    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() for v in row if pd.notna(v)]
        if 'transaction date' in vals or ('num' in vals and 'amount' in vals):
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"Could not find header row for '{company_name}'.")

    df = pd.read_excel(filepath_or_buffer, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    first_col = df.columns[0]
    has_txn_type_col = any('transaction type' in c.lower() for c in df.columns)

    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if 'transaction date' in cl: col_map['date'] = c
        elif 'transaction type' in cl: col_map['transaction_type'] = c
        elif cl == 'num': col_map['num'] = c
        elif cl == 'name': col_map['name'] = c
        elif 'description' in cl and 'full' not in cl: col_map['description'] = c
        elif 'amount' in cl and 'balance' not in cl: col_map['amount'] = c

    if 'date' not in col_map:
        raise ValueError(f"Could not find 'Transaction date' column for '{company_name}'.")

    # #5 — Infer transaction_type from col A section headers
    if not has_txn_type_col:
        txn_types = []
        current_type = 'Unknown'
        known_types = {'bill','expense','check','journal entry','credit card credit'}
        for val in df[first_col]:
            v = str(val).strip()
            if v.lower() in known_types:
                current_type = v.title()
                txn_types.append(None)  # section header row — dropped later
            else:
                txn_types.append(current_type)
        df['transaction_type'] = txn_types
        col_map['transaction_type'] = 'transaction_type'

    df['qbo_company'] = company_name
    df = df.rename(columns={v: k for k, v in col_map.items()})

    # #5 — Remove original column A (section label column)
    if first_col in df.columns and first_col not in ('transaction_type','date','num','name','description','amount','qbo_company'):
        df = df.drop(columns=[first_col])

    final_cols = ['qbo_company','transaction_type','date','num','name','description','amount']
    for c in final_cols:
        if c not in df.columns: df[c] = None
    df = df[final_cols].copy()

    df = df[df['date'].notna()]
    df = df[~df['date'].astype(str).str.strip().str.lower().isin(['nan','none','transaction date','beginning balance',''])]
    # Drop rows with no transaction type (#4)
    df = df[df['transaction_type'].notna()]
    df = df[~df['transaction_type'].astype(str).str.strip().str.lower().isin(['nan','none','','unknown'])]

    df['date'] = pd.to_datetime(df['date'], errors='coerce').dt.normalize()
    df = df.dropna(subset=['date'])
    df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0)
    df['num'] = df['num'].astype(str).str.strip()
    df['qbo_company'] = df['qbo_company'].astype(str).str.strip()
    df['transaction_type'] = df['transaction_type'].astype(str).str.strip()
    df['source_file'] = company_name
    return df


def parse_ticketvault(filepath_or_buffers):
    """
    Parse one or more TV Purchase Details files. Returns (recon_df, raw_df).
    recon_df is grouped by company+date for reconciliation.
    raw_df is the full merged cleaned dataframe (for tab display and check4).
    """
    if not isinstance(filepath_or_buffers, list):
        filepath_or_buffers = [filepath_or_buffers]

    frames = []
    for buf in filepath_or_buffers:
        df = pd.read_excel(buf)
        df.columns = [str(c).strip() for c in df.columns]
        frames.append(df)
    df = pd.concat(frames, ignore_index=True)

    if 'Cancelled' in df.columns:
        df = df[df['Cancelled'].astype(str).str.strip().str.lower() != 'yes']

    # #2 — Remove unwanted columns
    for col in ['Delivery Type', 'Notes', 'Tags']:
        if col in df.columns:
            df = df.drop(columns=[col])

    # #3 — Format all date columns as date-only
    for col in ['PO Created', 'Event Date', 'Created']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.normalize()

    df['tv_company'] = df['Company'].astype(str).str.strip()
    df['total_cost'] = pd.to_numeric(df['Total Cost'], errors='coerce').fillna(0)
    df['qbo_company'] = df['tv_company'].apply(_map_tv_to_qbo)

    raw_df = df.copy()

    recon_df = df[['tv_company','qbo_company','PO Created','total_cost']].copy()
    recon_df = recon_df.rename(columns={'PO Created': 'date'})
    recon_df = recon_df.dropna(subset=['date'])

    return recon_df, raw_df


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

def check1_daily_reconciliation(qbo_df, tv_recon_df):
    bills = qbo_df[qbo_df['transaction_type'].str.lower() == 'bill'].copy()
    qbo_grouped = bills.groupby(['qbo_company','date'])['amount'].sum().reset_index().rename(columns={'amount':'qbo_total'})
    tv_grouped = tv_recon_df.groupby(['qbo_company','date'])['total_cost'].sum().reset_index().rename(columns={'total_cost':'tv_total'})
    merged = pd.merge(qbo_grouped, tv_grouped, on=['qbo_company','date'], how='outer').fillna(0)
    merged['variance'] = merged['qbo_total'] - merged['tv_total']
    merged['match'] = merged['variance'].abs() <= 1.00

    unmapped = tv_recon_df[tv_recon_df['qbo_company'].isna()][['tv_company','date','total_cost']].copy()
    unmapped = unmapped.groupby(['tv_company','date'])['total_cost'].sum().reset_index()
    return merged.sort_values(['qbo_company','date']), unmapped


def check2_duplicate_bills(qbo_df):
    """2a: Same Company + Bill # + Amount."""
    txns = qbo_df[qbo_df['transaction_type'].str.lower().isin(['bill','expense'])].copy()
    txns = txns[txns['num'].notna() & (txns['num'] != '') & (txns['num'].str.lower() != 'nan')]
    dupes = txns[txns.duplicated(subset=['qbo_company','num'], keep=False)].copy()
    return dupes.sort_values(['qbo_company','num','date'])


def check2b_duplicate_bills_detail(qbo_df):
    """2b: Same Company + Date + Name + Description + Amount."""
    txns = qbo_df[qbo_df['transaction_type'].str.lower().isin(['bill','expense'])].copy()
    txns['_name_n'] = txns['name'].astype(str).str.strip().str.lower()
    txns['_desc_n'] = txns['description'].astype(str).str.strip().str.lower()
    dupes = txns[txns.duplicated(subset=['qbo_company','date','_name_n','_desc_n','amount'], keep=False)].copy()
    dupes = dupes.drop(columns=['_name_n','_desc_n'])
    return dupes.sort_values(['qbo_company','date','name'])


def check3_description_mismatch(qbo_df):
    relevant = qbo_df[qbo_df['transaction_type'].str.lower().isin(['bill','expense'])].copy()
    relevant['desc_company_raw'] = relevant['description'].apply(_extract_desc_company)
    relevant['desc_qbo_mapped'] = relevant['desc_company_raw'].apply(lambda x: _map_desc_to_qbo(x) if x else None)

    mismatches = relevant[
        relevant['desc_company_raw'].notna() &
        relevant['desc_qbo_mapped'].notna() &
        (relevant['desc_qbo_mapped'].str.lower() != relevant['qbo_company'].str.lower())
    ].copy()
    mismatches['mismatch_reason'] = 'Expected: ' + mismatches['desc_qbo_mapped'] + ' | Got (Company col): ' + mismatches['qbo_company']

    unmappable = relevant[relevant['desc_company_raw'].notna() & relevant['desc_qbo_mapped'].isna()].copy()
    unmappable['mismatch_reason'] = 'Description company not in mapping'

    return pd.concat([mismatches, unmappable], ignore_index=True).sort_values(['qbo_company','date'])


def check4_po_vs_created(tv_raw_df):
    """Rows where PO Created date != Created date."""
    if 'PO Created' not in tv_raw_df.columns or 'Created' not in tv_raw_df.columns:
        return pd.DataFrame()
    df = tv_raw_df.copy()
    po = pd.to_datetime(df['PO Created'], errors='coerce').dt.normalize()
    cr = pd.to_datetime(df['Created'], errors='coerce').dt.normalize()
    return df[po.notna() & cr.notna() & (po != cr)].copy()


# ---------------------------------------------------------------------------
# Excel Report Builder
# ---------------------------------------------------------------------------

RED_FILL    = PatternFill("solid", fgColor="FFCCCC")
GREEN_FILL  = PatternFill("solid", fgColor="CCFFCC")
ORANGE_FILL = PatternFill("solid", fgColor="FFE0B2")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL    = PatternFill("solid", fgColor="2F5496")
WHITE_FONT  = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
BOLD_FONT   = Font(bold=True, name="Calibri", size=10)
NORMAL_FONT = Font(name="Calibri", size=10)


def _style_header(ws, row, cols, fill=HEADER_FILL):
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = WHITE_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def _auto_width(ws, min_w=10, max_w=60):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        header_len = len(str(col[0].value)) if col[0].value else 0
        ws.column_dimensions[col_letter].width = min(max_w, max(min_w, header_len + 3, max_len + 2))


def _write_dupes_sheet(ws, df, headers, col_fns):
    if len(df) == 0:
        ws["A1"] = "No duplicates found."
        ws["A1"].font = Font(bold=True, color="006600", name="Calibri", size=12)
        return
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    _style_header(ws, 1, len(headers))
    ws.row_dimensions[1].height = 22
    amt_col = len(col_fns)  # last column is always Amount
    for r, row in enumerate(df.itertuples(), 2):
        for c, fn in enumerate(col_fns, 1):
            cell = ws.cell(r, c, fn(row))
            cell.font = NORMAL_FONT
            if c == amt_col:
                cell.number_format = '$#,##0.00'
    _auto_width(ws, max_w=60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def build_report(qbo_df, tv_recon_df, tv_raw_df=None, period_label="", input_files=None):
    wb = Workbook()

    recon_df, unmapped_tv    = check1_daily_reconciliation(qbo_df, tv_recon_df)
    dupes_df                 = check2_duplicate_bills(qbo_df)
    dupes2_df                = check2b_duplicate_bills_detail(qbo_df)
    mismatch_df              = check3_description_mismatch(qbo_df)
    po_mismatch_df           = check4_po_vs_created(tv_raw_df if tv_raw_df is not None else pd.DataFrame())

    # ── Summary tab ────────────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"
    SUMMARY_FILL  = PatternFill("solid", fgColor="EEF2FA")
    SUMMARY_ITEMS = [
        ("Tab", "What it checks", "Flag criteria"),
        ("1. Daily Reconciliation",
         "Compares daily QBO Bill totals per company against TicketVault daily Total Cost (by PO Created date).",
         "Match (green) = within $1.00  |  Missing (red) = one side is $0  |  Discrepancy (orange) = both sides have $ but differ by more than $1.00"),
        ("2. Duplicate Bills (Bill #s)",
         "Finds QBO Bills or Expenses where the same Company + Bill/Expense # appears more than once.",
         "Any Company + Bill # combination that appears 2 or more times"),
        ("3. Duplicate Bills (Detail)",
         "Finds QBO Bills or Expenses where the same Company, Date, Name, Description, and Amount all match.",
         "All five fields must match across two or more rows"),
        ("4. Description Mismatches",
         "For Bills and Expenses, checks that the company name in parentheses at the end of the Description field matches the QBO Company column.",
         "Company in parentheses maps to a different QBO company, or is not found in the mapping at all"),
        ("5. TV Date Mismatch",
         "Flags TicketVault purchase rows where the PO Created date does not match the Created date.",
         "PO Created date ≠ Created date (date portion only, time ignored)"),
    ]
    col_widths = [28, 70, 70]
    ws_sum.column_dimensions['A'].width = col_widths[0]
    ws_sum.column_dimensions['B'].width = col_widths[1]
    ws_sum.column_dimensions['C'].width = col_widths[2]
    for r, (tab, desc, criteria) in enumerate(SUMMARY_ITEMS, 1):
        for c, val in enumerate([tab, desc, criteria], 1):
            cell = ws_sum.cell(r, c, val)
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if r == 1:
                cell.fill = HEADER_FILL
                cell.font = WHITE_FONT
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                cell.fill = SUMMARY_FILL if r % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
                cell.font = BOLD_FONT if c == 1 else NORMAL_FONT
        ws_sum.row_dimensions[r].height = 60 if r > 1 else 22
    ws_sum.row_dimensions[1].height = 22

    # ── 1. Daily Reconciliation ─────────────────────────────────────────────
    ws1 = wb.create_sheet("1. Daily Reconciliation")
    headers1 = ["QBO Company", "Date", "QBO Bills Total ($)", "TV Total Cost ($)", "Variance ($)", "Status"]
    for c, h in enumerate(headers1, 1): ws1.cell(1, c, h)
    _style_header(ws1, 1, len(headers1))
    ws1.row_dimensions[1].height = 22

    for r, row in enumerate(recon_df.itertuples(), 2):
        ws1.cell(r, 1, row.qbo_company).font = NORMAL_FONT
        ws1.cell(r, 2, _fmt_date(row.date)).font = NORMAL_FONT
        ws1.cell(r, 3, round(row.qbo_total, 2)).number_format = '$#,##0.00'
        ws1.cell(r, 4, round(row.tv_total, 2)).number_format = '$#,##0.00'
        ws1.cell(r, 5, round(row.variance, 2)).number_format = '$#,##0.00'
        qbo_zero = abs(row.qbo_total) < 0.02
        tv_zero  = abs(row.tv_total)  < 0.02
        if row.match:
            label, fill = "Match", GREEN_FILL
        elif qbo_zero or tv_zero:
            label, fill = "Missing", RED_FILL
        else:
            label, fill = "Discrepancy", ORANGE_FILL
        ws1.cell(r, 6, label).font = NORMAL_FONT
        for c in range(1, 7): ws1.cell(r, c).fill = fill

    if len(unmapped_tv) > 0:
        start = len(recon_df) + 3
        ws1.cell(start, 1, "⚠️ Unmapped TicketVault Companies (no QBO mapping found)").font = Font(bold=True, color="CC6600", name="Calibri", size=10)
        ws1.merge_cells(f"A{start}:F{start}")
        for c, h in enumerate(["TV Company","Date","TV Total Cost ($)"], 1): ws1.cell(start+1, c, h)
        _style_header(ws1, start+1, 3, fill=SUB_FILL)
        for r2, row2 in enumerate(unmapped_tv.itertuples(), start+2):
            ws1.cell(r2, 1, row2.tv_company).font = NORMAL_FONT
            ws1.cell(r2, 2, _fmt_date(row2.date)).font = NORMAL_FONT
            ws1.cell(r2, 3, round(row2.total_cost, 2)).number_format = '$#,##0.00'
            for c in range(1, 4): ws1.cell(r2, c).fill = ORANGE_FILL

    _auto_width(ws1)
    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = f"A1:F{len(recon_df)+1}"

    # ── 2. Duplicate Bills (2a: Company + Bill # + Amount) ──────────────────
    ws2 = wb.create_sheet("2. Duplicate Bills (Bill #s)")
    h2 = ["QBO Company","Transaction Type","Bill / Expense #","Date","Name","Description","Amount ($)"]
    _write_dupes_sheet(ws2, dupes_df, h2, [
        lambda r: r.qbo_company,
        lambda r: r.transaction_type,
        lambda r: _clean_num(r.num),
        lambda r: _fmt_date(r.date),
        lambda r: str(r.name) if pd.notna(r.name) else '',
        lambda r: str(r.description) if pd.notna(r.description) else '',
        lambda r: round(r.amount, 2),
    ])
    if len(dupes_df) > 0:
        for r2 in range(2, ws2.max_row+1):
            ws2.cell(r2, 7).number_format = '$#,##0.00'

    # ── 3. Duplicate Bills Detail (2b: Company + Date + Name + Desc + Amount) ─
    ws3 = wb.create_sheet("3. Duplicate Bills (Detail)")
    h3 = ["QBO Company","Transaction Type","Date","Bill / Expense #","Name","Description","Amount ($)"]
    _write_dupes_sheet(ws3, dupes2_df, h3, [
        lambda r: r.qbo_company,
        lambda r: r.transaction_type,
        lambda r: _fmt_date(r.date),
        lambda r: _clean_num(r.num),
        lambda r: str(r.name) if pd.notna(r.name) else '',
        lambda r: str(r.description) if pd.notna(r.description) else '',
        lambda r: round(r.amount, 2),
    ])

    # ── 4. Description Mismatches ───────────────────────────────────────────
    ws4 = wb.create_sheet("4. Description Mismatches")
    if len(mismatch_df) == 0:
        ws4["A1"] = "No description company mismatches found."
        ws4["A1"].font = Font(bold=True, color="006600", name="Calibri", size=12)
    else:
        h4 = ["QBO Company","Transaction Type","Date","Bill / Expense #","Description","Mismatch Detail","Amount ($)"]
        for c, h in enumerate(h4, 1): ws4.cell(1, c, h)
        _style_header(ws4, 1, len(h4))
        ws4.row_dimensions[1].height = 22
        for r, row in enumerate(mismatch_df.itertuples(), 2):
            ws4.cell(r, 1, row.qbo_company).font = NORMAL_FONT
            ws4.cell(r, 2, row.transaction_type).font = NORMAL_FONT
            ws4.cell(r, 3, _fmt_date(row.date)).font = NORMAL_FONT
            ws4.cell(r, 4, _clean_num(row.num)).font = NORMAL_FONT
            ws4.cell(r, 5, str(row.description) if pd.notna(row.description) else '').font = NORMAL_FONT
            ws4.cell(r, 6, str(row.mismatch_reason)).font = NORMAL_FONT
            ws4.cell(r, 7, round(row.amount, 2)).number_format = '$#,##0.00'
    _auto_width(ws4, max_w=60)
    ws4.freeze_panes = "A2"
    if len(mismatch_df) > 0: ws4.auto_filter.ref = ws4.dimensions

    # ── 5. TV PO Created vs Created Date Mismatch ───────────────────────────
    ws5 = wb.create_sheet("5. TV Date Mismatch")
    DATE_COLS = {'PO Created','Event Date','Created'}
    MONEY_COLS = {'Cost','Total Cost'}
    if len(po_mismatch_df) == 0:
        ws5["A1"] = "No PO Created vs Created date mismatches found."
        ws5["A1"].font = Font(bold=True, color="006600", name="Calibri", size=12)
    else:
        tv_cols = [c for c in po_mismatch_df.columns if not c.startswith('_') and c not in ('tv_company','total_cost','qbo_company')]
        for c, h in enumerate(tv_cols, 1): ws5.cell(1, c, h)
        _style_header(ws5, 1, len(tv_cols))
        ws5.row_dimensions[1].height = 22
        for r, row in enumerate(po_mismatch_df[tv_cols].itertuples(index=False), 2):
            for c, (col_name, val) in enumerate(zip(tv_cols, row), 1):
                if col_name in DATE_COLS:
                    ws5.cell(r, c, _fmt_date(val)).font = NORMAL_FONT
                elif col_name in MONEY_COLS:
                    cell = ws5.cell(r, c, val)
                    cell.number_format = '$#,##0.00'
                    cell.font = NORMAL_FONT
                else:
                    ws5.cell(r, c, val if val != '' else None).font = NORMAL_FONT
    # Header-only column widths for TV date mismatch tab
    for col in ws5.columns:
        col_letter = get_column_letter(col[0].column)
        header_len = len(str(col[0].value)) if col[0].value else 10
        ws5.column_dimensions[col_letter].width = header_len + 3
    ws5.freeze_panes = "A2"
    if len(po_mismatch_df) > 0: ws5.auto_filter.ref = ws5.dimensions

    # ── Input file tabs ─────────────────────────────────────────────────────
    if input_files:
        DATE_COLS_TV = {'PO Created', 'Event Date', 'Created'}
        MONEY_COLS_TV = {'Cost', 'Total Cost'}
        REMOVE_COLS_TV = {'Delivery Type', 'Notes', 'Tags'}

        tv_input_files = [(f, b) for f, b in input_files if 'purchase' in f.lower() or 'ticketvault' in f.lower()]
        qbo_input_files = [(f, b) for f, b in input_files if f not in [x[0] for x in tv_input_files]]

        # ── Merged TV Purchases tab ──────────────────────────────────────────
        if tv_input_files:
            ws_tv = wb.create_sheet("TV - Purchases")
            tv_all_rows = []
            tv_headers = None
            for fname, fbytes in tv_input_files:
                try:
                    src_wb = load_workbook(BytesIO(fbytes), read_only=True, data_only=True)
                    src_ws = src_wb.active
                    rows = list(src_ws.iter_rows(values_only=True))
                    if not rows:
                        src_wb.close()
                        continue
                    # Find which column indices to keep (remove Delivery Type, Notes, Tags)
                    header = [str(v).strip() if v is not None else '' for v in rows[0]]
                    keep_idx = [i for i, h in enumerate(header) if h not in REMOVE_COLS_TV]
                    filtered_header = [header[i] for i in keep_idx]
                    if tv_headers is None:
                        tv_headers = filtered_header
                    for row_vals in rows[1:]:
                        tv_all_rows.append([row_vals[i] if i < len(row_vals) else None for i in keep_idx])
                    src_wb.close()
                except Exception as e:
                    pass

            if tv_headers:
                for c, h in enumerate(tv_headers, 1):
                    ws_tv.cell(1, c, h)
                _style_header(ws_tv, 1, len(tv_headers))
                ws_tv.row_dimensions[1].height = 22
                for r, row_vals in enumerate(tv_all_rows, 2):
                    for c, (col_name, val) in enumerate(zip(tv_headers, row_vals), 1):
                        if col_name in DATE_COLS_TV:
                            ws_tv.cell(r, c, _fmt_date(val)).font = NORMAL_FONT
                        elif col_name in MONEY_COLS_TV:
                            cell = ws_tv.cell(r, c, val)
                            cell.number_format = '$#,##0.00'
                            cell.font = NORMAL_FONT
                        else:
                            ws_tv.cell(r, c, val if val is not None else '').font = NORMAL_FONT
            # Header-only column widths for TV tab
            for col in ws_tv.columns:
                col_letter = get_column_letter(col[0].column)
                header_len = len(str(col[0].value)) if col[0].value else 10
                ws_tv.column_dimensions[col_letter].width = header_len + 3
            ws_tv.freeze_panes = "A2"
            ws_tv.auto_filter.ref = ws_tv.dimensions

        # ── QBO Bills + Expenses tabs (no Journal Entries) ─────────────────
        QBO_OUT_HEADERS = ["Company", "Transaction Type", "Transaction Date", "Num", "Name", "Description", "Amount ($)"]

        def _write_qbo_tab(ws, df_subset):
            for c, h in enumerate(QBO_OUT_HEADERS, 1):
                ws.cell(1, c, h)
            _style_header(ws, 1, len(QBO_OUT_HEADERS))
            ws.row_dimensions[1].height = 22
            for r, row in enumerate(df_subset.itertuples(), 2):
                ws.cell(r, 1, row.qbo_company).font = NORMAL_FONT
                ws.cell(r, 2, row.transaction_type).font = NORMAL_FONT
                ws.cell(r, 3, _fmt_date(row.date)).font = NORMAL_FONT
                ws.cell(r, 4, _clean_num(row.num)).font = NORMAL_FONT
                ws.cell(r, 5, str(row.name) if pd.notna(row.name) else '').font = NORMAL_FONT
                ws.cell(r, 6, str(row.description) if pd.notna(row.description) else '').font = NORMAL_FONT
                cell = ws.cell(r, 7, round(row.amount, 2))
                cell.number_format = '$#,##0.00'
                cell.font = NORMAL_FONT
            _auto_width(ws, max_w=60)
            ws.freeze_panes = "A2"
            if len(df_subset) > 0:
                ws.auto_filter.ref = ws.dimensions

        qbo_excl_je = qbo_df[~qbo_df['transaction_type'].str.lower().isin(['journal entry'])]
        qbo_bills    = qbo_excl_je[qbo_excl_je['transaction_type'].str.lower() == 'bill']
        qbo_expenses = qbo_excl_je[qbo_excl_je['transaction_type'].str.lower() != 'bill']

        ws_qbo_b = wb.create_sheet("QBO - Bills")
        _write_qbo_tab(ws_qbo_b, qbo_bills)

        ws_qbo_e = wb.create_sheet("QBO - Expenses")
        _write_qbo_tab(ws_qbo_e, qbo_expenses)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out
