"""
AnH (Agency & Health) Grid Processor
Converts the Agency Health Excel grid to the target magma_anh_monthly CSV format.

Key transformation rules:
  1. Each source row with a Fresh % → one output row  (type_of_business = New)
  2. Each source row with a Renewal % → one output row (type_of_business = Renewal)
  3. If the renewal60 column has a value → an extra row with insured_age_ll=61
  4. Product Code/Name → resolved via Product Master sheet
  5. Vol/NBP/Slab → mapped to Final Nm GWP Ll/Ul
  6. 'Common grid' rows → use pol_agent_code = ANY; outgo = GWP
  7. Specific rate rows → use pol_agent_code = IMD Code; outgo = RGWP
  8. Rates in % string format (e.g. '33.90%') are cleaned to float
"""

from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import openpyxl
import os, uuid, traceback, re
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED = {'xlsx', 'xls', 'xlsb', 'csv'}
def ok_file(f): return '.' in f and f.rsplit('.',1)[1].lower() in ALLOWED


# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def get_sheets(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    s = wb.sheetnames; wb.close(); return s


def clean_rate(val):
    """
    Convert rate to float percentage suitable for the output comm_percentage column.
    Handles: 0.3814 (decimal) → 38.14, '33.90%' (percent string) → 33.90,
             '-' / 'Common grid' → None
    """
    if val is None: return None
    s = str(val).strip()
    if s in ('-', 'Common grid', '', 'None', '0'): return None
    # Remove % sign
    s_clean = s.replace('%', '').split(':')[0].strip()  # handle '19.07% :- on ...'
    try:
        f = float(s_clean)
        # If value < 1 it's a decimal fraction → multiply by 100
        if f < 1.0:
            f = round(f * 100, 6)
        return round(f, 4)
    except (ValueError, TypeError):
        return None


def extract_numeric_codes(val):
    """Return list of 4+ digit code strings found in val, preserving order."""
    if val is None:
        return []
    return re.findall(r'\b(\d{4,})\b', str(val))


def resolve_product_codes(val, all_codes):
    """
    Resolve a raw Product Master code string to a final comma-separated
    code string, using all_codes (ordered list) as the full universe.

    Cases:
      'Arogya Sanjeevani (6116), ..., IPA (8101)'  -> codes extracted directly
      'Excluding 8101, 8103'  -> all_codes minus [8101, 8103]
      '8101, 8103'            -> '8101, 8103'
      '6119'                  -> '6119'
      8103  (int)             -> '8103'
    """
    if val is None:
        return None
    s = str(val).strip()
    nums = extract_numeric_codes(s)

    if re.search(r'\bexcluding\b', s, re.IGNORECASE):
        # Subtract excluded codes from the full universe, preserving order
        exclude_set = set(nums)
        remaining = [c for c in all_codes if c not in exclude_set]
        return ', '.join(remaining) if remaining else s

    if nums:
        return ', '.join(nums)

    return s  # fallback: no digits found, return as-is


def is_common_grid(val):
    """Return True if the rate value means 'use common grid rate'."""
    if val is None: return True
    return str(val).strip() in ('Common grid', '-', '', 'None')


# Product codes that are exclusively covered by the renewal60 (age > 60) row.
# The column header itself says: "If the insured or member age is > 60 and
# product code is 6111, 6117, 6118".  When a renewal60 rate is present these
# codes must be REMOVED from the regular renewal row to prevent overlap.
RENEWAL60_EXCLUSIVE_CODES = {'6111', '6117', '6118'}


def exclude_codes_from_product_string(product_code_str, codes_to_exclude):
    """
    Remove specific product codes from a comma-separated product code string.
    Returns the filtered string, or None if no codes remain.

    Example:
        exclude_codes_from_product_string('6116, 6111, 6117, 6118, 6119', {'6111','6117','6118'})
        → '6116, 6119'
    """
    if not product_code_str:
        return product_code_str
    codes = [c.strip() for c in str(product_code_str).split(',')]
    filtered = [c for c in codes if c not in codes_to_exclude]
    return ', '.join(filtered) if filtered else None


VOL_SLAB_MAP = {
    '-':                (-1e10, 1e11),
    '':                 (-1e10, 1e11),
    '<=50 K':           (-1e10, 49999),
    '>50 K':            (50000, 1e11),
    '>50K - 1.5 Lac':   (50000, 149999),
    '>1.5 Lac - 2 lac': (150000, 199999),
    '>2 Lac':           (200000, 1e11),
    'Upto 199999':      (-1e10, 199999),
    '200000 & Above':   (200000, 1e11),
}

def vol_to_gwp(vol_str):
    """Map Vol/NBP/Slab string to (gwp_ll, gwp_ul)."""
    s = str(vol_str).strip() if vol_str else '-'
    return VOL_SLAB_MAP.get(s, (-1e10, 1e11))


# ══════════════════════════════════════════════════════════════
# PRODUCT MASTER LOADING
# ══════════════════════════════════════════════════════════════

def load_product_master(wb, sheet_name='Product Master', col_name=1, col_code=2):
    """
    Load the Product Master sheet and return a dict:
      { product_name: resolved_product_code_string }

    Resolution logic:
      1. Find the full universe of codes from the row with the most
         numeric codes (typically the 'All Non-Motor Product' row).
      2. For every row, resolve_product_codes() either extracts codes
         directly or subtracts 'Excluding' codes from the full universe.
         e.g. 'Excluding 8101, 8103' -> '6116, 6111, 6117, 6118, 6119, 9993, 9997'
    """
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]

    # Pass 1: collect raw rows
    raw = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, col_name).value
        code = ws.cell(r, col_code).value
        if name and code:
            raw.append((str(name).strip(), str(code).strip()))

    # Determine full code universe: row with the most numeric codes
    all_codes = []
    for _, code_val in raw:
        nums = extract_numeric_codes(code_val)
        if len(nums) > len(all_codes):
            all_codes = nums

    # Pass 2: resolve each entry against the full universe
    pm = {}
    for name, code_val in raw:
        pm[name] = resolve_product_codes(code_val, all_codes)
    return pm


# ══════════════════════════════════════════════════════════════
# CORE PROCESSOR
# ══════════════════════════════════════════════════════════════

def process_anh(config):
    """
    Main AnH processing engine.

    config keys:
      filepath                  Path to the Excel file
      imd_sheet                 Sheet name for the IMD List (default: 'IMD List')
      product_master_sheet      Sheet for product codes (default: 'Product Master')
      header_row                1-indexed row with column headers (default: 2)
      data_start_row            1-indexed first data row (default: 3)

      -- Column positions (1-indexed, 0 = absent) --
      col_rel_code              Relationship Code (default: 2)
      col_imd_code              IMD Code (default: 3)
      col_imd_name              IMD Name (default: 4)
      col_product_name          Product Code/Name (default: 5)
      col_80_10_10              80.10.10 flag (default: 6)
      col_vol_slab              Vol/NBP/Slab (default: 7)
      col_fresh_pct             Fresh Business % (default: 8)
      col_rollover_pct          Rollover payout % (default: 9)
      col_renewal_pct           Renewal Business % (default: 10)
      col_renewal60_pct         Renewal % for age>60 (default: 11)
      col_fresh_rnr             Fresh Business RNR % (default: 12)
      col_renewal_rnr           Renewal Business RNR % (default: 13)
      col_status                Status column (default: 17)

      -- Filters --
      active_statuses           List of status values to include (default: ['active'])
      skip_common_grid_fresh    If True, skip rows where fresh = 'Common grid' (default: False)
      skip_common_grid_renewal  If True, skip rows where renewal = 'Common grid' (default: False)
      common_grid_agent         pol_agent_code for 'Common grid' rows (default: 'ANY')
      common_grid_outgo         outgo for 'Common grid' rows (default: 'GWP')
      specific_rate_outgo       outgo for specific rate rows (default: 'RGWP')

      -- Output static fields --
      version_id                Version Id* value (default: 'magma_anh_monthly_Grid_4')
      default_insured_age_ll    Default insured age lower (default: 0.0)
      default_insured_age_ul    Default insured age upper (default: 1000.0)
      renewal60_age_ll          Insured age ll for renewal60 rows (default: 61.0)
      renewal60_age_ul          Insured age ul for renewal60 rows (default: 1000.0)
      default_tsi_ll            Total sum insured LL (default: -1e10)
      default_tsi_ul            Total sum insured UL (default: 1e11)
      policy_year               Policy year value (default: 'ANY')
      output_name               Base name for output file

      -- Vol/Slab mapping override --
      vol_slab_map              Dict of { slab_string: [ll, ul] } to override defaults

      -- Product master columns --
      pm_col_name               Product Master name column index (default: 1)
      pm_col_code               Product Master code column index (default: 2)
    """
    path                 = config['filepath']
    imd_sheet            = config.get('imd_sheet', 'IMD List')
    pm_sheet             = config.get('product_master_sheet', 'Product Master')
    header_row           = config.get('header_row', 2)
    data_start           = config.get('data_start_row', 3)

    c_rel    = config.get('col_rel_code', 2)
    c_imd    = config.get('col_imd_code', 3)
    c_name   = config.get('col_imd_name', 4)
    c_prod   = config.get('col_product_name', 5)
    c_8010   = config.get('col_80_10_10', 6)
    c_vol    = config.get('col_vol_slab', 7)
    c_fresh  = config.get('col_fresh_pct', 8)
    c_roll   = config.get('col_rollover_pct', 9)
    c_renew  = config.get('col_renewal_pct', 10)
    c_r60    = config.get('col_renewal60_pct', 11)
    c_status = config.get('col_status', 17)

    active_statuses      = [s.lower() for s in config.get('active_statuses', ['active'])]
    skip_cg_fresh        = config.get('skip_common_grid_fresh', False)
    skip_cg_renewal      = config.get('skip_common_grid_renewal', False)
    cg_agent             = config.get('common_grid_agent', 'ANY')
    cg_outgo             = config.get('common_grid_outgo', 'GWP')
    spec_outgo           = config.get('specific_rate_outgo', 'RGWP')

    version_id           = config.get('version_id', 'magma_anh_monthly_Grid_4')
    def_age_ll           = config.get('default_insured_age_ll', 0.0)
    def_age_ul           = config.get('default_insured_age_ul', 1000.0)
    r60_age_ll           = config.get('renewal60_age_ll', 61.0)
    r60_age_ul           = config.get('renewal60_age_ul', 1000.0)
    def_tsi_ll           = config.get('default_tsi_ll', -1e10)
    def_tsi_ul           = config.get('default_tsi_ul', 1e11)
    policy_year          = config.get('policy_year', 'ANY')

    # Vol/slab map override
    vol_map = dict(VOL_SLAB_MAP)
    for k, v in config.get('vol_slab_map', {}).items():
        vol_map[str(k)] = (float(v[0]), float(v[1]))

    pm_col_name = config.get('pm_col_name', 1)
    pm_col_code = config.get('pm_col_code', 2)

    wb = openpyxl.load_workbook(path, data_only=True)

    # Load product master
    product_master = load_product_master(wb, pm_sheet, pm_col_name, pm_col_code)

    if imd_sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{imd_sheet}' not found. Available: {wb.sheetnames}")
    ws = wb[imd_sheet]
    max_row = ws.max_row

    def cell(r, c):
        if not c or c <= 0: return None
        v = ws.cell(r, c).value
        return v

    def cell_str(r, c):
        v = cell(r, c)
        return str(v).strip() if v is not None else ''

    rows = []
    uid  = 1
    skipped = 0

    for row_idx in range(data_start, max_row + 1):
        # Skip by status
        if c_status and c_status > 0:
            status = cell_str(row_idx, c_status).lower()
            if status and status not in active_statuses:
                skipped += 1; continue

        rel_code    = cell_str(row_idx, c_rel)
        imd_code    = cell_str(row_idx, c_imd)
        prod_name   = cell_str(row_idx, c_prod)
        vol_slab    = cell_str(row_idx, c_vol)
        flag_8010   = cell_str(row_idx, c_8010)

        # Skip truly blank rows
        if not imd_code and not rel_code:
            skipped += 1; continue

        # Product code lookup (already resolved in load_product_master)
        product_code = product_master.get(prod_name, resolve_product_codes(prod_name, []))

        # GWP ll/ul from vol slab
        gwp_ll, gwp_ul = vol_map.get(vol_slab if vol_slab else '-', (-1e10, 1e11))

        # Raw rate values
        fresh_raw   = cell(row_idx, c_fresh)
        renewal_raw = cell(row_idx, c_renew)
        r60_raw     = cell(row_idx, c_r60)

        fresh_rate   = clean_rate(fresh_raw)
        renewal_rate = clean_rate(renewal_raw)
        r60_rate     = clean_rate(r60_raw)

        cg_fresh   = is_common_grid(fresh_raw)
        cg_renewal = is_common_grid(renewal_raw)

        # Skip if both rates are absent / common grid
        if cg_fresh and cg_renewal:
            skipped += 1; continue
        if cg_fresh and skip_cg_fresh and cg_renewal:
            skipped += 1; continue

        # Determine pol_agent_code: always use IMD Code
        agent_code = imd_code

        def _make_row(tob, rate, age_ll, age_ul, tsi_ll=None, tsi_ul=None, outgo=None,
                      product_code_override=None):
            nonlocal uid
            # If rate is None or common grid → use common grid logic
            is_cg = (rate is None)
            pol_code = cg_agent if is_cg else agent_code
            out_go   = outgo if outgo else (cg_outgo if is_cg else spec_outgo)
            r = {
                'Version Id*':          version_id,
                'Pol Agent Code*':      pol_code,
                'Product Code*':        product_code_override if product_code_override is not None else product_code,
                'Final Nm Gwp Ll*':     gwp_ll,
                'Final Nm Gwp Ul*':     gwp_ul,
                'Type Of Business*':    tob,
                'Insured Age Ll*':      age_ll,
                'Insured Age Ul*':      age_ul,
                'Outgo*':               out_go,
                'Comm Percentage*':     rate if rate is not None else '',
                'Total Sum Insured Ll*': tsi_ll if tsi_ll is not None else def_tsi_ll,
                'Total Sum Insured Ul*': tsi_ul if tsi_ul is not None else def_tsi_ul,
                'Policy Year*':         policy_year,
                'Unique Id':            uid,
            }
            uid += 1
            return r

        # ── Fresh Business row ──────────────────────────────────────────────
        if not cg_fresh or not skip_cg_fresh:
            if fresh_rate is not None:
                rows.append(_make_row('New', fresh_rate, def_age_ll, def_age_ul))

        # ── Renewal Business rows ───────────────────────────────────────────
        # When a renewal60 rate exists for codes 6111/6117/6118, the renewal
        # rows split into:
        #   A) 6111, 6117, 6118  | age 0-60    | renewal_rate
        #   B) other codes       | age 0-1000  | renewal_rate
        #   C) 6111, 6117, 6118  | age 61-1000 | r60_rate       (r60 row below)
        # When no renewal60 split applies, a single row covers all codes.
        if not cg_renewal or not skip_cg_renewal:
            if renewal_rate is not None:
                row_codes_set = {c.strip() for c in str(product_code).split(',')} if product_code else set()
                r60_present = r60_rate is not None and r60_rate != renewal_rate
                overlap_codes = ', '.join(c for c in ['6111', '6117', '6118'] if c in row_codes_set)

                if r60_present and overlap_codes:
                    # A) 6111/6117/6118 at renewal rate, age 0 to 60
                    rows.append(_make_row('Renewal', renewal_rate, def_age_ll, r60_age_ll - 1,
                                         product_code_override=overlap_codes))
                    # B) remaining codes at renewal rate, full age range
                    other_codes = exclude_codes_from_product_string(product_code, RENEWAL60_EXCLUSIVE_CODES)
                    if other_codes:
                        rows.append(_make_row('Renewal', renewal_rate, def_age_ll, def_age_ul,
                                             product_code_override=other_codes))
                else:
                    # No r60 split — single renewal row for all codes
                    rows.append(_make_row('Renewal', renewal_rate, def_age_ll, def_age_ul))

        # ── Renewal age > 60 row ────────────────────────────────────────────
        # C) Only codes 6111, 6117, 6118 (as per column header), age 61-1000.
        #    Only when renewal60 col has a rate different from the renewal rate.
        if r60_rate is not None and r60_rate != renewal_rate:
            row_codes_set = {c.strip() for c in str(product_code).split(',')} if product_code else set()
            r60_codes = ', '.join(c for c in ['6111', '6117', '6118'] if c in row_codes_set)
            if r60_codes:
                rows.append(_make_row('Renewal', r60_rate, r60_age_ll, r60_age_ul,
                                      product_code_override=r60_codes))

    wb.close()
    return pd.DataFrame(rows), skipped, uid - 1


def process_policy_wise(config, wb):
    """Process the Policy Wise sheet into additional output rows."""
    sheet = config.get('policy_wise_sheet', 'Policy Wise')
    if not wb or sheet not in wb.sheetnames:
        return pd.DataFrame()

    ws  = wb[sheet]
    ver = config.get('version_id', 'magma_anh_monthly_Grid_4')
    def_tsi_ll = config.get('default_tsi_ll', -1e10)
    def_tsi_ul = config.get('default_tsi_ul', 1e11)
    policy_year = config.get('policy_year', 'ANY')

    rows = []
    # Headers: IMD Code, IMD Name, Policy Number, Type of Business, %, Provision %, From, To, Entry Status
    for r in range(2, ws.max_row + 1):
        imd  = ws.cell(r, 1).value
        tob  = ws.cell(r, 4).value
        pct  = ws.cell(r, 5).value
        stat = ws.cell(r, 9).value
        if not imd or not pct: continue
        rate = clean_rate(pct)
        if rate is None: continue
        rows.append({
            'Version Id*':           ver,
            'Pol Agent Code*':       str(imd).strip(),
            'Product Code*':         'ANY',
            'Final Nm Gwp Ll*':      -1e10,
            'Final Nm Gwp Ul*':      1e11,
            'Type Of Business*':     str(tob).strip() if tob else 'ANY',
            'Insured Age Ll*':       0.0,
            'Insured Age Ul*':       1000.0,
            'Outgo*':                'RGWP',
            'Comm Percentage*':      rate,
            'Total Sum Insured Ll*': def_tsi_ll,
            'Total Sum Insured Ul*': def_tsi_ul,
            'Policy Year*':          policy_year,
            'Unique Id':             f'PW_{r}',
        })
    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════
# API ROUTES
# ══════════════════════════════════════════════════════════════

@app.route('/')
def index(): return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files: return jsonify({'error':'No file'}), 400
        f = request.files['file']
        if not f.filename or not ok_file(f.filename): return jsonify({'error':'Invalid file type'}), 400
        sid = str(uuid.uuid4())[:8]
        fn  = secure_filename(f.filename)
        sp  = os.path.join(UPLOAD_DIR, f"{sid}_{fn}"); f.save(sp)
        sheets = get_sheets(sp)
        return jsonify({'session_id':sid,'filepath':sp,'sheets':sheets,'filename':fn})
    except Exception as e:
        return jsonify({'error':str(e)}), 500


@app.route('/api/preview', methods=['POST'])
def preview():
    """Preview IMD List columns and sample data + product master."""
    try:
        d = request.json
        fp = d['filepath']
        imd_sheet = d.get('imd_sheet', 'IMD List')
        header_row = d.get('header_row', 2)
        data_start = d.get('data_start_row', 3)

        wb = openpyxl.load_workbook(fp, data_only=True)

        # Headers
        if imd_sheet not in wb.sheetnames:
            return jsonify({'error': f"Sheet '{imd_sheet}' not found", 'sheets': wb.sheetnames}), 400
        ws = wb[imd_sheet]
        headers = {}
        for c in range(1, ws.max_column+1):
            v = ws.cell(header_row, c).value
            if v: headers[c] = str(v).strip()

        # Sample rows
        sample = []
        for r in range(data_start, min(data_start+5, ws.max_row+1)):
            row = {}
            for c, h in headers.items():
                v = ws.cell(r, c).value
                row[h] = str(v)[:60] if v is not None else ''
            sample.append(row)

        # Product master
        pm = load_product_master(wb, d.get('product_master_sheet', 'Product Master'))

        # Unique vol slab values
        vol_col = d.get('col_vol_slab', 7)
        vol_vals = set()
        for r in range(data_start, ws.max_row+1):
            v = ws.cell(r, vol_col).value
            if v: vol_vals.add(str(v).strip())

        # Unique product names
        prod_col = d.get('col_product_name', 5)
        prod_vals = set()
        for r in range(data_start, ws.max_row+1):
            v = ws.cell(r, prod_col).value
            if v: prod_vals.add(str(v).strip())

        # Unique statuses
        stat_col = d.get('col_status', 17)
        stat_vals = set()
        for r in range(data_start, ws.max_row+1):
            v = ws.cell(r, stat_col).value
            if v: stat_vals.add(str(v).strip())

        wb.close()
        return jsonify({
            'headers':       {str(k): v for k, v in headers.items()},
            'sample':        sample,
            'product_master': pm,
            'vol_slab_vals': sorted(vol_vals),
            'product_names': sorted(prod_vals),
            'status_vals':   sorted(stat_vals),
            'max_row':       ws.max_row,
            'max_col':       ws.max_column,
        })
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/process', methods=['POST'])
def process():
    try:
        d = request.json
        config = d.get('config', {})
        config['filepath'] = d['filepath']
        session_id = d.get('session_id', 'x')
        output_name = d.get('output_name', 'anh_output')

        df, skipped, total_uid = process_anh(config)

        if df.empty:
            return jsonify({'error': f'No output rows generated. {skipped} rows skipped.'}), 400

        fn  = f"{session_id}_{secure_filename(output_name)}.csv"
        op  = os.path.join(OUTPUT_DIR, fn)

        # Write with the standard AnH CSV header
        header_note = '* means mandatory.\nLeave blank for DB Null.\n\n\n'
        with open(op, 'w', newline='') as fh:
            fh.write(header_note)
        df.to_csv(op, mode='a', index=False)

        return jsonify({
            'success':        True,
            'output_path':    op,
            'output_filename': fn,
            'rows':           len(df),
            'cols':           len(df.columns),
            'skipped':        skipped,
            'preview':        df.head(8).fillna('').astype(str).to_dict('records'),
            'columns':        list(df.columns),
        })
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/download/<filename>')
def download(filename):
    p = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(p): return jsonify({'error': 'Not found'}), 404
    return send_file(p, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    app.run(debug=True, port=5052, host='0.0.0.0')
