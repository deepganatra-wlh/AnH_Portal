# AnH Grid Processor

A Flask web application that converts the **Agency & Health (AnH) Special Grid Master** Excel file into the `magma_anh_monthly` CSV format required for commission processing.

---

## Table of Contents

- [Overview](#overview)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [Running the App](#running-the-app)
- [How to Use](#how-to-use)
- [Input File Format](#input-file-format)
- [Output File Format](#output-file-format)
- [Transformation Rules](#transformation-rules)
- [API Reference](#api-reference)
- [Configuration Reference](#configuration-reference)
- [Vol/Slab Mapping](#volslab-mapping)
- [Product Code Resolution](#product-code-resolution)

---

## Overview

The AnH Grid Processor reads an Agency & Health grid Excel file, applies a set of business transformation rules, and outputs a clean CSV that can be directly loaded into the commission database.

Each source row in the Excel file can generate up to **3 output rows** depending on what rate columns are populated:

- One row for **Fresh Business** (Type of Business = `New`)
- One row for **Renewal Business** (Type of Business = `Renewal`)
- One extra row for **Renewal age > 60** when a separate rate is specified

---

## Project Structure

```
.
├── app.py                  # Main Flask application
├── templates/
│   └── index.html          # Frontend UI (5-step workflow)
├── uploads/                # Uploaded Excel files (auto-created)
├── outputs/                # Generated CSV files (auto-created)
└── README.md
```

---

## Requirements

- Python 3.8+
- Flask
- pandas
- openpyxl
- werkzeug

Install all dependencies:

```bash
pip install flask pandas openpyxl werkzeug
```

---

## Installation

```bash
git clone <your-repo-url>
cd anh-grid-processor
pip install flask pandas openpyxl werkzeug
```

---

## Running the App

```bash
python app.py
```

The server starts at `http://0.0.0.0:5052`. Open `http://localhost:5052` in your browser.

---

## How to Use

The UI walks you through a 5-step workflow:

**Step 1 — Upload File**
Drag and drop (or click to browse) your Agency & Health grid Excel file. Accepted formats: `.xlsx`, `.xls`, `.xlsb`, `.csv`. Maximum file size: 100 MB.

**Step 2 — Configure Sheets**
Select which sheets to use and set the header and data start row positions. Defaults match the standard grid format (header row 2, data from row 3).

**Step 3 — Map Columns**
Enter the 1-indexed column number for each field (IMD Code, Product Name, Fresh %, Renewal %, etc.). Set a column to `0` if it is absent in your file.

**Step 4 — Preview Data**
Loads a 5-row sample from the IMD List sheet along with detected product codes, vol/slab values, and statuses. Use this to verify the column mapping before processing.

**Step 5 — Process & Download**
Runs the full transformation and presents an 8-row output preview with row/skip counts. Download the resulting CSV directly from this page.

---

## Input File Format

The Excel file must contain:

### IMD List sheet *(required)*

| Col | Default Position | Field |
|-----|-----------------|-------|
| Relationship Code | 2 | Agent relationship code |
| IMD Code | 3 | Agent IMD code |
| IMD Name | 4 | Agent name |
| Product Code / Name | 5 | Product name (looked up in Product Master) |
| 80.10.10 Flag | 6 | Special scheme flag |
| Vol / NBP / Slab | 7 | Business volume slab |
| Fresh Business % | 8 | Commission rate for new business |
| Rollover Payout % | 9 | Rollover commission rate |
| Renewal Business % | 10 | Commission rate for renewals |
| Renewal % (Age > 60) | 11 | Commission rate for renewals where insured age > 60 |
| Status | 17 | Row active/inactive status |

### Product Master sheet *(optional but recommended)*

Maps product names to numeric codes. Two-column format:

| Column 1 | Column 2 |
|----------|----------|
| Product Code / Name | Product Code |
| All Non-Motor Product | `Arogya Sanjeevani (6116), One Health (6111), ...` |
| All Non-Motor Product (Except 8101, 8103) | `Excluding 8101, 8103` |

See [Product Code Resolution](#product-code-resolution) for how these values are processed.

### Policy Wise sheet *(optional)*

Policy-specific rates. Columns expected: IMD Code (1), Type of Business (4), Rate % (5), Entry Status (9).

---

## Output File Format

The output is a CSV with a 4-line preamble followed by the data:

```
* means mandatory.
Leave blank for DB Null.


Version Id*,Pol Agent Code*,Product Code*,...
```

### Output Columns

| Column | Description |
|--------|-------------|
| `Version Id*` | Grid version identifier |
| `Pol Agent Code*` | IMD or Relationship code; `ANY` for common grid rows |
| `Product Code*` | Resolved numeric product codes (comma-separated) |
| `Final Nm Gwp Ll*` | GWP lower bound from Vol/Slab mapping |
| `Final Nm Gwp Ul*` | GWP upper bound from Vol/Slab mapping |
| `Type Of Business*` | `New` or `Renewal` |
| `Insured Age Ll*` | Insured age lower bound |
| `Insured Age Ul*` | Insured age upper bound |
| `Outgo*` | `GWP` for common grid rows, `RGWP` for specific rates |
| `Comm Percentage*` | Commission rate as a float (e.g. `33.9`) |
| `Total Sum Insured Ll*` | TSI lower bound |
| `Total Sum Insured Ul*` | TSI upper bound |
| `Policy Year*` | Policy year filter value |
| `Unique Id` | Sequential row identifier |

---

## Transformation Rules

1. **Fresh Business row** — generated when the Fresh % column has a valid rate. `Type Of Business` = `New`.
2. **Renewal Business row** — generated when the Renewal % column has a valid rate. `Type Of Business` = `Renewal`.
3. **Renewal Age > 60 row** — generated when the Renewal60 column has a rate that differs from the standard renewal rate. `Insured Age Ll` = `61`.
4. **Common Grid rows** — rows where the rate value is `Common grid` or `-` use `Pol Agent Code` = `ANY` and `Outgo` = `GWP`.
5. **Specific rate rows** — rows with an actual numeric rate use the IMD/Relationship code and `Outgo` = `RGWP`.
6. **Agent code** — the IMD Code column is always used as `Pol Agent Code*`. The Relationship Code column is not used for this mapping.
7. **Rate cleaning** — rates stored as decimals (e.g. `0.3814`) are multiplied by 100 to give a percentage (`38.14`). Percent strings like `33.90%` are stripped of the `%` sign.
8. **Status filtering** — rows whose status column does not match the configured active statuses (default: `active`) are skipped.
9. **Blank rows skipped** — rows with no IMD Code and no Relationship Code are skipped.

---

## API Reference

All endpoints are JSON-based except `/api/upload` (multipart form) and `/api/download` (file download).

### `POST /api/upload`

Upload an Excel file and get back the session ID, file path, and detected sheet names.

**Request:** `multipart/form-data` with field `file`

**Response:**
```json
{
  "session_id": "a1b2c3d4",
  "filepath": "/path/to/uploads/a1b2c3d4_file.xlsx",
  "sheets": ["IMD List", "Product Master", "Policy Wise"],
  "filename": "file.xlsx"
}
```

---

### `POST /api/preview`

Preview headers, sample rows, detected product codes, vol/slab values, and statuses from the uploaded file.

**Request body:**
```json
{
  "filepath": "/path/to/uploads/a1b2c3d4_file.xlsx",
  "imd_sheet": "IMD List",
  "header_row": 2,
  "data_start_row": 3
}
```

**Response:**
```json
{
  "headers": {"1": "Sr No", "2": "Relationship Code", "...": "..."},
  "sample": [{"Relationship Code": "10000055685", "...": "..."}],
  "product_master": {"All Non-Motor Product": "6116, 6111, 6117, ..."},
  "vol_slab_vals": ["-", "<=50 K", ">50 K"],
  "status_vals": ["active", "inactive"],
  "max_row": 250,
  "max_col": 18
}
```

---

### `POST /api/process`

Run the full transformation and save the output CSV.

**Request body:**
```json
{
  "filepath": "/path/to/uploads/a1b2c3d4_file.xlsx",
  "session_id": "a1b2c3d4",
  "output_name": "anh_output",
  "config": {
    "imd_sheet": "IMD List",
    "header_row": 2,
    "data_start_row": 3,
    "col_imd_code": 3,
    "col_fresh_pct": 8,
    "active_statuses": ["active"]
  }
}
```

**Response:**
```json
{
  "success": true,
  "output_filename": "a1b2c3d4_anh_output.csv",
  "rows": 412,
  "cols": 14,
  "skipped": 38,
  "columns": ["Version Id*", "Pol Agent Code*", "..."],
  "preview": [{"Version Id*": "magma_anh_monthly_Grid_4", "...": "..."}]
}
```

---

### `GET /api/download/<filename>`

Download a previously generated output CSV by filename.

---

## Configuration Reference

All `config` keys accepted by `/api/process`:

| Key | Default | Description |
|-----|---------|-------------|
| `imd_sheet` | `IMD List` | Sheet name for the main grid data |
| `product_master_sheet` | `Product Master` | Sheet name for product code mapping |
| `header_row` | `2` | Row number containing column headers |
| `data_start_row` | `3` | First row of actual data |
| `pm_col_name` | `1` | Product Master: name column index |
| `pm_col_code` | `2` | Product Master: code column index |
| `col_rel_code` | `2` | Relationship Code column |
| `col_imd_code` | `3` | IMD Code column |
| `col_imd_name` | `4` | IMD Name column |
| `col_product_name` | `5` | Product Name column |
| `col_80_10_10` | `6` | 80-10-10 flag column |
| `col_vol_slab` | `7` | Vol/NBP/Slab column |
| `col_fresh_pct` | `8` | Fresh Business % column |
| `col_rollover_pct` | `9` | Rollover Payout % column |
| `col_renewal_pct` | `10` | Renewal Business % column |
| `col_renewal60_pct` | `11` | Renewal % for age > 60 column |
| `col_status` | `17` | Status column |
| `active_statuses` | `["active"]` | Status values to include; case-insensitive |
| `skip_common_grid_fresh` | `false` | Skip rows where Fresh = `Common grid` |
| `skip_common_grid_renewal` | `false` | Skip rows where Renewal = `Common grid` |
| `common_grid_agent` | `ANY` | `Pol Agent Code` for common grid rows |
| `common_grid_outgo` | `GWP` | `Outgo` value for common grid rows |
| `specific_rate_outgo` | `RGWP` | `Outgo` value for specific rate rows |
| `version_id` | `magma_anh_monthly_Grid_4` | Value written to `Version Id*` column |
| `policy_year` | `ANY` | Value written to `Policy Year*` column |
| `default_insured_age_ll` | `0.0` | Default insured age lower bound |
| `default_insured_age_ul` | `1000.0` | Default insured age upper bound |
| `renewal60_age_ll` | `61.0` | Insured age lower bound for age > 60 rows |
| `renewal60_age_ul` | `1000.0` | Insured age upper bound for age > 60 rows |
| `default_tsi_ll` | `-1e10` | Total Sum Insured lower bound |
| `default_tsi_ul` | `1e11` | Total Sum Insured upper bound |

---

## Vol/Slab Mapping

The `Vol/NBP/Slab` column value is mapped to `Final Nm Gwp Ll*` and `Final Nm Gwp Ul*`:

| Slab String | GWP Lower | GWP Upper |
|-------------|-----------|-----------|
| `-` or blank | -1e10 | 1e11 |
| `<=50 K` | -1e10 | 49,999 |
| `>50 K` | 50,000 | 1e11 |
| `>50K - 1.5 Lac` | 50,000 | 1,49,999 |
| `>1.5 Lac - 2 lac` | 1,50,000 | 1,99,999 |
| `>2 Lac` | 2,00,000 | 1e11 |
| `Upto 199999` | -1e10 | 1,99,999 |
| `200000 & Above` | 2,00,000 | 1e11 |

Any unrecognised slab value falls back to `-1e10` / `1e11`.

---

## Product Code Resolution

Product codes are resolved from the **Product Master** sheet using a two-pass approach:

**Pass 1** — The row with the most numeric codes is identified as the full universe (typically the `All Non-Motor Product` row).

**Pass 2** — Each entry is resolved:

| Product Master Entry | Raw Code Column | Resolved Output |
|---------------------|-----------------|-----------------|
| All Non-Motor Product | `Arogya Sanjeevani (6116), One Health (6111), ...` | `6116, 6111, 6117, 6118, 6119, 8101, 8103, 9993, 9997` |
| All Non-Motor Product (Except 8101, 8103) | `Excluding 8101, 8103` | `6116, 6111, 6117, 6118, 6119, 9993, 9997` |
| All Non-Motor Product (Except 6119) | `Excluding 6119` | `6116, 6111, 6117, 6118, 8101, 8103, 9993, 9997` |
| All Non-Motor Product (Except 6119, 8101, 8103) | `Excluding 6119, 8101, 8103` | `6116, 6111, 6117, 6118, 9993, 9997` |
| IPA (8101) + OneProtect (8103) | `8101, 8103` | `8101, 8103` |
| Double Suraksha (6119) | `6119` | `6119` |

The key rule: **`Excluding X, Y` means take the full code universe and remove X and Y**, not just output the string literally.
