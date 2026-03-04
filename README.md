# Superbill Processor

A desktop GUI tool for processing Lynx Superbill export files (`.xlsx`) and appending cleaned, remapped rows to a consolidated output workbook.

## Features

- **File search UI** — search for input files by name within a selected folder
- **Empty-column validation** — detects and removes known empty columns; warns on any discrepancy
- **Duplicate detection** — identifies rows already present in the output (matched by Date of Service / Patient Name / Billing Code) and stops before writing duplicates
- **Column remapping** — maps input columns to the output schema automatically
- **Append mode** — safely appends new rows to an existing output file or creates a new one

## Column mapping

| Input col | Output col |
|-----------|-----------|
| A | A |
| B | C |
| C | D |
| D | E |
| E | F |
| F | G |
| G | H |
| O | I |
| P | J |
| Q | K |
| R | L |
| S | M |
| T | N |
| U | O |
| V | P |
| AD | X |
| AE | Y |
| AF | Z |

## Requirements

- Python 3.9+
- `pandas`
- `openpyxl`

Install dependencies:

```bash
pip install pandas openpyxl
```

## Usage

```bash
python superbill_processor.py
```

1. Type a filename (or part of it) in the **Search by name** box, or click **Browse…** to locate your Superbill export.
2. Optionally click **Browse folder…** to change the search root directory.
3. Select the matching file from the results list.
4. Choose or create an **output file** with **Browse / Create…**.
5. Click **▶ Process**.
6. Review messages in the **Messages** pane. Any discrepancies, duplicates, or errors are reported there.

## Input file format

The tool expects a Lynx Superbill report `.xlsx` where:
- Rows 0–5 are report metadata/title
- Row 6 is the column header row
- Data starts at row 7

## Expected empty columns (removed automatically)

- Primary Carrier
- Primary Policy
- Secondary Carrier
- Secondary Policy
- Tertiary Carrier
- Tertiary Policy
- Clinical Trial
- Seq No
- Comment
