# Superbill Processor

A desktop GUI tool for processing Lynx Superbill export files (`.xlsx`) and appending cleaned, remapped rows to a consolidated output workbook.

## Features

- **Drag-and-drop support** — drop an Excel file directly onto the Input or Output file field to populate its path
- **Browse buttons** — alternatively use **Browse…** / **Open…** to select files via a file dialog
- **Empty-column validation** — detects and removes known empty columns; warns on any discrepancy
- **Duplicate detection** — identifies rows already present in the output (matched by Date of Service / Patient Name / Billing Code); reports each duplicate with its Excel row number; rows with no identifying information are skipped automatically
- **Column remapping** — maps input columns to the output schema automatically
- **Append mode** — safely appends new rows after the last non-empty row in the output file
- **Backup & verify** — creates a backup before writing and verifies all appended rows after saving
- **About dialog** — click **ℹ About** to view this README inside the application

## Column mapping

| Input col | Output col |
|-----------|------------|
| A         | A          |
| B         | C          |
| C         | D          |
| D         | E          |
| E         | F          |
| F         | G          |
| G         | H          |
| O         | I          |
| P         | J          |
| Q         | K          |
| R         | L          |
| S         | M          |
| T         | N          |
| U         | O          |
| V         | P          |
| AD        | X          |
| AE        | Y          |
| AF        | Z          |

## Requirements

- Python 3.9+
- `pandas`
- `openpyxl`
- `tkinterdnd2`

Install dependencies:

```bash
pip install pandas openpyxl tkinterdnd2
```

## Usage

```bash
python superbill_processor.py
```

1. Provide the **Input Superbill file** — type the path, click **Browse…**, or drag and drop the file onto the field.
2. Provide the **Output file** — type the path, click **Open…**, or drag and drop the file onto the field.
3. Click **▶ Process**.
4. If duplicates are found, review the list (each entry shows the Excel row number, Date of Service, Patient Name, and Billing Code) then choose **Proceed** or **Abort**.
5. All messages appear in the **Messages** pane. Click **Clear log** to reset it.
6. Click **ℹ About** at any time to view this documentation.

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
