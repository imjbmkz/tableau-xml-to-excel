# tableau-xml-to-excel

A utility script that extracts calculated field metadata from a Tableau workbook's XML and exports it to a formatted Excel file.

It's useful for documenting dashboards, auditing logic, handing off workbooks, and finding calculated fields that are no longer used on any chart.

## What you get

The script produces an `.xlsx` file with a single **Calculated Fields** sheet containing one row per calculated field:

| Column | Description |
| --- | --- |
| Calculated Field Name | The field's display name (caption) in Tableau |
| Internal Name | Tableau's internal identifier, e.g. `Calculation_1234567890` |
| Datasource | The datasource the field belongs to |
| Data Type | `string`, `real`, `integer`, `date`, etc. |
| Calculation Logic (Tableau) | The full formula, as written in Tableau |
| Dependencies | Other fields the formula references, shown by caption where possible |
| Used on Chart | `Y` (green) if the field appears on any worksheet, `N` (red) if not |
| Worksheet(s) | Worksheets that use the field |
| Dashboard Page(s) | Dashboards containing those worksheets |

Rows are sorted with used fields first, then by datasource and field name. The header row is frozen and filters are enabled so you can slice the list in Excel.

A summary is also printed to the terminal:

```
Excel file saved successfully!
Path: tableau-def.xlsx

Summary:
  Total calculated fields: 42
  Fields used on charts: 35
  Fields NOT used on charts: 7

  Dashboard pages: 3
    - Overview (4 charts)
    - Regional Breakdown (6 charts)
    - Trends (3 charts)

  Total worksheets (charts): 13
```

## Requirements

- Python 3.11 or newer
- The packages in `requirements.txt` (`pandas`, `openpyxl`, `numpy`)

## Installation

```bash
git clone https://github.com/imjbmkz/tableau-xml-to-excel.git
cd tableau-xml-to-excel

python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

## Getting the XML from Tableau

A Tableau workbook (`.twb`) is already an XML file, so you can point the script at it directly.

If you have a packaged workbook (`.twbx`), it's a zip archive. Extract it (rename to `.zip` and unzip, or use any archive tool) and use the `.twb` file inside.

The script doesn't care about the file extension; `.twb`, `.xml`, and `.txt` all work as long as the contents are the workbook XML.

## Usage

```bash
python3 main.py --source-xml tableau-def.xml
```

### Options

| Flag | Required | Description |
| --- | --- | --- |
| `--source-xml` | Yes | Path to the Tableau workbook XML |
| `--output` | No | Path for the Excel file. See default below. |

### Default output name

If `--output` isn't given, the Excel file is named after the source file, lowercased with spaces replaced by underscores, and saved in the same folder as the source:

| Source | Output |
| --- | --- |
| `tableau-def.xml` | `tableau-def.xlsx` |
| `Sales Dashboard.twb` | `sales_dashboard.xlsx` |
| `reports/My Workbook Calc Fields.txt` | `reports/my_workbook_calc_fields.xlsx` |

### Examples

```bash
# Default output name
python3 main.py --source-xml "Sales Dashboard.twb"

# Custom output path
python3 main.py --source-xml "Sales Dashboard.twb" --output docs/sales_calc_fields.xlsx
```

## How it works

1. **Calculated fields.** For every datasource, it collects fields that have a formula, along with their caption and data type (from `<column>` and `<metadata-record>` elements). HTML entities in formulas are decoded and line endings are normalized.
2. **Dependencies.** Field references in square brackets (e.g. `[Sales]`) are parsed out of each formula and mapped back to their display names. Parameter prefixes are ignored.
3. **Worksheet usage.** For every worksheet, it collects the fields listed under `datasource-dependencies` and `column-instance`.
4. **Dashboard mapping.** Each dashboard's zones are matched against worksheet names to link worksheets to dashboards.
5. **Export.** Everything is assembled into a table and written to Excel with styling.

## Limitations

- "Used on Chart" only reflects direct use on a worksheet. A field marked `N` might still be used indirectly, for example inside another calculated field that is on a chart. Check the **Dependencies** column of other fields before deleting anything.
- Dependency detection is regex-based, so it lists every bracketed reference in the formula, including references to regular (non-calculated) fields.
- Parameters themselves are not exported as separate rows.
