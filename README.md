# Excel Reports Merger

A Python command-line script that merges multiple Excel reports from a folder into one formatted workbook. It tags every row with the file it came from, skips problem files instead of crashing, and logs everything.

Useful when you receive the same report from different people or months (sales, inventory, timesheets) and need a single clean file.

## Features

- Reads every `.xlsx` file in the `reports/` folder (ignores Excel temp files)
- Adds a `source_file` column so you can trace each row to its original file
- Detects the common column layout and skips files that don't match, with a clear reason
- Skips unreadable or empty files instead of crashing
- Creates `merged_report.xlsx` with two sheets:
  - `Merged`: all rows from the valid files
  - `Summary`: number of rows read from each file
- Formatted output: styled header, auto column widths, frozen first row, and filters
- Friendly message if the output file is open in Excel
- Writes a timestamped log to `merge.log`

## How It Works

1. **Read**: each report is loaded and tagged with its file name.
2. **Validate**: the column layout shared by most files is used as the expected layout. Files that differ are skipped and reported.
3. **Merge**: valid files are combined into one table.
4. **Export**: the result is saved as a formatted Excel workbook.

## Requirements

- Python 3.9 or newer
- `pandas` and `openpyxl`

```bash
pip install pandas openpyxl
```

## Usage

1. Put your `.xlsx` reports in the `reports/` folder (it is created automatically on first run).
2. Run:

```bash
python merge_reports.py
```

Optional arguments:

```bash
python merge_reports.py --input my_folder --output result.xlsx
```

| Argument   | Description                                   | Default             |
|------------|-----------------------------------------------|---------------------|
| `--input`  | Folder that contains the `.xlsx` reports      | `reports/`          |
| `--output` | Name of the merged workbook                   | `merged_report.xlsx`|

## Example Output

```
$ python merge_reports.py
merged 2 files, 3 rows -> merged_report.xlsx
  skipped broken.xlsx: unreadable file
  skipped bad_columns.xlsx: columns differ - missing: price, product, quantity; unexpected: item, qty
```

`Merged` sheet:

| source_file    | date       | product  | quantity | price |
|----------------|------------|----------|----------|-------|
| sales_feb.xlsx | 2026-10-03 | Monitor  | 1        | 150   |
| sales_jan.xlsx | 2026-10-01 | Keyboard | 3        | 25.5  |
| sales_jan.xlsx | 2026-10-02 | Mouse    | 5        | 12    |

`Summary` sheet:

| source_file    | rows |
|----------------|------|
| sales_feb.xlsx | 1    |
| sales_jan.xlsx | 2    |

## Notes

- Only the first sheet of each workbook is read.
- Only `.xlsx` files are supported.
- All reports should share the same columns; files with a different layout are skipped and listed in the terminal and in `merge.log`.
- Paths are relative to the script's folder, so it works from any directory.

## Project Structure

```
.
├── merge_reports.py
├── README.md
├── reports/             # put your .xlsx reports here
├── merged_report.xlsx   # created automatically
└── merge.log            # created automatically
```

## Libraries Used

`pandas`, `openpyxl`, plus `argparse`, `logging`, `collections`, `pathlib`, `sys` from the Python standard library

## Author

GitHub: https://github.com/firasaltawil1714
