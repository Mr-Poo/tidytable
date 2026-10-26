# TidyTable

> **Windows app — no Python needed.** A ready-to-run `.exe` build of this tool is available on [Gumroad](https://duonghai8.gumroad.com/l/tidytable) ($19.00). Buying it supports development; this source code stays free under the MIT license.

Offline command-line cleaner for messy CSV and Excel sheets. It trims and
normalizes text, repairs mistyped numbers, removes duplicates, enforces simple
validation rules and writes a cleaned file plus an optional report of flagged
rows. The input file is never modified. No network access, no telemetry.

## Install

    pip install .          # Python 3.12+, needs pandas and openpyxl

## Usage

    tidytable INPUT -o OUTPUT [options]

    tidytable inventory.csv -o inventory_clean.csv --required product,qty --no-negative qty --report problems.csv
    tidytable contacts.xlsx -o contacts_clean.xlsx --case title:name --key email

Options:

- `-o, --output` output `.csv` or `.xlsx` (required, must differ from input)
- `--sheet NAME` sheet of an `.xlsx` input (default: first sheet)
- `--case STYLE:COLS` `upper`, `lower` or `title` for columns (repeatable)
- `--key COLS` dedupe on these columns (default: all columns)
- `--required COLS` rows with a blank value in these columns are removed and flagged
- `--no-negative COLS` rows with a negative number in these columns are removed and flagged
- `--report FILE.csv` write flagged rows with `row`, `action` and `reason`

## What it does

- Trims and collapses whitespace, removes non-printing characters.
- Numeric columns (detected automatically) are repaired: currency symbols,
  `1.234,56` vs `1,234.56`, stray letters (`12kg`, `USD 7`) become numbers.
- Values that cannot be parsed safely (`abc`, ambiguous `1.234`) are left
  untouched and listed in the report as `kept for review`.
- Duplicates are removed after normalization, comparing case-insensitively
  and keeping the first occurrence.
- Rows violating `--required` / `--no-negative` are removed and reported
  as `removed`. `row` is the row number in the source sheet (header = 1).
- A summary (rows in/out, duplicates removed, cells fixed, rows flagged) is
  printed. Exit code is 0 on success and non-zero on bad input.
- Without `--report`, no extra file is written; flagged rows are listed on stdout.

## Limitations

- Only `.csv` and `.xlsx`; no PDFs, images, Word files or emails.
- Formulas are not preserved; cleaned values are written as static data.
- One sheet per run.
- Number-format detection is heuristic; ambiguous values are flagged.
- Column names containing commas cannot be used in option lists.
- No charts, pivot tables or online lookups.

## Tests

    pip install -r requirements.txt pytest && pytest

## Windows app

Don't want to install Python? Get the ready-to-run Windows build (single `.exe`, runs fully offline) on [Gumroad](https://duonghai8.gumroad.com/l/tidytable).

## License

MIT — see [LICENSE](LICENSE).
