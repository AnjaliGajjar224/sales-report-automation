# Sales Report Automation

Starter repository for the **Build a Real-World Automation** Codex lesson.

## Input data

- `sales_data.xlsx` contains six fictional transactions in the `Sales Data` worksheet.
- Columns: `Order ID`, `Date`, `Region`, `Product`, `Units`, `Unit Price`.
- Region values intentionally have inconsistent capitalization and whitespace.
- There is **no Revenue column** in the source; Codex will build the automation to calculate it.
- Do not modify the original input file.

## Expected totals (for independent verification)

- North: 2,500
- South: 1,700
- West: 2,300
- Overall: **6,500**

## Demonstration

Use `chatgpt.com/codex` to connect this repository in a prepared/published Codex Cloud environment. The instructor will supply the prompts during the recording. This repository now includes the automation and generated sample report.

## Later lessons

After completing the recording, merge the verified automation changes to the GitHub branch before proceeding to the debugging and refactoring lessons.

## Usage

Requires Python 3.10 or newer. From this directory:

```bash
python -m pip install -r requirements.txt
python automate_report.py sales_data.xlsx sales_report.xlsx
```

Both paths are required positional arguments and may be absolute or relative to
your current directory. Quote paths containing spaces. The output directory must
already exist; an existing output report is replaced. The input is never saved
or overwritten, and using the same file for both paths is rejected.

The script reads the first worksheet, with headers in row 1, and locates the six
required columns by name (column order may vary). Fully blank rows are skipped.
Regions have surrounding and repeated whitespace removed and use title case.
Units and Unit Price must be finite numbers or numeric text. Missing/duplicate
required headers, blank regions, and invalid numbers cause an error with a
nonzero exit status; row validation errors include the source row number.
Formula inputs require cached values previously saved by Excel.

`Cleaned Data` contains the six source columns plus calculated Revenue.
`Regional Summary` contains alphabetically sorted regional totals and Overall
Total. Calculations use decimal arithmetic without rounding; Excel displays
prices and revenue to two decimal places. Dates retain their source values.

The sample run processed all six transactions without errors. The generated
report was reopened and its transaction revenues and regional totals were
independently checked against the source: North 2,500; South 1,700; West 2,300;
overall 6,500. The source workbook's SHA-256 hash was unchanged.
