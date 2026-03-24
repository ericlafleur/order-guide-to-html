"""Convert xlsx files from the workbooks/ folder to HTML files in workbook_html/."""

import argparse
import pathlib
import pandas as pd

WORKBOOKS_DIR = pathlib.Path("workbooks")
OUTPUT_DIR = pathlib.Path("workbook_html")

parser = argparse.ArgumentParser(description="Convert xlsx workbooks to HTML.")
parser.add_argument(
    "files",
    nargs="*",
    type=pathlib.Path,
    help="Specific xlsx files to convert. Defaults to all files in workbooks/.",
)
args = parser.parse_args()

OUTPUT_DIR.mkdir(exist_ok=True)

xlsx_files = args.files if args.files else list(WORKBOOKS_DIR.glob("*.xlsx"))

if not xlsx_files:
    print("No xlsx files found in workbooks/")
else:
    for xlsx_path in xlsx_files:
        if not xlsx_path.exists():
            print(f"Skipping {xlsx_path}: file not found")
            continue
        stem = xlsx_path.stem
        all_sheets_html = []

        with pd.ExcelFile(xlsx_path) as xls:
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                table_html = df.to_html(index=False, border=1, classes="order-guide-table")
                all_sheets_html.append(
                    f"<h2>{sheet_name}</h2>\n{table_html}"
                )

        body_content = "\n".join(all_sheets_html)
        html = f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{stem}</title>
    <style>
      body {{ font-family: sans-serif; padding: 1rem; }}
      table.order-guide-table {{ border-collapse: collapse; width: 100%; margin-bottom: 2rem; }}
      table.order-guide-table th, table.order-guide-table td {{ border: 1px solid #ccc; padding: 0.4rem 0.8rem; text-align: left; }}
      table.order-guide-table th {{ background: #f0f0f0; }}
    </style>
  </head>
  <body>
    <h1>{stem}</h1>
    {body_content}
  </body>
</html>
"""
        output_path = OUTPUT_DIR / f"{stem}.html"
        output_path.write_text(html, encoding="utf-8")
        print(f"Converted {xlsx_path} -> {output_path}")
