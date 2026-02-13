#!/usr/bin/env python3
"""
Export all tables from an Access .mdb/.accdb into a single Excel file
(one worksheet per table) using mdbtools + pandas.

Usage:
  python access_to_excel_linux.py input.mdb output.xlsx
  python access_to_excel_linux.py input.accdb output.xlsx
"""
import io
import subprocess
import sys
from pathlib import Path

import pandas as pd


def run(cmd):
    cp = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if cp.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(cmd)}\n{cp.stderr}")
    return cp.stdout

def list_tables(db_path):
    # -1 : one table per line; filters out system tables with -S (if available)
    # Some Ubuntu builds don’t have -S, so we’ll filter in Python as well.
    out = run(["mdb-tables", "-1", db_path])
    tables = [t.strip() for t in out.splitlines() if t.strip()]
    # Drop common system tables
    return [t for t in tables if not (t.startswith("MSys") or t.startswith("USys"))]

def export_table_csv(db_path, table):
    """
    Returns CSV text for a table using mdb-export.
    -H: include header row
    -D: date/time format
    -d: delimiter
    -q: quote character
    """
    cmd = [
        "mdb-export",
        "-D", "%Y-%m-%d %H:%M:%S",
        "-d", ",",
        "-q", '"',
        db_path,
        table,
    ]
    # print(f"Exporting table: {table}: {' '.join(cmd)}")
    return run(cmd)

def main(db_path, xlsx_path):
    db = Path(db_path)
    if not db.exists():
        sys.exit(f"Database not found: {db}")

    tables = list_tables(str(db))
    if not tables:
        sys.exit("No user tables found (or mdbtools couldn’t read the catalog).")

    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        used_sheet_names = set()
        for t in tables:
            try:
                csv_text = export_table_csv(str(db), t)
                # Heuristic: mdbtools outputs UTF-8 on modern builds
                df = pd.read_csv(io.StringIO(csv_text))
            except Exception as e:
                # Write an error sheet instead of failing the whole export
                err = pd.DataFrame({"table": [t], "error": [str(e)]})
                name = t[:25] + "_ERR" if len(t) > 25 else f"{t}_ERR"
                if name in used_sheet_names:
                    name = f"{name}_1"
                err.to_excel(writer, sheet_name=name, index=False)
                used_sheet_names.add(name)
                continue

            # Sanitize sheet name (Excel max 31 chars, disallowed chars: : \ / ? * [ ])
            sheet = t
            for ch in r':\/?*[]':
                sheet = sheet.replace(ch, "_")
            sheet = (sheet or "Sheet")[:31]
            if sheet in used_sheet_names:
                # ensure uniqueness after truncation
                base = sheet[:29]
                i = 1
                while sheet in used_sheet_names:
                    sheet = f"{base}_{i}"
                    i += 1
            used_sheet_names.add(sheet)

            df.to_excel(writer, sheet_name=sheet, index=False)

            # Nice-to-have: simple autofilter and column width
            ws = writer.sheets[sheet]
            if not df.empty:
                ws.autofilter(0, 0, len(df), len(df.columns)-1)
                for i, col in enumerate(df.columns):
                    try:
                        w = min(max(len(str(col)), *(len(str(x)) for x in df[col].head(200).fillna("").tolist())) + 2, 60)
                        writer.sheets[sheet].set_column(i, i, w)
                    except Exception:
                        pass

    print(f"Exported {len(tables)} table(s) to {xlsx_path}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python access_to_excel_linux.py input.mdb output.xlsx", file=sys.stderr)
        sys.exit(2)
    main(sys.argv[1], sys.argv[2])
