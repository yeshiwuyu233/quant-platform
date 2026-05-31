"""Import Whole Market.xlsx sheets into the SQLite market cache."""
import argparse
import os
import sys
from datetime import datetime

import pandas as pd

from db_service import get_db, init_db

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_XLSX = os.path.join(PROJECT_ROOT, "Whole Market.xlsx")


def _clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def _clean_float(value):
    if pd.isna(value):
        return None
    try:
        return float(value)
    except Exception:
        return None


def import_sheet(conn, xlsx_path: str, sheet_name: str) -> int:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    required = {"代码", "全称", "行业", "准确率", "今日指标"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"sheet {sheet_name} missing columns: {sorted(missing)}")

    rows = []
    for row_order, (_, row) in enumerate(df.iterrows()):
        code = _clean_text(row.get("代码"))
        if not code:
            continue
        rows.append((
            sheet_name,
            row_order,
            code,
            _clean_text(row.get("全称")),
            _clean_text(row.get("行业")),
            _clean_float(row.get("准确率")),
            _clean_float(row.get("今日指标")),
            _clean_text(row.get("指标历史")),
        ))

    conn.execute("DELETE FROM market_snapshot WHERE trade_date = ?", (sheet_name,))
    conn.executemany(
        """
        INSERT INTO market_snapshot
            (trade_date, row_order, code, name, industry, accuracy, indicator, indicator_history)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(trade_date, code) DO UPDATE SET
            row_order = excluded.row_order,
            name = excluded.name,
            industry = excluded.industry,
            accuracy = excluded.accuracy,
            indicator = excluded.indicator,
            indicator_history = excluded.indicator_history
        """,
        rows,
    )
    conn.execute(
        """
        INSERT INTO import_meta (trade_date, rows_count, imported_at)
        VALUES (?, ?, ?)
        ON CONFLICT(trade_date) DO UPDATE SET
            rows_count = excluded.rows_count,
            imported_at = excluded.imported_at
        """,
        (sheet_name, len(rows), datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    )
    return len(rows)


def import_workbook(xlsx_path: str, limit_sheets: int | None = None, validate_only: bool = False) -> list[tuple[str, int]]:
    xls = pd.ExcelFile(xlsx_path)
    sheets = [s for s in xls.sheet_names if s.isdigit() and len(s) == 4]
    sheets = sorted(sheets)
    if limit_sheets:
        sheets = sheets[-limit_sheets:]

    if validate_only:
        return [(s, len(pd.read_excel(xlsx_path, sheet_name=s, usecols=["代码"]))) for s in sheets]

    conn = get_db()
    try:
        init_db(conn)
        results = []
        for sheet in sheets:
            count = import_sheet(conn, xlsx_path, sheet)
            results.append((sheet, count))
        conn.commit()
        return results
    finally:
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Import Whole Market.xlsx into market.db")
    parser.add_argument("--xlsx", default=DEFAULT_XLSX)
    parser.add_argument("--limit-sheets", type=int, default=None)
    parser.add_argument("--validate-only", action="store_true")
    args = parser.parse_args()

    if not os.path.exists(args.xlsx):
        print(f"missing workbook: {args.xlsx}", file=sys.stderr)
        return 1

    results = import_workbook(args.xlsx, args.limit_sheets, args.validate_only)
    for sheet, count in results:
        print(f"{sheet}: {count}")
    print(f"sheets: {len(results)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())