"""
SQLite data access helpers for the market cache.

Phase 1 keeps Excel as the source of truth and uses SQLite as a read-through
cache/query layer. All helpers fail closed so callers can fall back to Excel.
"""
import os
import sqlite3
from typing import Iterable, Optional

import pandas as pd

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.environ.get("MARKET_DB_PATH", os.path.join(PROJECT_ROOT, "market.db"))


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db(conn: Optional[sqlite3.Connection] = None) -> None:
    close = False
    if conn is None:
        conn = get_db()
        close = True
    try:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS market_snapshot (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                trade_date TEXT NOT NULL,
                code TEXT NOT NULL,
                name TEXT NOT NULL DEFAULT '',
                industry TEXT NOT NULL DEFAULT '',
                accuracy REAL,
                indicator REAL,
                indicator_history TEXT,
                raw_json TEXT,
                row_order INTEGER NOT NULL DEFAULT 0,
                UNIQUE(trade_date, code)
            );
            CREATE INDEX IF NOT EXISTS idx_ms_date ON market_snapshot(trade_date);
            CREATE INDEX IF NOT EXISTS idx_ms_ind ON market_snapshot(industry);
            CREATE INDEX IF NOT EXISTS idx_ms_industry ON market_snapshot(industry);
            CREATE INDEX IF NOT EXISTS idx_ms_date_acc ON market_snapshot(trade_date, accuracy);

            CREATE TABLE IF NOT EXISTS daily_tracking (
                trade_date TEXT PRIMARY KEY,
                acc_08_raw TEXT,
                all_raw TEXT,
                next_10 INTEGER,
                raw_json TEXT,
                updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
            CREATE INDEX IF NOT EXISTS idx_dt_date ON daily_tracking(trade_date);

            CREATE TABLE IF NOT EXISTS import_meta (
                trade_date TEXT PRIMARY KEY,
                rows_count INTEGER NOT NULL DEFAULT 0,
                imported_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
        """)
        cols = {row[1] for row in conn.execute("PRAGMA table_info(market_snapshot)").fetchall()}
        if "row_order" not in cols:
            conn.execute("ALTER TABLE market_snapshot ADD COLUMN row_order INTEGER NOT NULL DEFAULT 0")
        conn.commit()
    finally:
        if close:
            conn.close()


def fetchone(sql: str, params: tuple = ()) -> Optional[dict]:
    try:
        conn = get_db()
        row = conn.execute(sql, params).fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception:
        return None


def fetchall(sql: str, params: tuple = ()) -> list[dict]:
    try:
        conn = get_db()
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return [dict(r) for r in rows]
    except Exception:
        return []


def execute(sql: str, params: tuple = ()) -> int:
    try:
        conn = get_db()
        conn.execute(sql, params)
        conn.commit()
        affected = conn.total_changes
        conn.close()
        return affected
    except Exception:
        return 0


def executemany(sql: str, seq: Iterable[tuple]) -> int:
    try:
        conn = get_db()
        conn.executemany(sql, list(seq))
        conn.commit()
        affected = conn.total_changes
        conn.close()
        return affected
    except Exception:
        return 0


def get_market_dates() -> list[str]:
    rows = fetchall("SELECT trade_date FROM import_meta ORDER BY trade_date")
    return [str(r["trade_date"]) for r in rows]


def read_market_sheet(trade_date: str) -> Optional[pd.DataFrame]:
    if not trade_date:
        return None
    rows = fetchall(
        """
        SELECT code AS 代码,
               name AS 全称,
               industry AS 行业,
               accuracy AS 准确率,
               indicator AS 今日指标,
               indicator_history AS 指标历史
        FROM market_snapshot
        WHERE trade_date = ?
        ORDER BY row_order, code
        """,
        (trade_date,),
    )
    if not rows:
        return None
    return pd.DataFrame(rows)


def count_market_rows(trade_date: str) -> Optional[int]:
    row = fetchone("SELECT rows_count FROM import_meta WHERE trade_date = ?", (trade_date,))
    return int(row["rows_count"]) if row else None


def get_industries(trade_date: str) -> list[str]:
    rows = fetchall(
        """
        SELECT DISTINCT industry
        FROM market_snapshot
        WHERE trade_date = ? AND industry <> ''
        ORDER BY industry
        """,
        (trade_date,),
    )
    return [str(r["industry"]).strip() for r in rows if str(r["industry"]).strip()]