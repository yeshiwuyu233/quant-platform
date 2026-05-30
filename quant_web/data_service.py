"""
数据访问层 — 统一管理 xlsx/JSON 文件读取与缓存。
"""
import glob
import json
import os
import re
from datetime import datetime
from typing import Any, Optional

import numpy as np
import pandas as pd

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ── 轻量 TTL 缓存 ──

def ttl_cache(seconds: int):
    """Simple TTL cache decorator — no external dependencies."""
    def decorator(func):
        cache = {}

        def wrapper(*args, **kwargs):
            key = (args, tuple(sorted(kwargs.items())))
            now = datetime.now().timestamp()
            if key in cache:
                val, ts = cache[key]
                if now - ts < seconds:
                    return val
            result = func(*args, **kwargs)
            cache[key] = (result, now)
            return result
        return wrapper
    return decorator

# 追踪表列位置 (Excel col 0 是 index)
_COL = {
    "date": 1, "acc_08": 2, "acc_12": 3, "all": 4,
    "nc": 5, "t3": 6, "cold_alpha": 7,
    "cold_count": 8, "top3_count": 9,
    "next_08": 10, "next_10": 11, "next_12": 12,
}


# ── 底层 IO ──

def _latest(pattern):
    files = glob.glob(os.path.join(PROJECT_ROOT, pattern))
    return max(files, key=os.path.getmtime) if files else None


def _read_safe(path, sheet_name=0, **kwargs):
    try:
        return pd.read_excel(path, sheet_name=sheet_name, **kwargs)
    except Exception:
        return None


def _table(df):
    if df is None or df.empty:
        return None

    def clean(v):
        if pd.isna(v):
            return ""
        if isinstance(v, float) and v == int(v):
            return int(v)
        return v

    rows = [{k: clean(v) for k, v in row.items()}
            for row in df.to_dict(orient="records")]
    return {"cols": list(rows[0].keys()), "rows": rows}


def _first_row(df):
    if df is None or df.empty:
        return None
    clean = df.dropna(how="all")
    return clean.iloc[0] if not clean.empty else None


def _parse_pct(s):
    if pd.isna(s) or not str(s).strip():
        return None
    m = re.search(r"([-+]?\d+\.?\d*)%", str(s))
    return float(m.group(1)) if m else None


def _parse_ratio(s):
    if pd.isna(s) or not str(s).strip():
        return None, None
    m = re.search(r"\((\d+)/(\d+)\)", str(s))
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


# ── 公开工具函数（供 batch 脚本复用） ──

def df_to_table(df):
    """Convert DataFrame → {{cols, rows}} for JSON serialization."""
    return _table(df)


def parse_pct(s):
    """Extract percentage value from string like '49.25%(33/67)' → 49.25"""
    return _parse_pct(s)


def parse_ratio(s):
    """Extract (success, total) from string like '49.25%(33/67)' → (33, 67)"""
    return _parse_ratio(s)


def native_type(v):
    """Convert numpy types to native Python types for JSON serialization."""
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return float(v) if not np.isnan(v) else None
    return v


def get_adjacent_dates(date, all_dates):
    """Return (prev_date, next_date) for date navigation."""
    if not all_dates or date not in all_dates:
        return None, None
    idx = all_dates.index(date)
    prev_date = all_dates[idx - 1] if idx > 0 else None
    next_date = all_dates[idx + 1] if idx < len(all_dates) - 1 else None
    return prev_date, next_date


def _to_scalar(v):
    if pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return int(v) if v == int(v) else v
    return v


def _mmdd_label(mmdd):
    if not mmdd or len(mmdd) != 4:
        return mmdd
    return f"{mmdd[:2]}/{mmdd[2:]}"


def _xlsx_to_json(path):
    if not path:
        return None
    return path.rsplit('.', 1)[0] + '.json'


def _try_json(path):
    """Read JSON report if available; returns (full_dict, True) or (None, False)."""
    jp = _xlsx_to_json(path)
    if jp and os.path.exists(jp):
        try:
            with open(jp, 'r', encoding='utf-8') as f:
                return json.load(f), True
        except Exception:
            pass
    return None, False


def _read_tracking_from_xlsx(path):
    """Fallback xlsx reading for tracking data."""
    df = _read_safe(path, sheet_name="0.每日追踪总表", header=[0, 1])
    row = _first_row(df)
    if row is None:
        return None

    def v(pos):
        try:
            return _to_scalar(row.iloc[pos])
        except Exception:
            return None

    date_val = v(_COL["date"])
    full = str(int(float(date_val))) if date_val is not None else ""
    date_str = full[4:] if len(full) >= 8 else full

    return {
        "date_raw": date_str,
        "date": _mmdd_label(date_str),
        "acc_08_raw": v(_COL["acc_08"]),
        "acc_12_raw": v(_COL["acc_12"]),
        "all_raw": v(_COL["all"]),
        "nc_raw": v(_COL["nc"]),
        "t3_raw": v(_COL["t3"]),
        "cold_alpha_raw": v(_COL["cold_alpha"]),
        "cold_stock_count": v(_COL["cold_count"]),
        "top3_stock_count": v(_COL["top3_count"]),
        "next_08": v(_COL["next_08"]),
        "next_10": v(_COL["next_10"]),
        "next_12": v(_COL["next_12"]),
    }


def _read_tracking_from_path(path):
    if not path:
        return None

    # Try JSON first
    report, ok = _try_json(path)
    if ok and report and 'tracking' in report:
        return report['tracking']

    return _read_tracking_from_xlsx(path)


# ── 公开 API ──

@ttl_cache(60)
def get_available_dates() -> list[str]:
    """返回 Whole Market 中所有交易日日期列表 (已排序)。"""
    market_path = os.path.join(PROJECT_ROOT, "Whole Market.xlsx")
    try:
        xls = pd.ExcelFile(market_path)
        dates = [s for s in xls.sheet_names if s.isdigit() and len(s) == 4]
        return sorted(dates)
    except Exception:
        return []


@ttl_cache(30)
def get_latest_tracking() -> Optional[dict[str, Any]]:
    return _read_tracking_from_path(_latest("*量化复盘报告*.xlsx"))


@ttl_cache(60)
def get_tracking_data(date: str) -> Optional[dict[str, Any]]:
    exact = os.path.join(PROJECT_ROOT, f"{date}量化复盘报告.xlsx")
    if os.path.exists(exact):
        return _read_tracking_from_path(exact)
    files = glob.glob(os.path.join(PROJECT_ROOT, f"*{date}*量化复盘报告*"))
    if files:
        return _read_tracking_from_path(max(files, key=os.path.getmtime))
    return None


@ttl_cache(60)
def get_backtest_sheets(date: Optional[str] = None) -> dict[str, Any]:
    """读取指定日期的复盘报告。date=None 则取最新。"""
    path = None
    if date:
        exact = os.path.join(PROJECT_ROOT, f"{date}量化复盘报告.xlsx")
        if os.path.exists(exact):
            path = exact
        else:
            files = glob.glob(os.path.join(PROJECT_ROOT, f"*{date}*量化复盘报告*"))
            if files:
                path = max(files, key=os.path.getmtime)
    if not path:
        path = _latest("*量化复盘报告*.xlsx")
    if not path:
        return {}

    # Try JSON first
    report, ok = _try_json(path)
    if ok and report and 'sheets' in report:
        return {k: v for k, v in report['sheets'].items() if v is not None}

    # Fall back to xlsx
    raw = {
        "backtest": _read_safe(path, "1.回测明细(跨日合并)"),
        "win_rates": _read_safe(path, "2.三阶胜率全景对比"),
        "industry_dist": _read_safe(path, "3.回测行业分布(含名单)"),
        "hot_split": _read_safe(path, "4.并列热门拆分对比"),
        "today_dist": _read_safe(path, "5.当日最新策略分布"),
        "today_industry": _read_safe(path, "6.当日最新行业热度"),
    }
    return {k: _table(v) for k, v in raw.items() if _table(v) is not None}


@ttl_cache(60)
def get_available_weekly_dates() -> list[str]:
    """返回所有礼拜攻势报告对应的 MMDD 日期列表。"""
    dates: set[str] = set()
    for f in glob.glob(os.path.join(PROJECT_ROOT, "*的选股策略礼拜攻势.xlsx")):
        d = os.path.basename(f)[:4]
        if d.isdigit():
            dates.add(d)
    return sorted(dates)


def _extract_trend_record(mmdd, summary_rows):
    """从 summary 行列表提取单条趋势记录。"""
    record = {"date": _mmdd_label(mmdd), "date_raw": mmdd}
    for row in summary_rows:
        group = row.get("策略分组", "")
        suffix = group.replace("指标大于", "gt_").replace(".", "_")
        count = int(row.get("入选股票数", 0))
        ret_str = str(row.get("平均持仓回报", "0%"))
        win_str = str(row.get("策略胜率(>0%)", "0%"))

        try:
            ret_val = float(ret_str.replace("%", ""))
        except ValueError:
            ret_val = 0.0
        try:
            win_val = float(win_str.replace("%", ""))
        except ValueError:
            win_val = 0.0

        record[f"{suffix}_回报"] = ret_val
        record[f"{suffix}_胜率"] = win_val
        record[f"{suffix}_股票数"] = count
    return record


@ttl_cache(120)
def get_weekly_trend_data() -> list[dict[str, Any]]:
    """聚合所有日期的礼拜攻势数据，生成趋势图用时间序列。
    优先读 batch_weekly.py 生成的缓存文件 weekly_trend.json。
    """
    cache_path = os.path.join(PROJECT_ROOT, 'weekly_trend.json')
    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass

    # 缓存不存在时回退：逐文件解析
    records = []
    files = sorted(glob.glob(os.path.join(PROJECT_ROOT, "*的选股策略礼拜攻势.json")))
    for f in files:
        base = os.path.basename(f)
        mmdd = base[:4]
        if not mmdd.isdigit():
            continue
        try:
            with open(f, 'r', encoding='utf-8') as jf:
                data = json.load(jf)
        except Exception:
            continue

        if 'standard' in data:
            std_summary = data.get('standard', {}).get('summary', [])
            if std_summary:
                records.append(_extract_trend_record(mmdd, std_summary))
        else:
            summary = data.get("整体回报总结")
            if not summary or not summary.get("rows"):
                continue
            records.append(_extract_trend_record(mmdd, summary["rows"]))

    return records


@ttl_cache(60)
def get_weekly_data(date: Optional[str] = None) -> dict[str, Any]:
    path = None
    if date:
        exact = os.path.join(PROJECT_ROOT, f"{date}的选股策略礼拜攻势.xlsx")
        if os.path.exists(exact):
            path = exact
    if not path:
        path = _latest("*选股策略礼拜攻势*.xlsx")
    if not path:
        return {}

    # Try JSON first
    report, ok = _try_json(path)
    if ok and report:
        return report

    # Fall back to xlsx
    try:
        xls = pd.ExcelFile(path)
    except Exception:
        return {}
    sheets = {}
    for name in xls.sheet_names:
        t = _table(_read_safe(path, name))
        if t:
            sheets[name] = t
    return sheets


def extract_weekly_view(report, strategy='standard'):
    """从双轨 report 中提取指定策略的 sheets + chart_data。

    支持新旧两种 JSON 格式，返回 (sheets: dict, chart_data: list | None).
    """
    if not report:
        return {}, None

    # ── 新格式：嵌套结构 ──
    if 'standard' in report:
        strat = report.get(strategy, report.get('standard', {}))
        sheets = strat.get('sheets', {})
        chart_data = strat.get('summary', [])
        return sheets, chart_data

    # ── 旧格式：扁平结构 ──
    chart_data = None
    if "整体回报总结" in report:
        chart_data = report["整体回报总结"].get("rows", [])
    return report, chart_data


def _normalize_history(raw_list: list) -> list:
    """Convert raw tracking format (acc_08_raw strings) → parsed history format."""
    result = []
    for r in raw_list:
        # Already in parsed format — pass through
        if 'acc_08' in r and 'all_pct' in r:
            result.append(r)
            continue

        all_raw = r.get('all_raw', '')
        all_pct = _parse_pct(all_raw)
        all_s, all_t = _parse_ratio(all_raw)

        result.append({
            'date': r.get('date', ''),
            'date_raw': r.get('date_raw', ''),
            'acc_08': _parse_pct(r.get('acc_08_raw', '')),
            'acc_12': _parse_pct(r.get('acc_12_raw', '')),
            'all_pct': all_pct,
            'all_success': all_s,
            'all_total': all_t,
            'cold_alpha': _parse_pct(r.get('cold_alpha_raw', '')),
            'next_10': int(r.get('next_10', 0)),
        })
    return result


@ttl_cache(120)
def get_history_data() -> list[dict[str, Any]]:
    # Try history.json first
    history_path = os.path.join(PROJECT_ROOT, 'history.json')
    if os.path.exists(history_path):
        try:
            with open(history_path, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            return _normalize_history(raw)
        except Exception:
            pass

    # Fall back to xlsx aggregation
    files = glob.glob(os.path.join(PROJECT_ROOT, "*量化复盘报告*.xlsx"))
    records = []
    for f in files:
        df = _read_safe(f, sheet_name="0.每日追踪总表", header=[0, 1])
        row = _first_row(df)
        if row is None:
            continue

        def v(pos):
            try:
                return row.iloc[pos]
            except Exception:
                return None

        date_raw = v(_COL["date"])
        try:
            date_str = str(int(float(str(date_raw))))
            dt = datetime.strptime(date_str, "%Y%m%d")
            date_label = dt.strftime("%m/%d")
        except Exception:
            date_label = str(date_raw)

        acc_08 = _parse_pct(v(_COL["acc_08"]))
        acc_12 = _parse_pct(v(_COL["acc_12"]))
        all_str = v(_COL["all"])
        all_pct = _parse_pct(all_str)
        all_s, all_t = _parse_ratio(all_str)
        cold_alpha = _parse_pct(v(_COL["cold_alpha"]))
        next_10 = v(_COL["next_10"])

        records.append({
            "date": date_label,
            "date_raw": date_str,
            "acc_08": acc_08,
            "acc_12": acc_12,
            "all_pct": all_pct,
            "all_success": all_s,
            "all_total": all_t,
            "cold_alpha": cold_alpha,
            "next_10": int(next_10) if pd.notna(next_10) else 0,
        })
    records.sort(key=lambda r: r["date_raw"])
    seen = set()
    deduped = []
    for r in records:
        if r["date_raw"] not in seen:
            seen.add(r["date_raw"])
            deduped.append(r)
    return deduped


# ── 筛选器（全市场股票） ──

def get_screener_data(date: Optional[str] = None) -> list[dict[str, Any]]:
    """返回 Whole Market 中指定日期的全量股票数据。"""
    return _query_market(date, acc_min=0, ind_min=-999, industry="")


@ttl_cache(300)
def get_screener_count(date: Optional[str] = None) -> int:
    """返回全市场股票总数（只读单列统计行数）。"""
    if not date:
        dates = get_available_dates()
        date = dates[-1] if dates else None
    if not date:
        return 0
    try:
        df = pd.read_excel(os.path.join(PROJECT_ROOT, "Whole Market.xlsx"),
                           sheet_name=date, usecols=['代码'])
        return len(df)
    except Exception:
        return 0


def query_screener(date: str, acc_min: float = 0, ind_min: float = -999,
                   industry: str = "", page: int = 1, per_page: int = 100) -> dict:
    """分页查询全市场股票。返回 {rows, total, page, per_page, pages}。"""
    rows = _query_market(date, acc_min=acc_min, ind_min=ind_min, industry=industry)
    total = len(rows)
    pages = max(1, (total + per_page - 1) // per_page)
    start = (page - 1) * per_page
    end = start + per_page
    return {
        "rows": rows[start:end],
        "total": total,
        "page": page,
        "per_page": per_page,
        "pages": pages,
    }


@ttl_cache(120)
def _read_market_sheet(date: str):
    """读取 Whole Market 单日 sheet，返回 cleaned DataFrame（2min 缓存）。"""
    if not date:
        dates = get_available_dates()
        date = dates[-1] if dates else None
    if not date:
        return None
    path = os.path.join(PROJECT_ROOT, "Whole Market.xlsx")
    try:
        df = pd.read_excel(path, sheet_name=date)
    except Exception:
        return None
    df['准确率'] = pd.to_numeric(df['准确率'], errors='coerce')
    df['今日指标'] = pd.to_numeric(df['今日指标'], errors='coerce')
    return df


def _query_market(date: str, acc_min: float = 0, ind_min: float = -999,
                  industry: str = "") -> list[dict[str, Any]]:
    """按条件筛选全市场股票，返回 dict 列表。"""
    df = _read_market_sheet(date)
    if df is None:
        return []

    mask = (df['准确率'] >= acc_min) & (df['今日指标'] >= ind_min)
    if industry:
        mask &= (df['行业'].astype(str).str.strip() == industry)
    df = df[mask]

    def _extract_return(h):
        if pd.isna(h):
            return ""
        matches = re.findall(r'\(([+-]?\d+\.?\d*)%\)', str(h))
        if not matches:
            return ""
        return f"{float(matches[-1]):+.2f}%"

    df = df.copy()
    df['今日收益率'] = df['指标历史'].apply(_extract_return)

    display_cols = ['代码', '全称', '行业', '准确率', '今日指标', '今日收益率']
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        d = {}
        for c in display_cols:
            v = row.get(c)
            if pd.isna(v):
                d[c] = ""
            elif isinstance(v, (np.floating, float)):
                d[c] = round(float(v), 2)
            elif isinstance(v, (np.integer,)):
                d[c] = int(v)
            else:
                d[c] = str(v)
        rows.append(d)

    return rows


@ttl_cache(300)
def get_all_industries() -> list[str]:
    """从 Whole Market 最新单日 sheet 获取行业列表（只读行业列）。"""
    dates = get_available_dates()
    if not dates:
        return []
    path = os.path.join(PROJECT_ROOT, "Whole Market.xlsx")
    try:
        df = pd.read_excel(path, sheet_name=dates[-1], usecols=['行业'])
    except Exception:
        return []
    inds = [str(v).strip() for v in df['行业'].dropna().unique() if str(v).strip()]
    return sorted(inds)
