"""每日量化流水线 —— 供 Docker 定时任务调用。

运行顺序:
  1. daily_spider.py  爬取今日数据（URL 自动按日期构造）
  2. batch_backtest.py  批量回测所有新日期
  3. batch_weekly.py    批量生成所有新礼拜攻势

用法:
  docker exec quant-web python /app/quant_web/run_pipeline.py
"""
import datetime
import logging
import smtplib
import subprocess
import sys
import os
import json
import glob
from email.mime.text import MIMEText

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def is_trading_day() -> bool:
    import pandas as pd
    import akshare as ak

    today = datetime.datetime.now().date()
    try:
        log.info(">>> 核对A股交易日历...")
        calendar_df = ak.tool_trade_date_hist_sina()
        trade_dates = pd.to_datetime(calendar_df["trade_date"]).dt.date.tolist()
        return today in trade_dates
    except Exception as e:
        log.warning(f"[!] 获取交易日历失败: {e}，按工作日判断")
        return today.weekday() < 5


def crawl_today():
    log.info(">>> 爬取今日数据...")
    crawler = os.path.join(PROJECT_ROOT, "quant_web", "daily_spider.py")
    env = os.environ.copy()
    env['PYTHONPATH'] = os.path.join(PROJECT_ROOT, 'quant_web')
    subprocess.run(
        [sys.executable, crawler],
        check=True, cwd=PROJECT_ROOT, env=env,
    )


def run_backtest():
    log.info(">>> 批量回测...")
    script = os.path.join(PROJECT_ROOT, "quant_web", "batch_backtest.py")
    env = os.environ.copy()
    env['PYTHONPATH'] = os.path.join(PROJECT_ROOT, 'quant_web')
    subprocess.run([sys.executable, script], check=True, cwd=PROJECT_ROOT, env=env)


def run_weekly():
    log.info(">>> 批量礼拜攻势...")
    script = os.path.join(PROJECT_ROOT, "quant_web", "batch_weekly.py")
    env = os.environ.copy()
    env['PYTHONPATH'] = os.path.join(PROJECT_ROOT, 'quant_web')
    subprocess.run([sys.executable, script], check=True, cwd=PROJECT_ROOT, env=env)


def send_notification(success: bool, summary: str):
    """流水线结束后发送邮件通知"""
    smtp_user = os.environ.get("SMTP_USER", "")
    smtp_pass = os.environ.get("SMTP_PASS", "")
    smtp_host = os.environ.get("SMTP_HOST", "")
    smtp_port = int(os.environ.get("SMTP_PORT", "465"))

    if not smtp_user or not smtp_pass:
        log.info("[通知] SMTP 未配置，跳过邮件通知")
        return

    notify_to = os.environ.get("NOTIFY_EMAIL", smtp_user)
    date_str = datetime.datetime.now().strftime("%m/%d")
    status = "✅ 成功" if success else "❌ 失败"

    body = f"""量化流水线 — {date_str} 执行报告

状态: {status}

{summary}

—— 量化分析系统自动通知"""

    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = f"量化流水线 {date_str} {status}"
        msg["From"] = smtp_user
        msg["To"] = notify_to
        with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=10) as s:
            s.login(smtp_user, smtp_pass)
            s.send_message(msg)
        log.info(f"[通知] 邮件已发送至 {notify_to}")
    except Exception as e:
        log.warning(f"[通知] 邮件发送失败: {e}")


def _count_today_results():
    """统计今日最新回测和礼拜攻势结果"""
    today_mmdd = datetime.datetime.now().strftime("%m%d")
    lines = []

    # 最新回测
    bt_files = sorted(glob.glob(os.path.join(PROJECT_ROOT, "*量化复盘报告.json")))
    if bt_files:
        latest = bt_files[-1]
        try:
            with open(latest, 'r', encoding='utf-8') as f:
                data = json.load(f)
            t = data.get('tracking', {})
            lines.append(f"回测 {t.get('date', '?')}")
            lines.append(f"  全样本胜率: {t.get('all_raw', '?')}")
            lines.append(f"  冷门Alpha: {t.get('cold_alpha_raw', '?')}")
            lines.append(f"  明日>1.0 达标: {t.get('next_10', '?')} 只")
        except Exception:
            pass

    # 最新礼拜攻势
    wk_files = sorted(glob.glob(os.path.join(PROJECT_ROOT, "*选股策略礼拜攻势.json")))
    if wk_files:
        latest = wk_files[-1]
        try:
            with open(latest, 'r', encoding='utf-8') as f:
                data = json.load(f)
            std = data.get('standard', {}).get('summary', [])
            top = data.get('top_industries', {}).get('summary', [])
            cold = data.get('cold_industry', {}).get('summary', [])
            lines.append(f"\n礼拜攻势 {data.get('meta', {}).get('date', '?')}")
            for s in std:
                lines.append(f"  全市场-{s.get('策略分组', '?')}: {s.get('入选股票数', 0)}只, 回报{s.get('平均持仓回报', '?')}")
            lines.append(f"  前三行业: {data.get('top_industries', {}).get('meta', {}).get('industries', [])}")
            lines.append(f"  冷门行业: {data.get('meta', {}).get('cold_industries', [])}")
        except Exception:
            pass

    return "\n".join(lines) if lines else "暂无数据"


def main():
    if not is_trading_day():
        date_str = datetime.datetime.now().strftime("%Y-%m-%d")
        log.info(f"[{date_str}] 非交易日，跳过。")
        return

    start = datetime.datetime.now()
    log.info(f"\n[{start.strftime('%Y-%m-%d %H:%M:%S')}] 开始量化流水线\n")

    success = True
    try:
        crawl_today()
    except Exception as e:
        log.error(f"爬取失败: {e}")
        success = False

    if success:
        try:
            run_backtest()
        except Exception as e:
            log.error(f"回测失败: {e}")
            success = False

    if success:
        try:
            run_weekly()
        except Exception as e:
            log.error(f"礼拜攻势失败: {e}")
            success = False

    elapsed = (datetime.datetime.now() - start).total_seconds()
    summary = _count_today_results()
    summary += f"\n\n总耗时: {elapsed:.0f} 秒"

    log.info(f"\n流水线{'完成' if success else '异常'}！耗时 {elapsed:.0f}s")
    send_notification(success, summary)


if __name__ == "__main__":
    main()
