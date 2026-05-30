import io
import re
import os
import sys
import pandas as pd
from datetime import datetime, timedelta
from DrissionPage import ChromiumPage, ChromiumOptions

# ================= 基础配置 =================
# 自动获取当天的日期，格式为 YYYYMMDD (例如 20260428)
today_str = datetime.now().strftime("%Y%m%d")

# 如果你想爬取"昨天"的数据（比如由于时差或数据更新晚），把上面那行注释掉，用下面这行：
# today_str = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")

# 动态拼接今天的目标 URL
# 从环境变量读取爬虫认证信息（不可硬编码）
CRAWLER_URL = os.environ.get("CRAWLER_URL", "")
CRAWLER_USER = os.environ.get("CRAWLER_USER", "")
CRAWLER_PASS = os.environ.get("CRAWLER_PASS", "")
USERNAME = CRAWLER_USER
PASSWORD = CRAWLER_PASS

TARGET_URL = os.environ.get("CRAWLER_URL", "")

# 目标汇总文件（与 batch_backtest / batch_weekly 共用同一个文件）
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER_FILE = os.path.join(PROJECT_ROOT, "Whole Market.xlsx")

# ============================================

def extract_date(url: str) -> str:
    """从 URL 中提取月日作为 Sheet 名称 (如 0428)"""
    match = re.search(r'/(\d{4})(\d{4})/', url)
    return match.group(2) if match else "未知日期"


def save_to_master_excel(df: pd.DataFrame, file_path: str, sheet_name: str):
    """将数据保存到指定的 Excel Sheet 中，支持追加模式（原子写入，崩溃不坏文件）。"""
    import shutil
    if not os.path.exists(file_path):
        df.to_excel(file_path, sheet_name=sheet_name, index=False)
        print(f"[+] 已新建汇总表 {file_path} 并创建 Sheet: {sheet_name}")
    else:
        tmp = file_path + '.tmp'
        shutil.copy2(file_path, tmp)
        with pd.ExcelWriter(tmp, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        os.replace(tmp, file_path)
        print(f"[+] 已在 {file_path} 中更新/新建 Sheet: {sheet_name}")


def fetch_and_sync_data():
    """抓取数据并同步到 Whole Market.xlsx"""
    sheet_name = extract_date(TARGET_URL)

    co = ChromiumOptions()
    co.set_browser_path('/usr/bin/chromium-browser' if os.path.exists('/usr/bin/chromium-browser') else '/usr/bin/chromium')
    co.set_argument('--headless')
    co.set_argument('--no-sandbox')
    co.set_argument('--disable-gpu')

    page = None
    try:
        page = ChromiumPage(co)
        auth_url = TARGET_URL.replace("https://", f"https://{USERNAME}:{PASSWORD}@")
        page.get(auth_url)

        print(f"[*] 正在抓取页面并提取 {sheet_name} 的数据...")
        page.wait.load_start()
        page.wait(3)

        tables = pd.read_html(io.StringIO(page.html))
        if not tables:
            print("[-] 未能识别到表格数据，可能是该日期的网页不存在。")
            sys.exit(1)

        df = tables[0]

        # 调用保存函数
        save_to_master_excel(df, MASTER_FILE, sheet_name)
        print(f"[+] 任务完成！当前 {today_str} 的数据已汇总至 {MASTER_FILE}")

    except Exception as e:
        print(f"[-] 运行出错: 目标网站可能未更新今日数据，或网络连接异常。详细信息: {e}")
        sys.exit(1)
    finally:
        if page:
            try:
                page.quit()
            except Exception:
                pass

if __name__ == "__main__":
    fetch_and_sync_data()
