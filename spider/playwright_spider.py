"""
Windows 爬虫脚本：抓取量化网页表格数据并保存。

环境变量:
    CRAWLER_URL     目标 URL (必填)
    CRAWLER_USER    HTTP 基本认证用户名 (必填)
    CRAWLER_PASS    HTTP 基本认证密码 (必填)
    MASTER_FILE     保存路径 (默认: Whole Market.xlsx)
"""
import io
import re
import os
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

# ================= 基础配置 =================
TARGET_URL = os.environ.get("CRAWLER_URL", "")
USERNAME = os.environ.get("CRAWLER_USER", "")
PASSWORD = os.environ.get("CRAWLER_PASS", "")
MASTER_FILE = os.environ.get("MASTER_FILE", "Whole Market.xlsx")
# ============================================


def extract_date(url: str) -> str:
    match = re.search(r'/(\d{4})(\d{4})/', url)
    return match.group(2) if match else "未知日期"


def save_to_master_excel(df: pd.DataFrame, file_path: str, sheet_name: str):
    if not os.path.exists(file_path):
        df.to_excel(file_path, sheet_name=sheet_name, index=False)
        print(f"[+] 已新建 {file_path} -> Sheet: {sheet_name}")
    else:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"[+] 已更新 {file_path} -> Sheet: {sheet_name}")


def fetch_and_sync_data():
    if not TARGET_URL or not USERNAME or not PASSWORD:
        print("[-] 请设置环境变量 CRAWLER_URL, CRAWLER_USER, CRAWLER_PASS")
        return

    sheet_name = extract_date(TARGET_URL)
    print(f"[*] 开始抓取 {sheet_name} 的数据...")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox',
                '--disable-gpu',
                '--disable-dev-shm-usage',
                '--window-size=1920,1080',
                '--disable-blink-features=AutomationControlled',
            ],
        )

        context = browser.new_context(
            http_credentials={"username": USERNAME, "password": PASSWORD},
            user_agent=(
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            ),
            viewport={"width": 1920, "height": 1080},
        )

        page = context.new_page()

        try:
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(3000)

            html = page.content()
            print(f"[DEBUG] 页面标题: {page.title()}")
            print(f"[DEBUG] 页面长度: {len(html)} 字符")
            table_count = html.count("<table")
            print(f"[DEBUG] <table> 标签数: {table_count}")
            if table_count == 0:
                print(f"[DEBUG] 页面开头 800 字符:\n{html[:800]}")

            tables = pd.read_html(io.StringIO(html))
            if not tables:
                print("[-] 未找到表格数据。")
                return

            df = tables[0]
            print(f"[*] 抓取到 {len(df)} 条数据")

            save_to_master_excel(df, MASTER_FILE, sheet_name)
            print(f"[+] 任务完成！")

        except Exception as e:
            print(f"[-] 运行出错: {e}")
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    fetch_and_sync_data()
