"""
测试爬虫脚本: DrissionPage 爬取 - 从环境变量读取配置。

环境变量:
    CRAWLER_URL     目标 URL (必填)
    CRAWLER_USER    HTTP 基本认证用户名 (必填)
    CRAWLER_PASS    HTTP 基本认证密码 (必填)
    SAVE_DIR        保存目录 (默认 ./data)
    BROWSER_PATH    浏览器路径 (可选)
"""
import io
import re
import os
import pandas as pd
from datetime import datetime
from DrissionPage import ChromiumPage, ChromiumOptions

# ================= 基础配置 =================
CRAWLER_URL = os.environ.get("CRAWLER_URL", "")
CRAWLER_USER = os.environ.get("CRAWLER_USER", "")
CRAWLER_PASS = os.environ.get("CRAWLER_PASS", "")

today_str = datetime.now().strftime("%Y%m%d")
TARGET_URL = os.environ.get("CRAWLER_URL", "")

SAVE_DIR = os.environ.get("SAVE_DIR", "./data")
BROWSER_PATH = os.environ.get("BROWSER_PATH", "")
# ============================================


def extract_date_from_url(url: str) -> str:
    match = re.search(r'/(\d{8})/', url)
    return match.group(1) if match else "download_data"


def fetch_and_sync_data():
    if not CRAWLER_URL or not CRAWLER_USER or not CRAWLER_PASS:
        print("[-] 请设置环境变量 CRAWLER_URL, CRAWLER_USER, CRAWLER_PASS")
        return

    file_date = extract_date_from_url(TARGET_URL)
    filename = f"{file_date}.xlsx"

    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    filepath = os.path.join(SAVE_DIR, filename)

    co = ChromiumOptions()
    if BROWSER_PATH:
        co.set_browser_path(BROWSER_PATH)
    co.set_argument('--headless')
    co.set_argument('--no-sandbox')
    co.set_argument('--disable-gpu')
    co.set_argument('--window-size=1920,1080')
    co.set_user_agent(
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
        '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    )

    page = None
    try:
        page = ChromiumPage(co)
        auth_url = TARGET_URL.replace("https://", f"https://{CRAWLER_USER}:{CRAWLER_PASS}@")
        page.get(auth_url)

        print(f"[*] 正在访问页面，准备保存至 {filepath} ...")
        page.wait.load_start()
        page.wait(3)

        tables = pd.read_html(io.StringIO(page.html))
        if not tables:
            print("[-] 未能识别到表格数据。")
            return

        df = tables[0]
        print(f"[*] 抓取到 {len(df)} 条数据")

        df.to_excel(filepath, index=False)
        print(f"[+] 任务完成！数据已保存至: {filepath}")

    except Exception as e:
        print(f"[-] 运行出错: {e}")
    finally:
        if page:
            page.quit()


if __name__ == "__main__":
    fetch_and_sync_data()
