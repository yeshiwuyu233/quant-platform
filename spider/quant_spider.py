"""
服务器爬虫脚本: 用 DrissionPage 抓取量化网页表格数据。

环境变量:
    CRAWLER_URL     目标 URL (必填)
    CRAWLER_USER    HTTP 基本认证用户名 (必填)
    CRAWLER_PASS    HTTP 基本认证密码 (必填)
    SAVE_DIR        保存目录 (默认 ./data)
    CHROME_PATH     浏览器路径 (可选, 默认自动查找)
"""
import io
import re
import os
import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions

# ================= 基础配置 =================
TARGET_URL = os.environ.get("CRAWLER_URL", "")
USERNAME = os.environ.get("CRAWLER_USER", "")
PASSWORD = os.environ.get("CRAWLER_PASS", "")

SAVE_DIR = os.environ.get("SAVE_DIR", "./data")
CHROME_PATH = os.environ.get("CHROME_PATH", "/usr/bin/google-chrome-stable")
# ============================================


def extract_date_from_url(url: str) -> str:
    match = re.search(r'/(\d{8})/', url)
    return match.group(1) if match else "download_data"


def fetch_and_sync_data():
    if not TARGET_URL or not USERNAME or not PASSWORD:
        print("[-] 请设置环境变量 CRAWLER_URL, CRAWLER_USER, CRAWLER_PASS")
        return

    file_date = extract_date_from_url(TARGET_URL)
    filename = f"{file_date}.xlsx"

    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    filepath = os.path.join(SAVE_DIR, filename)

    co = ChromiumOptions()
    if CHROME_PATH and os.path.exists(CHROME_PATH):
        co.set_browser_path(CHROME_PATH)

    co.set_argument('--headless')
    co.set_argument('--no-sandbox')
    co.set_argument('--disable-dev-shm-usage')
    co.set_argument('--disable-gpu')

    page = None
    try:
        page = ChromiumPage(co)
        auth_url = TARGET_URL.replace("https://", f"https://{USERNAME}:{PASSWORD}@")
        page.get(auth_url)

        print(f"[*] 正在访问页面并提取数据，准备保存至 {filepath}...")
        page.wait.load_start()
        page.wait(2)

        tables = pd.read_html(io.StringIO(page.html))
        if not tables:
            print("[-] 未能识别到表格数据。")
            return

        df = tables[0]
        df.to_excel(filepath, index=False)
        print(f"[+] 任务完成！数据已保存至新文件: {filepath}")

    except Exception as e:
        print(f"[-] 运行出错: {e}")
    finally:
        if page:
            page.quit()


if __name__ == "__main__":
    fetch_and_sync_data()
