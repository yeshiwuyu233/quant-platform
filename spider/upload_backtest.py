"""
PC 端上传脚本：登录量化系统并上传 Excel 文件触发回测流水线。

用法:
    python upload_backtest.py <excel文件>

环境变量:
    SERVER_URL      量化平台地址 (默认 http://localhost:5000)
    ADMIN_USER      管理员用户名 (默认 admin)
    ADMIN_PASS      管理员密码 (必填)

示例:
    SERVER_URL=http://example.com ADMIN_PASS=xxx python upload_backtest.py 20260527.xlsx
"""
import re
import sys
import os
from pathlib import Path

import requests

SERVER = os.environ.get("SERVER_URL", "http://localhost:5000")
USERNAME = os.environ.get("ADMIN_USER", "admin")
PASSWORD = os.environ.get("ADMIN_PASS", "")


def extract_date_from_filename(path: str) -> str:
    """从文件名中提取 YYYYMMDD 格式日期。"""
    name = Path(path).stem
    match = re.search(r'(\d{8})', name)
    if not match:
        print("[-] 文件名中未找到 8 位日期 (YYYYMMDD)")
        print("   示例: 20260527.xlsx")
        sys.exit(1)
    return match.group(1)


def main():
    if not PASSWORD:
        print("[-] 请设置环境变量 ADMIN_PASS")
        print("   示例: ADMIN_PASS=xxx python upload_backtest.py <file>")
        sys.exit(1)

    if len(sys.argv) < 2:
        print("用法: python upload_backtest.py <excel文件>")
        sys.exit(1)

    filepath = sys.argv[1]

    if not os.path.exists(filepath):
        print(f"[-] 文件不存在: {filepath}")
        sys.exit(1)

    if not filepath.endswith(".xlsx"):
        print(f"[-] 文件必须是 .xlsx 格式")
        sys.exit(1)

    date_str = extract_date_from_filename(filepath)

    session = requests.Session()

    # ── 1. 登录 ──
    print(f"[*] 登录 {SERVER} ...")
    resp = session.post(
        f"{SERVER}/login",
        data={"username": USERNAME, "password": PASSWORD},
    )
    if resp.url.endswith("/login"):
        print("[-] 登录失败，请检查用户名/密码")
        sys.exit(1)
    print("[+] 登录成功")

    # ── 2. 上传文件 ──
    print(f"[*] 上传 {filepath}，日期: {date_str} ...")
    with open(filepath, "rb") as f:
        resp = session.post(
            f"{SERVER}/admin/upload",
            data={"date": date_str},
            files={"file": (os.path.basename(filepath), f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        )

    result = resp.json()
    if resp.status_code != 200 or not result.get("ok"):
        error = result.get("error", resp.text)
        print(f"[-] 上传失败: {error}")
        sys.exit(1)

    print(f"[+] 上传成功！{result.get('message', '')}")
    print(f"[*] 可登录后台查看进度: {SERVER}/admin/upload")


if __name__ == "__main__":
    main()
