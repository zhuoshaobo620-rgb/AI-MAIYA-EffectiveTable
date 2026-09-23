# 作用：一次性将 legacy 内嵌凭据写入 Windows Credential Manager（不打印 secret）。
# 何时改：target 名或迁移源解析规则变化时。禁止纳入定时任务。

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from credential_provider import (
    CREDENTIAL_TARGET_CRM,
    CREDENTIAL_TARGET_EXCEL,
    get_crm_credentials,
    get_excel_password,
    write_credential,
)


def _从日报解析内嵌凭据(日报路径: Path) -> tuple[str, str, str]:
    文本 = 日报路径.read_text(encoding="utf-8")
    用户 = re.search(r'^CRM_USER\s*=\s*["\']([^"\']+)["\']', 文本, re.M)
    密码 = re.search(r'^CRM_PASS\s*=\s*["\']([^"\']+)["\']', 文本, re.M)
    excel = re.search(r'^EXCEL_PASSWORD\s*=\s*["\']([^"\']+)["\']', 文本, re.M)
    if not 用户 or not 密码 or not excel:
        raise SystemExit("LEGACY_PARSE_FAILED")
    return 用户.group(1), 密码.group(1), excel.group(1)


def _校验一致(期望用户: str, 期望密码: str, 期望excel: str) -> None:
    读用户, 读密码 = get_crm_credentials()
    读excel = get_excel_password()
    if 读用户 != 期望用户 or 读密码 != 期望密码 or 读excel != 期望excel:
        raise SystemExit("VERIFY_MISMATCH")


def main() -> int:
    解析 = argparse.ArgumentParser(description="迁移有效表凭据到 Credential Manager")
    解析.add_argument(
        "--legacy-daily-report",
        default=str(Path(__file__).resolve().parent / "daily_report.py"),
        help="仍含内嵌凭据的 daily_report 路径（迁移前）",
    )
    参数 = 解析.parse_args()
    日报 = Path(参数.legacy_daily_report)
    if not 日报.is_file():
        print("DAILY_REPORT_MISSING", flush=True)
        return 2
    用户, 密码, excel = _从日报解析内嵌凭据(日报)
    write_credential(CREDENTIAL_TARGET_CRM, 用户, 密码)
    write_credential(CREDENTIAL_TARGET_EXCEL, "EffectiveTableExcel", excel)
    _校验一致(用户, 密码, excel)
    print("MIGRATION_OK", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
