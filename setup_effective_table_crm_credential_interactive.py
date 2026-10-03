# 作用：交互式写入舜鼎 CRM 凭据到 Windows Credential Manager（不 echo、不落盘、不打印 secret）。
# 何时改：target 名变化时。禁止纳入定时任务。

from __future__ import annotations

import getpass
import sys

from credential_provider import CREDENTIAL_TARGET_CRM, CredentialError, get_crm_credentials, write_credential


def main() -> int:
    if sys.platform != "win32":
        print("UNSUPPORTED_PLATFORM", flush=True)
        return 2
    print(f"目标: {CREDENTIAL_TARGET_CRM}", flush=True)
    用户 = input("CRM 用户名: ").strip()
    密码 = getpass.getpass("CRM 密码（不回显）: ")
    if not 用户 or not 密码:
        print("INPUT_INCOMPLETE", flush=True)
        return 3
    write_credential(CREDENTIAL_TARGET_CRM, 用户, 密码)
    try:
        读用户, 读密码 = get_crm_credentials()
    except CredentialError:
        print("VERIFY_FAILED", flush=True)
        return 4
    if 读用户 != 用户 or 读密码 != 密码:
        print("VERIFY_MISMATCH", flush=True)
        return 5
    print("SETUP_OK", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
