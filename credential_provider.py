# 作用：信息流有效表生产凭据集中读取（Windows Credential Manager）；禁止日志/异常携带 secret。
# 何时改：Credential target 命名或读取策略变化时。禁止 fallback 到源码默认值。

from __future__ import annotations

import ctypes
import sys
from ctypes import Structure, byref, c_byte, c_void_p
from ctypes import wintypes
from typing import Protocol

CREDENTIAL_TARGET_CRM = "AI-MAIYA/EffectiveTableCRM"
CREDENTIAL_TARGET_EXCEL = "AI-MAIYA/EffectiveTableExcel"

CREDENTIAL_NOT_FOUND = "CREDENTIAL_NOT_FOUND"
CREDENTIAL_READ_FAILED = "CREDENTIAL_READ_FAILED"

CRED_TYPE_GENERIC = 1
CRED_PERSIST_CURRENT_USER = 1


class CredentialError(RuntimeError):
    """凭据缺失或读取失败（消息不含 secret）。"""


class _FILETIME(Structure):
    _fields_ = [
        ("dwLowDateTime", wintypes.DWORD),
        ("dwHighDateTime", wintypes.DWORD),
    ]


class _CREDENTIAL(Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD),
        ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR),
        ("LastWritten", _FILETIME),
        ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.POINTER(c_byte)),
        ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD),
        ("Attributes", c_void_p),
        ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


def _advapi32():
    if sys.platform != "win32":
        raise OSError("Credential Manager 仅支持 win32")
    return ctypes.windll.advapi32


def _blob_to_text(凭据: _CREDENTIAL) -> str:
    if not 凭据.CredentialBlob or 凭据.CredentialBlobSize == 0:
        return ""
    原始 = bytes(凭据.CredentialBlob[i] for i in range(凭据.CredentialBlobSize))
    return 原始.decode("utf-16-le").rstrip("\x00")


def _text_to_blob(文本: str) -> tuple[ctypes.Array[c_byte], int]:
    编码 = 文本.encode("utf-16-le") + b"\x00\x00"
    缓冲 = (c_byte * len(编码)).from_buffer_copy(编码)
    return 缓冲, len(编码)


def _read_generic(target: str) -> tuple[str | None, str | None]:
    if sys.platform != "win32":
        return None, None
    指针 = ctypes.POINTER(_CREDENTIAL)()
    成功 = _advapi32().CredReadW(target, CRED_TYPE_GENERIC, 0, byref(指针))
    if not 成功:
        return None, None
    try:
        凭据 = 指针.contents
        用户名 = (凭据.UserName or "").strip() or None
        密码 = _blob_to_text(凭据).strip() or None
        return 用户名, 密码
    except Exception:
        return None, None
    finally:
        _advapi32().CredFree(指针)


def write_credential(target: str, username: str, secret: str) -> None:
    """仅 setup / 迁移脚本调用；禁止业务代码写入。"""
    if sys.platform != "win32":
        raise OSError("Credential Manager 仅支持 win32")
    缓冲, 大小 = _text_to_blob(secret)
    凭据 = _CREDENTIAL()
    凭据.Type = CRED_TYPE_GENERIC
    凭据.TargetName = target
    凭据.Persist = CRED_PERSIST_CURRENT_USER
    凭据.UserName = username
    凭据.CredentialBlobSize = 大小
    凭据.CredentialBlob = ctypes.cast(缓冲, ctypes.POINTER(c_byte))
    成功 = _advapi32().CredWriteW(byref(凭据), 0)
    if not 成功:
        raise CredentialError(CREDENTIAL_READ_FAILED)


def get_credential(name: str) -> str:
    """
    读取单一 secret（Excel 等）。
    name: EffectiveTableExcel
    """
    if name == "EffectiveTableExcel":
        用户, 密码 = _read_generic(CREDENTIAL_TARGET_EXCEL)
        if not 密码 or not str(密码).strip():
            if 用户 is None and 密码 is None:
                raise CredentialError(CREDENTIAL_NOT_FOUND)
            raise CredentialError(CREDENTIAL_NOT_FOUND)
        return str(密码).strip()
    raise CredentialError(CREDENTIAL_NOT_FOUND)


class CredentialProvider(Protocol):
    def read_crm_username(self) -> str | None: ...

    def read_crm_password(self) -> str | None: ...

    def read_excel_password(self) -> str | None: ...


class WindowsCredentialProvider:
    def read_crm_username(self) -> str | None:
        用户, _ = _read_generic(CREDENTIAL_TARGET_CRM)
        return 用户

    def read_crm_password(self) -> str | None:
        _, 密码 = _read_generic(CREDENTIAL_TARGET_CRM)
        return 密码

    def read_excel_password(self) -> str | None:
        _, 密码 = _read_generic(CREDENTIAL_TARGET_EXCEL)
        return 密码


_default_provider: CredentialProvider | None = None


def set_credential_provider(provider: CredentialProvider | None) -> None:
    global _default_provider
    _default_provider = provider


def _provider() -> CredentialProvider:
    return _default_provider or WindowsCredentialProvider()


def get_crm_credentials(
    *,
    provider: CredentialProvider | None = None,
) -> tuple[str, str]:
    来源 = provider or _provider()
    try:
        用户名 = 来源.read_crm_username()
        密码 = 来源.read_crm_password()
    except Exception:
        raise CredentialError(CREDENTIAL_READ_FAILED) from None
    if not 用户名 or not str(用户名).strip():
        raise CredentialError(CREDENTIAL_NOT_FOUND)
    if not 密码 or not str(密码).strip():
        raise CredentialError(CREDENTIAL_NOT_FOUND)
    return str(用户名).strip(), str(密码).strip()


def get_excel_password(*, provider: CredentialProvider | None = None) -> str:
    来源 = provider or _provider()
    try:
        密码 = 来源.read_excel_password()
    except Exception:
        raise CredentialError(CREDENTIAL_READ_FAILED) from None
    if not 密码 or not str(密码).strip():
        raise CredentialError(CREDENTIAL_NOT_FOUND)
    return str(密码).strip()
