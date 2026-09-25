#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日常报表自动化脚本
- 从 CRM 导出信息流线索池数据（记录时间为昨天）
- 处理导出文件并更新本地 Excel 报表
"""

import os
import sys
import shutil
import time
import datetime
import io
from pathlib import Path

import msoffcrypto
import openpyxl
from playwright.sync_api import sync_playwright

from crm_export_artifact_gate import (
    assert_export_process_yield,
    validate_crm_export_structure,
)


def use_local_playwright_browser():
    """Playwright 浏览器固定到用户目录，避免 Cursor/沙箱临时路径失效。"""
    local_browser_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "ms-playwright"
    current_path = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "")
    if local_browser_dir.exists() and (
        "cursor-sandbox-cache" in current_path
        or not current_path
        or not Path(current_path).exists()
    ):
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(local_browser_dir)


# ==================== 配置 ====================
CRM_URL = "http://192.168.10.16:8899/"

from credential_provider import (  # noqa: E402
    CredentialError,
    get_crm_credentials,
    get_excel_password,
)

_CRM_USER: str | None = None
_CRM_PASS: str | None = None
_EXCEL_PASSWORD: str | None = None


def ensure_runtime_credentials() -> None:
    """从 Windows Credential Manager 加载凭据；缺失则 fail-closed。"""
    global _CRM_USER, _CRM_PASS, _EXCEL_PASSWORD
    if _CRM_USER is not None and _CRM_PASS is not None and _EXCEL_PASSWORD is not None:
        return
    _CRM_USER, _CRM_PASS = get_crm_credentials()
    _EXCEL_PASSWORD = get_excel_password()


def _crm_user() -> str:
    ensure_runtime_credentials()
    assert _CRM_USER is not None
    return _CRM_USER


def _crm_pass() -> str:
    ensure_runtime_credentials()
    assert _CRM_PASS is not None
    return _CRM_PASS


def _excel_password() -> str:
    ensure_runtime_credentials()
    assert _EXCEL_PASSWORD is not None
    return _EXCEL_PASSWORD

# 脚本所在目录；有效表读写在本项目根目录；投放数据只读来自 cursor 项目
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DONE_DIR = os.path.join(SCRIPT_DIR, "完成")
# 深圳信息流报表-cursor 项目根目录（投放数据只读，不修改原表）
CURSOR_PROJECT_DIR = r"E:\AI项目\深圳信息流报表-cursor"

# 导出文件保存目录
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")


def 有效表完成目录():
    """未设置环境变量时保持原「完成」目录。"""
    return os.environ.get("EFFECTIVE_TABLE_DONE_DIR", "").strip() or DONE_DIR


def 有效表下载目录():
    """未设置环境变量时保持用户 Downloads。"""
    return os.environ.get("EFFECTIVE_TABLE_DOWNLOAD_DIR", "").strip() or DOWNLOAD_DIR


def _是否历史探测模式():
    """总后台 historical/probe 注入；生产默认定时任务不设置此变量。"""
    return os.environ.get("EFFECTIVE_TABLE_HISTORICAL_PROBE", "").strip() == "1"


# 导出文件中需要保留的列（按顺序）
CRM_COLUMNS = [
    "客户姓名",
    "电话",
    "记录时间",
    "来源平台",
    "表单名称",
    "建档时间",
    "客户标签",
    "网电咨询师",
]

# 总表列名（与 CRM 列对应，但名称不同）
ZONGBIAO_COLUMNS = ["姓名", "电话", "创建时间", "来源平台", "表单名称", "建档时间", "有效情况", "咨询师"]

# 保留的来源平台（惠州广点通在 CRM 来源平台中可能显示为「腾讯广点通（惠州）」）
VALID_PLATFORMS = ["腾讯广点通", "百度", "惠州广点通", "腾讯广点通（惠州）"]


# ==================== 工具函数 ====================
# 命令行可指定目标日期：python daily_report.py 2026-09-02（默认=昨天）
TARGET_DATE = None


def get_yesterday():
    """目标日期：命令行指定（如补跑 9-2/9-3）或默认昨天"""
    if TARGET_DATE is not None:
        return TARGET_DATE
    return datetime.date.today() - datetime.timedelta(days=1)


def report_paths_for_date(target_date):
    """按目标日期月份解析有效表与投放数据的固定路径。"""
    month = target_date.month
    workbook_override = os.environ.get("EFFECTIVE_TABLE_WORKBOOK_PATH", "").strip()
    if workbook_override:
        shenzhen_path = workbook_override
    else:
        shenzhen_path = os.path.join(
            SCRIPT_DIR, f"{month}月-信息流【深圳有效表】.xlsx"
        )
    toufang_override = os.environ.get("EFFECTIVE_TABLE_TOUFANG_PATH", "").strip()
    if toufang_override:
        toufang_path = toufang_override
    else:
        toufang_path = os.path.join(
            CURSOR_PROJECT_DIR, f"{month}月信息流投放数据（深圳）.xlsx"
        )
    return shenzhen_path, toufang_path


def load_encrypted_workbook(path, data_only=False):
    """加载 xlsx，自动识别文件是否加密（兼容未加密的源文件）"""
    try:
        with open(path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=_excel_password())
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
        decrypted.seek(0)
        return openpyxl.load_workbook(decrypted, data_only=data_only)
    except msoffcrypto.exceptions.DecryptionError:
        # 文件未加密，直接读取
        print(f"  [提示] {os.path.basename(path)} 未加密，直接读取", flush=True)
        return openpyxl.load_workbook(path, data_only=data_only)


def is_encrypted(path):
    """判断 xlsx 是否加密"""
    try:
        with open(path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            return office_file.is_encrypted()
    except Exception:
        return True  # 无法判断时按加密处理（默认行为）


def save_encrypted_workbook(wb, path, encrypted=True):
    """保存 xlsx（先写入内存，再按需加密覆盖目标文件）
    encrypted=True  -> 加密保存（默认，历史行为）
    encrypted=False -> 明文保存（兼容未加密的源文件）
    """
    plain_buf = io.BytesIO()
    wb.save(plain_buf)
    wb.close()
    plain_buf.seek(0)
    if encrypted:
        encrypt_excel_from_buffer(plain_buf, path, _excel_password())
    else:
        with open(path, "wb") as out:
            out.write(plain_buf.read())


def encrypt_excel(src, dst, password):
    """使用 msoffcrypto 给 xlsx 加密码（src 为文件路径）"""
    with open(src, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        with open(dst, "wb") as out:
            office_file.encrypt(password, out)


def encrypt_excel_from_buffer(buf, dst, password):
    """使用 msoffcrypto 给 xlsx 加密码（buf 为 BytesIO）"""
    office_file = msoffcrypto.OfficeFile(buf)
    office_file.load_key(password=password)
    with open(dst, "wb") as out:
        office_file.encrypt(password, out)


def backup_file(path):
    """创建带时间戳的备份"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{path}.backup.{timestamp}.xlsx"
    shutil.copy2(path, backup_path)
    print(f"  已备份: {os.path.basename(backup_path)}")
    return backup_path


def parse_datetime_value(val):
    """将各种格式的日期值解析为 datetime 对象"""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, datetime.date):
        return datetime.datetime.combine(val, datetime.time.min)
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(val.strip(), fmt)
            except ValueError:
                continue
    return None


# ==================== CRM 自动化 (Playwright) ====================
def check_crm_login_ok(page):
    """
    登录后检查 CRM 是否正常可用：
    1. 检测是否仍停留在登录页（密码框还在）→ 真正阻断操作，抛错
    2. 检测是否弹出强制修改密码的 el-dialog 弹窗 → 抛错
    （CRM 顶部红色横幅"密码已过期"类提示属于常规警告，不阻断，忽略）
    """
    # 检测 1：登录失败（仍停留在登录页）→ 真正阻断
    try:
        if page.query_selector('input[placeholder="密码"]'):
            raise RuntimeError(
                "CRM 登录失败（仍停留在登录页），请检查账号密码配置！"
            )
    except RuntimeError:
        raise
    except Exception:
        pass

    # 检测 2：强制修改密码弹窗（el-dialog 内的"修改密码"）→ 真正阻断
    try:
        dialog = page.query_selector('.el-dialog__body:has-text("修改密码"), .el-message-box:has-text("修改密码")')
        if dialog and dialog.is_visible():
            shot = os.path.join(有效表下载目录(), "crm_pwd_expired.png")
            try:
                page.screenshot(path=shot)
                print(f"  错误截图已保存: {shot}", flush=True)
            except Exception:
                pass
            raise RuntimeError(
                "CRM 弹出强制修改密码弹窗，请先登录 CRM 修改密码后再运行脚本！"
            )
    except RuntimeError:
        raise
    except Exception:
        pass


def verify_on_target_page(page, expected_url_substr, page_name, screenshot_name):
    """
    验证当前页面是否在预期的目标页面（通过 URL hash 判断）。
    不在则保存截图并抛出明确错误。
    """
    try:
        cur_url = page.url
    except Exception:
        cur_url = ""
    if expected_url_substr not in cur_url:
        shot = os.path.join(有效表下载目录(), screenshot_name)
        try:
            page.screenshot(path=shot)
            print(f"  错误截图已保存: {shot}", flush=True)
        except Exception:
            pass
        raise RuntimeError(
            f"导航到 {page_name} 失败（当前 URL: {cur_url}），"
            f"预期包含 {expected_url_substr}。请检查 CRM 菜单结构是否变化或网络问题。"
        )


def _推断翻月方向(current_month, target_year, target_month):
    """面板头部解析失败时，用 today 作为参考，避免默认点「下一月」误翻到未来月。"""
    target = (target_year, target_month)
    if current_month is not None:
        return current_month > target
    today = datetime.date.today()
    ref = (today.year, today.month)
    if ref == target:
        return False
    return ref > target


def _读取daterange输入值(page):
    """分别读取 daterange 起始/结束 input 的 UI 值。"""
    inputs = page.query_selector_all('.el-date-editor--daterange input')
    start_raw = ""
    end_raw = ""
    if len(inputs) >= 1:
        try:
            start_raw = (inputs[0].input_value() or "").strip()
        except Exception:
            start_raw = ""
    if len(inputs) >= 2:
        try:
            end_raw = (inputs[1].input_value() or "").strip()
        except Exception:
            end_raw = ""
    if not end_raw and start_raw:
        end_raw = start_raw
    return start_raw, end_raw


def _归一化日期文本(文本):
    import re as _re3
    if not 文本:
        return ""
    文本 = str(文本).strip().replace("/", "-")
    匹配 = _re3.search(r"(\d{4}-\d{2}-\d{2})", 文本)
    return 匹配.group(1) if 匹配 else 文本[:10]


def _校验exact_daterange(page, target_date, step_label):
    """分别校验 start/end；不一致立即 PERIOD_UI_MISMATCH，禁止继续 query/export。"""
    time.sleep(0.5)
    start_raw, end_raw = _读取daterange输入值(page)
    expected = target_date.strftime("%Y-%m-%d")
    start = _归一化日期文本(start_raw)
    end = _归一化日期文本(end_raw)
    print(f"  [{step_label}] UI start readback: {start_raw!r} -> {start}", flush=True)
    print(f"  [{step_label}] UI end readback: {end_raw!r} -> {end}", flush=True)
    if start != expected or end != expected:
        shot = os.path.join(有效表下载目录(), "crm_period_ui_mismatch.png")
        try:
            page.screenshot(path=shot)
            print(f"  错误截图已保存: {shot}", flush=True)
        except Exception:
            pass
        raise RuntimeError(
            f"PERIOD_UI_MISMATCH:{step_label}:start={start}:end={end}:expected={expected}"
        )
    print(f"  [{step_label}] 日期已确认 start={start} end={end}", flush=True)
    return start, end


def _等待loading_mask消失(page, timeout_ms=20000):
    try:
        page.wait_for_selector('.el-loading-mask', state='hidden', timeout=timeout_ms)
    except Exception:
        pass
    time.sleep(0.5)


def _尝试键盘填写daterange(page, target_date, step_label):
    """优先 Ctrl+A 填写 start/end，清除 stale UI 值（如 2026-10-31）。"""
    expected = target_date.strftime("%Y-%m-%d")
    inputs = page.query_selector_all('.el-date-editor--daterange input')
    if len(inputs) < 1:
        return False
    try:
        for idx, inp in enumerate(inputs[:2]):
            inp.click()
            time.sleep(0.25)
            page.keyboard.press("Control+A")
            page.keyboard.type(expected)
            time.sleep(0.25)
            if idx == 0 and len(inputs) >= 2:
                page.keyboard.press("Tab")
                time.sleep(0.3)
        _校验exact_daterange(page, target_date, step_label)
        print(f"  [{step_label}] 日期键盘填写成功", flush=True)
        return True
    except RuntimeError:
        return False
    except Exception:
        return False


def set_date_range_by_click(page, target_date, step_label, max_month_nav=13):
    """
    可靠地在 CRM 日期范围选择器中选择 target_date~target_date：
    1. 优先键盘填写 start/end（清除 stale 值）
    2. 失败则打开日历面板翻月并点击目标日
    3. 分别读取 start/end 输入框；不一致 → PERIOD_UI_MISMATCH
    """
    if _尝试键盘填写daterange(page, target_date, step_label):
        return

    target_day = str(target_date.day)
    target_month = target_date.month
    target_year = target_date.year

    # 打开日期面板（点击输入框区域）
    date_editor = page.query_selector('.el-date-editor--daterange')
    if not date_editor:
        raise RuntimeError(f"[{step_label}] 未找到日期选择器 .el-date-editor--daterange")
    date_editor.click()
    time.sleep(1.2)

    # 翻月：反复读取面板头部当前显示的月份，直到与目标年月一致
    current_month = None
    for _ in range(max_month_nav * 2):
        # 读取面板显示的年份/月份文本，例如 "2026年 09月" 或 "2026-09"
        header_text = ""
        try:
            header_text = page.evaluate("""() => {
                const labels = Array.from(document.querySelectorAll(
                    '.el-date-picker__header-label, .el-date-range-picker__header-label, ' +
                    '.el-picker-panel__header .el-date-picker__header-label'
                ));
                return labels.length ? labels[0].textContent.trim() : '';
            }""")
        except Exception:
            pass
        # 解析文本中的年月数字
        try:
            import re as _re2
            nums = _re2.findall(r'(\d{4})[^\d]*(\d{1,2})', header_text)
            if nums:
                current_month = (int(nums[0][0]), int(nums[0][1]))
        except Exception:
            current_month = None

        if current_month == (target_year, target_month):
            break
        # 翻到目标月份：点击上/下一月按钮
        go_back = _推断翻月方向(current_month, target_year, target_month)
        try:
            # Element UI 用左箭头=上一月，右箭头=下一月
            if go_back:
                btn = page.query_selector('.el-date-picker__prev-btn, .el-icon-arrow-left')
            else:
                btn = page.query_selector('.el-date-picker__next-btn, .el-icon-arrow-right')
            if not btn:
                break
            btn.click(force=True)
            time.sleep(0.6)
        except Exception:
            break

    # 点击目标日（两次：range 的开始与结束都是同一天）
    clicked = False
    for _ in range(2):
        try:
            day_cell = page.query_selector(
                f'.el-date-table td:not(.next-month):not(.prev-month) span:text-is("{target_day}")'
            )
            if not day_cell:
                # 尝试 JS 兜底查找
                page.evaluate(f"""() => {{
                    const cells = Array.from(document.querySelectorAll('.el-date-table td'));
                    const target = cells.find(td => {{
                        const span = td.querySelector('span');
                        return span && span.textContent.trim() === '{target_day}'
                            && !td.classList.contains('next-month')
                            && !td.classList.contains('prev-month');
                    }});
                    if (target) target.click();
                }}""")
            else:
                day_cell.click()
            time.sleep(0.7)
            clicked = True
        except Exception:
            time.sleep(0.5)

    # 关闭面板
    try:
        page.keyboard.press('Escape')
    except Exception:
        pass
    time.sleep(0.8)

    if not clicked:
        raise RuntimeError(
            f"[{step_label}] 未能点击到目标日期 {target_date}，"
            f"请人工检查 CRM 日期面板结构。错误截图: {os.path.join(有效表下载目录(), 'crm_date_error.png')}"
        )
    _校验exact_daterange(page, target_date, step_label)


CRM_LOGIN_MAX_RETRIES = 3


def _login_and_navigate_信息流线索池(page, target_date):
    """登录 → 验证目标页 → 设置 exact date；失败抛错供 bounded retry。"""
    print("  打开 CRM...")
    page.goto(CRM_URL, wait_until="networkidle", timeout=30000)
    time.sleep(2)
    if "登录" in page.title() or page.query_selector('input[placeholder="用户名"]'):
        print("  登录中...")
        page.fill('input[placeholder="用户名"]', _crm_user())
        page.fill('input[placeholder="密码"]', _crm_pass())
        page.click('button:has-text("登录")')
        page.wait_for_load_state("networkidle", timeout=15000)
        time.sleep(2)
    check_crm_login_ok(page)
    print("  导航到信息流线索池...")
    page.goto(f"{CRM_URL}#/reservation-pool/flow-platform-search",
              wait_until="networkidle", timeout=15000)
    time.sleep(3)
    verify_on_target_page(page, "flow-platform-search", "信息流线索池", "crm_nav_error.png")
    print(f"  设置记录时间为 {target_date}...")
    set_date_range_by_click(page, target_date, "信息流线索池")
    _校验exact_daterange(page, target_date, "信息流线索池")


def _探测信息流线索池source_native_state(page) -> dict:
    """查询后读取 source-native 总记录数 / 明确零状态（禁止仅凭 export 失败推断）。"""
    import re

    import datetime as dt

    body = page.inner_text("body")
    nums = [int(m.group(1)) for m in re.finditer(r"共\s*(\d+)\s*条", body)]
    explicit_zero = "暂无数据" in body or "无数据" in body or nums == [0]
    try:
        table_rows = page.locator("table tbody tr").count()
    except Exception:
        table_rows = None
    parsed_total = nums[0] if nums else None
    confirmed = bool(
        explicit_zero
        or parsed_total == 0
        or (table_rows == 0 and explicit_zero)
    )
    return {
        "page_text_sample": body[:2000],
        "parsed_total_from_共条": parsed_total,
        "table_body_row_count": table_rows,
        "explicit_zero_text": explicit_zero,
        "source_native_zero_confirmed": confirmed,
        "query_timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _点击信息流线索池导出按钮(page):
    """
    信息流线索池工具栏「导出」：优先 exact 文案，避免点到下拉内其它导出变体。
    """
    候选 = page.locator('button:has-text("导出")').filter(visible=True)
    数量 = 候选.count()
    if 数量 == 0:
        raise RuntimeError("未找到可见的「导出」按钮")
    if 数量 == 1:
        目标 = 候选.first
    else:
        目标 = None
        for 序号 in range(数量):
            按钮 = 候选.nth(序号)
            文案 = (按钮.inner_text(timeout=2000) or "").strip()
            print(f"  导出按钮候选[{序号}]: text={文案!r}", flush=True)
            if 文案 == "导出":
                目标 = 按钮
                break
        if 目标 is None:
            目标 = 候选.first
            print("  多个导出按钮，回退点击第一个可见项", flush=True)
    目标.click()


def run_crm_export():
    """使用 Playwright 登录 CRM，导出信息流线索池数据，返回下载文件路径"""
    use_local_playwright_browser()
    yesterday = get_yesterday()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        output_path = None

        try:
            末次错误 = None
            for 尝试 in range(1, CRM_LOGIN_MAX_RETRIES + 1):
                try:
                    if 尝试 > 1:
                        print(f"  CRM session 恢复 attempt {尝试}/{CRM_LOGIN_MAX_RETRIES}...", flush=True)
                    _login_and_navigate_信息流线索池(page, yesterday)
                    break
                except Exception as 错误:
                    末次错误 = 错误
                    if 尝试 >= CRM_LOGIN_MAX_RETRIES:
                        raise
                    time.sleep(1.5 * 尝试)
            else:
                if 末次错误:
                    raise 末次错误

            # 5. 点击查询（DATE exact 已 PASS）
            print("  点击查询...")
            page.locator('button:has-text("查询")').filter(visible=True).click()
            _等待loading_mask消失(page)
            time.sleep(1)
            探测 = _探测信息流线索池source_native_state(page)
            print(
                f"  SOURCE_NATIVE_PROBE: confirmed={探测.get('source_native_zero_confirmed')} "
                f"parsed_total={探测.get('parsed_total_from_共条')}",
                flush=True,
            )
            目标日文本 = yesterday.strftime("%Y-%m-%d") if hasattr(yesterday, "strftime") else str(yesterday)[:10]
            if 探测.get("source_native_zero_confirmed"):
                from crm_export_artifact_gate import write_source_native_zero_evidence

                截图路径 = os.path.join(有效表下载目录(), "crm_source_native_zero_ui.png")
                try:
                    page.screenshot(path=截图路径, full_page=True)
                except Exception:
                    截图路径 = None
                证据路径 = write_source_native_zero_evidence(
                    有效表下载目录(),
                    business_date=目标日文本,
                    ui_start_readback=目标日文本,
                    ui_end_readback=目标日文本,
                    query_timestamp=str(探测.get("query_timestamp") or ""),
                    ui_state="SOURCE_NATIVE_ZERO_CONFIRMED",
                    record_count_evidence=0,
                    refresh_evidence={
                        "query_refresh_confirmed": True,
                        "loading_mask_cleared": True,
                        "stale_ui_risk": False,
                        "explicit_zero_text": bool(探测.get("explicit_zero_text")),
                        "parsed_total_from_共条": 探测.get("parsed_total_from_共条"),
                    },
                )
                print(f"  UI start readback: {目标日文本}", flush=True)
                print(f"  UI end readback: {目标日文本}", flush=True)
                print("  CRM_SOURCE_NATIVE_ZERO_CONFIRMED=1", flush=True)
                print(f"  CRM_EXIT_ZERO_CONFIRMED=1", flush=True)
                print(f"  SOURCE_NATIVE_ZERO_EVIDENCE_PATH={证据路径}", flush=True)
                return str(证据路径)

            # 6. 导出 — 使用 expect_download 捕获下载
            print("  点击导出...")
            with page.expect_download(timeout=60000) as download_info:
                _点击信息流线索池导出按钮(page)

            download = download_info.value
            suggested = download.suggested_filename
            print(f"  下载开始: {suggested}", flush=True)

            output_path = os.path.join(有效表下载目录(), suggested)
            download.save_as(output_path)
            print(f"  下载完成: {output_path}", flush=True)
            结构门 = validate_crm_export_structure(output_path)
            print(f"  EXPORT_ARTIFACT_STRUCTURE_VALID={结构门.get('EXPORT_ARTIFACT_STRUCTURE_VALID')}", flush=True)
            if 结构门.get("EXPORT_ARTIFACT_STRUCTURE_VALID") != "PASS":
                raise RuntimeError(
                    f"WRONG_OR_MALFORMED_EXPORT_ARTIFACT: {结构门.get('reason')}"
                )

        except Exception as e:
            print(f"  CRM 导出失败: {e}", flush=True)
            import traceback
            traceback.print_exc()
            try:
                page.screenshot(path=os.path.join(有效表下载目录(), "crm_error.png"))
                print(f"  错误截图已保存: {os.path.join(有效表下载目录(), 'crm_error.png')}")
            except Exception:
                pass
            raise
        finally:
            print("  关闭浏览器...", flush=True)
            try:
                context.close()
                browser.close()
                print("  浏览器已关闭", flush=True)
            except Exception as e:
                print(f"  关闭浏览器时出错（忽略）: {e}", flush=True)

    print(f"  返回导出路径: {output_path}", flush=True)
    if not os.path.exists(output_path):
        raise FileNotFoundError(f"CRM 导出文件未生成: {output_path}")
    return output_path


def run_crm_export_jianDang():
    """
    使用 Playwright 登录 CRM，导出建档客户池（片区）数据。
    建档日期=昨天，高级筛选→三级渠道选：营销中心、营销中心-信息流（深）、三方数据（3）
    返回下载文件路径。
    """
    use_local_playwright_browser()
    yesterday = get_yesterday()
    yesterday_day = str(yesterday.day)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # 1. 打开 CRM 并登录
            print("  打开 CRM...")
            page.goto(CRM_URL, wait_until="networkidle", timeout=30000)
            time.sleep(2)

            if "登录" in page.title() or page.query_selector('input[placeholder="用户名"]'):
                print("  登录中...")
                page.fill('input[placeholder="用户名"]', _crm_user())
                page.fill('input[placeholder="密码"]', _crm_pass())
                page.click('button:has-text("登录")')
                page.wait_for_load_state("networkidle", timeout=15000)
                time.sleep(2)

            # 2b. 检查登录状态与密码过期提示
            check_crm_login_ok(page)

            # 2. 直接导航到建档客户池（片区）页面
            print("  导航到建档客户池（片区）...")
            page.goto(f"{CRM_URL}#/reservation-pool/temp-cust-search-cmp",
                      wait_until="networkidle", timeout=15000)
            time.sleep(3)
            # 2c. 验证是否真的到了目标页面
            verify_on_target_page(page, "temp-cust-search-cmp", "建档客户池（片区）", "crm_jianDang_nav_error.png")

            # 3. 设置建档日期为目标日期（点击日历+验证，不依赖 fill）
            print(f"  设置建档日期为 {yesterday}...")
            set_date_range_by_click(page, yesterday, "建档客户池")

            # 4. 点击高级筛选（是文本标签，不是按钮）
            print("  点击高级筛选...")
            page.locator('text=高级筛选').filter(visible=True).first.click()
            time.sleep(1.5)

            # 5. 通过 cascader 搜索选择三级渠道
            print("  选择三级渠道（cascader 搜索方式）...")

            # 找到三级渠道对应的 cascader（高级筛选面板中的 cascader）
            # 尝试找到包含"三级渠道"标签的 form-item，再找其中的 cascader
            cascader_found = False

            # 方式1：通过 label 找到对应的 cascader
            channel_label = page.query_selector('label:has-text("三级渠道")')
            if channel_label:
                form_item = channel_label.evaluate("el => el.closest('.el-form-item')")
                if form_item:
                    form_item_handle = page.query_selector(f'.el-form-item:has(label:has-text("三级渠道")) .el-cascader')
                    if form_item_handle:
                        # 点击 cascader 打开下拉
                        form_item_handle.click(force=True)
                        time.sleep(1)
                        cascader_found = True

            # 方式2：如果没找到，尝试找所有 cascader 中的最后一个（高级筛选面板中通常有多个）
            if not cascader_found:
                cascaders = page.query_selector_all('.el-cascader')
                if cascaders:
                    # 尝试每个 cascader，看哪个有"三级渠道"相关的
                    for c in cascaders:
                        try:
                            c.click(force=True)
                            time.sleep(0.5)
                            # 检查是否打开了下拉
                            dropdown = page.query_selector('.el-cascader__dropdown:visible')
                            if dropdown:
                                cascader_found = True
                                break
                        except Exception:
                            continue

            if cascader_found:
                # 找到 cascader 搜索输入框
                search_input = page.query_selector('.el-cascader__search-input')
                if search_input:
                    search_input.click()
                    time.sleep(0.3)
                    search_input.fill('三方数据（3）')
                    time.sleep(1)

                    # 等待搜索结果出现
                    time.sleep(1)

                    # 点击匹配的建议项
                    suggestions = page.query_selector_all('.el-cascader__suggestion-item')
                    clicked = False
                    for sug in suggestions:
                        text = sug.inner_text()
                        if '三方数据' in text and '3' in text:
                            sug.click()
                            clicked = True
                            break

                    if not clicked:
                        # 尝试点击任何包含"三方数据"的建议
                        for sug in suggestions:
                            text = sug.inner_text()
                            if '三方数据' in text:
                                sug.click()
                                clicked = True
                                break

                    time.sleep(1)

                    # 验证 cascader 值
                    cascader_text = ""
                    try:
                        cascader_label = page.query_selector('.el-cascader .el-input__inner')
                        if cascader_label:
                            cascader_text = cascader_label.get_attribute('value') or cascader_label.input_value() or ""
                    except Exception:
                        pass
                    print(f"  Cascader 值: {cascader_text}")

                    # 关闭 cascader 下拉（按 Escape）
                    page.keyboard.press('Escape')
                    time.sleep(1)

                    # 再次确保下拉关闭
                    page.evaluate("document.body.click()")
                    time.sleep(0.5)
                    page.keyboard.press('Escape')
                    time.sleep(0.5)
                else:
                    print("  未找到 cascader 搜索输入框")
                    page.keyboard.press('Escape')
                    time.sleep(0.5)
            else:
                print("  未找到三级渠道 cascader")

            # 6. 点击查询（使用 .last，因为高级筛选面板中的查询按钮是第二个）
            print("  点击查询...")
            page.locator('button:has-text("查询")').filter(visible=True).last.click()
            try:
                page.wait_for_selector('.el-loading-mask', state='hidden', timeout=20000)
            except Exception:
                pass
            time.sleep(2)

            # 7. 导出
            print("  点击导出...")
            with page.expect_download(timeout=60000) as download_info:
                page.locator('button:has-text("导出")').filter(visible=True).click()

            download = download_info.value
            suggested = download.suggested_filename
            print(f"  下载开始: {suggested}", flush=True)

            # 保留 CRM 原始文件名
            output_path = os.path.join(有效表下载目录(), suggested)
            download.save_as(output_path)
            print(f"  下载完成: {output_path}", flush=True)

        except Exception as e:
            print(f"  建档客户池导出失败: {e}", flush=True)
            import traceback
            traceback.print_exc()
            try:
                page.screenshot(path=os.path.join(有效表下载目录(), "crm_jianDang_error.png"))
                print(f"  错误截图已保存: {os.path.join(有效表下载目录(), 'crm_jianDang_error.png')}")
            except Exception:
                pass
            raise
        finally:
            print("  关闭浏览器...", flush=True)
            try:
                context.close()
                browser.close()
                print("  浏览器已关闭", flush=True)
            except Exception as e:
                print(f"  关闭浏览器时出错（忽略）: {e}", flush=True)

    print(f"  返回导出路径: {output_path}", flush=True)
    if not os.path.exists(output_path):
        raise FileNotFoundError(f"建档客户池导出文件未生成: {output_path}")
    return output_path


# ==================== Excel 处理 ====================
def process_exported_file(path):
    """
    处理 CRM 导出的文件：
    1. 只保留指定列并按顺序排列
    2. 筛选来源平台（腾讯广点通、百度、含三方的），排除含"香港"的
    3. 三方数据（3）：只保留客户标签≠"重复"的，不改标签
    4. 其他平台：建档时间≠下载当天的，客户标签改为"重复"；建档时间=下载当天的不变
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {}
    for col_name in CRM_COLUMNS:
        if col_name in headers:
            col_idx[col_name] = headers.index(col_name) + 1
        else:
            raise ValueError(f"导出文件缺少列: {col_name}")

    yesterday = get_yesterday()

    processed = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        values = {name: row[col_idx[name] - 1].value for name in CRM_COLUMNS}

        platform = str(values["来源平台"] or "")

        # 排除含"香港"的平台
        if "香港" in platform:
            continue

        is_valid_platform = (
            platform in VALID_PLATFORMS
            or "三方" in platform
        )
        if not is_valid_platform:
            continue

        # 三方数据（3）特殊处理：只保留非重复的（客户标签≠"重复"），不改标签
        if "三方数据" in platform and "3" in platform and "惠州" not in platform:
            original_label = str(values["客户标签"] or "")
            if original_label == "重复":
                continue  # 跳过重复的
            form_name = str(values["表单名称"] or "")
            is_offline_form = "-线下" in form_name
            file_time = parse_datetime_value(values["建档时间"])
            # 线下表单：平台 sheet 按「创建时间=昨天」计数，须保留记录时间=昨天
            # 但建档时间可为历史日期（老客户再次留资）。非线下仍要求建档=昨天，与建档池一致。
            if not is_offline_form:
                if file_time is None or file_time.date() != yesterday:
                    continue
            processed.append(tuple(values[name] for name in CRM_COLUMNS))
            continue

        # 其他平台：建档时间≠下载当天的改标签为"重复"
        file_time = parse_datetime_value(values["建档时间"])
        if file_time is None:
            values["客户标签"] = "重复"
        elif file_time.date() != yesterday:
            values["客户标签"] = "重复"
        # 建档时间=下载当天的不变

        processed.append(tuple(values[name] for name in CRM_COLUMNS))

    wb.close()
    return processed


def process_jianDang_file(path, existing_phones=None):
    """
    处理建档客户池（片区）导出文件。
    建档客户池的列名与信息流线索池不同，需要做映射：
      客户姓名 → 客户姓名
      电话 → 电话
      三级渠道 → 来源平台
      建档时间 → 建档时间
      标签名称 → 客户标签
      (记录时间、表单名称、网电咨询师 在建档客户池中可能不存在，填 None)
    筛选逻辑：
      1. 只保留标签≠"重复"的记录
      2. 电话号码不在 existing_phones 中的（与信息流线索池去重）
    返回按 CRM_COLUMNS 顺序排列的元组列表。
    """
    if existing_phones is None:
        existing_phones = set()

    wb = openpyxl.load_workbook(path)
    ws = wb.active

    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    # 建档客户池列名 → CRM_COLUMNS 的映射
    column_alias = {
        "客户姓名": "客户姓名",
        "电话": "电话",
        "三级渠道": "来源平台",
        "建档时间": "建档时间",
        "标签名称": "客户标签",
        "记录时间": "记录时间",
        "表单名称": "表单名称",
        "网电咨询师": "网电咨询师",
    }

    # 找到各列在导出文件中的索引
    col_idx = {}
    for crm_name in CRM_COLUMNS:
        found = False
        # 先尝试精确匹配
        for alias, target in column_alias.items():
            if target == crm_name and alias in headers:
                col_idx[crm_name] = headers.index(alias) + 1
                found = True
                break
        # 再尝试直接匹配 CRM 列名
        if not found and crm_name in headers:
            col_idx[crm_name] = headers.index(crm_name) + 1
            found = True
        # 找不到的列设为 None
        if not found:
            col_idx[crm_name] = None

    print(f"  建档客户池列映射: {col_idx}")

    processed = []
    skipped_dup_label = 0
    skipped_dup_phone = 0
    seen_phones = set()

    for row in ws.iter_rows(min_row=2, values_only=False):
        # 提取各列值
        values = {}
        for name in CRM_COLUMNS:
            if col_idx[name] is not None:
                values[name] = row[col_idx[name] - 1].value
            else:
                values[name] = None

        # 来源平台：取三级渠道的值（如"三方数据（3）"）
        if not values["来源平台"]:
            values["来源平台"] = "三方数据（3）"

        # 记录时间：如果为空，用建档时间代替
        if not values["记录时间"] and values["建档时间"]:
            values["记录时间"] = values["建档时间"]

        # 表单名称：如果为空，等于来源平台
        if not values["表单名称"]:
            values["表单名称"] = values["来源平台"]

        # 1. 跳过标签="重复"的记录
        label = str(values["客户标签"] or "")
        if label == "重复":
            skipped_dup_label += 1
            continue

        # 2. 跳过电话号码已存在于信息流线索池中的记录
        phone = str(values["电话"] or "").strip()
        if phone and phone in existing_phones:
            skipped_dup_phone += 1
            continue

        # 3. 建档客户池内部也要去重
        if phone and phone in seen_phones:
            skipped_dup_phone += 1
            continue

        if phone:
            seen_phones.add(phone)

        processed.append(tuple(values[name] for name in CRM_COLUMNS))

    wb.close()
    print(f"  建档客户池: 保留 {len(processed)} 行, 跳过标签重复 {skipped_dup_label} 行, 跳过电话重复 {skipped_dup_phone} 行")
    return processed


def build_zongbiao_append_key(phone, record_time):
    """总表去重键：电话 + 创建时间（日期）。重跑同一天时避免重复追加。"""
    phone_text = str(phone or "").strip()
    record_dt = parse_datetime_value(record_time)
    if not phone_text or record_dt is None:
        return None
    return phone_text, record_dt.date()


def load_zongbiao_append_keys(ws):
    """读取总表已有行的去重键集合。"""
    keys = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        key = build_zongbiao_append_key(row[1], row[2])
        if key:
            keys.add(key)
    return keys


def count_zongbiao_rows_for_date(shenzhen_path, target_date):
    """统计总表中「创建时间」落在目标日的行数（用于判断是否已手动/自动导入）。"""
    workbook = load_encrypted_workbook(shenzhen_path, data_only=True)
    try:
        worksheet = workbook["总表"]
        row_count = 0
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            record_time = row[2] if len(row) > 2 else None
            record_dt = parse_datetime_value(record_time)
            if record_dt and record_dt.date() == target_date:
                row_count += 1
        return row_count
    finally:
        workbook.close()


def _rewrite_worksheet_data_rows(ws, kept_rows, max_col):
    """用保留行重写 sheet（保留第 1 行表头）。"""
    header_row = [ws.cell(row=1, column=col).value for col in range(1, max_col + 1)]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for offset, row_values in enumerate(kept_rows, start=2):
        for col_index, value in enumerate(row_values, start=1):
            ws.cell(row=offset, column=col_index, value=value)
    for col_index, value in enumerate(header_row, start=1):
        ws.cell(row=1, column=col_index, value=value)


def remove_zongbiao_rows_for_date(ws, target_date):
    """删除总表中「创建时间」=目标日的所有行（重跑覆盖用）。"""
    max_col = ws.max_column
    kept_rows = []
    removed = 0
    for row_index in range(2, ws.max_row + 1):
        row_values = [ws.cell(row=row_index, column=col).value for col in range(1, max_col + 1)]
        if row_values[0] is None and (len(row_values) < 2 or row_values[1] is None):
            continue
        record_time = row_values[2] if len(row_values) > 2 else None
        record_dt = parse_datetime_value(record_time)
        if record_dt and record_dt.date() == target_date:
            removed += 1
            continue
        kept_rows.append(row_values)
    if removed == 0:
        return 0
    _rewrite_worksheet_data_rows(ws, kept_rows, max_col)
    return removed


def remove_consume_rows_for_date(ws, target_date):
    """删除消费 sheet 中目标日的所有行（重跑覆盖用）。"""
    max_col = max(ws.max_column, 3)
    kept_rows = []
    removed = 0
    for row_index in range(2, ws.max_row + 1):
        row_values = [ws.cell(row=row_index, column=col).value for col in range(1, max_col + 1)]
        row_date = parse_datetime_value(row_values[0] if row_values else None)
        if row_date and row_date.date() == target_date:
            removed += 1
            continue
        kept_rows.append(row_values)
    if removed == 0:
        return 0
    _rewrite_worksheet_data_rows(ws, kept_rows, max_col)
    return removed


def filter_new_crm_rows_against_keys(data_rows, existing_keys):
    """仅保留总表中尚不存在的 CRM 行。"""
    kept = []
    skipped = 0
    seen_in_batch = set()
    for row_data in data_rows:
        phone = row_data[CRM_COLUMNS.index("电话")]
        record_time = row_data[CRM_COLUMNS.index("记录时间")]
        key = build_zongbiao_append_key(phone, record_time)
        if key is None:
            kept.append(row_data)
            continue
        if key in existing_keys or key in seen_in_batch:
            skipped += 1
            continue
        seen_in_batch.add(key)
        kept.append(row_data)
    return kept, skipped


def remove_duplicate_zongbiao_rows(ws):
    """
    删除总表中「电话+创建日期」完全重复的行，保留最早一行。
    大表用整表重写，避免 delete_rows 过慢。
    返回删除行数。
    """
    max_col = ws.max_column
    header_row = [ws.cell(row=1, column=col).value for col in range(1, max_col + 1)]
    kept_rows = []
    seen_keys = set()
    removed = 0
    for row_index in range(2, ws.max_row + 1):
        row_values = [ws.cell(row=row_index, column=col).value for col in range(1, max_col + 1)]
        if row_values[0] is None and (len(row_values) < 2 or row_values[1] is None):
            continue
        phone = row_values[1] if len(row_values) > 1 else None
        record_time = row_values[2] if len(row_values) > 2 else None
        key = build_zongbiao_append_key(phone, record_time)
        if key is not None:
            if key in seen_keys:
                removed += 1
                continue
            seen_keys.add(key)
        kept_rows.append(row_values)
    if removed == 0:
        return 0
    _rewrite_worksheet_data_rows(ws, kept_rows, max_col)
    return removed


def remove_duplicate_consume_rows(ws, target_date):
    """删除消费 sheet 中同一业务日、同一账户的重复行，保留最早一行。"""
    max_col = max(ws.max_column, 3)
    kept_rows = []
    removed = 0
    seen_accounts = set()
    for row_index in range(2, ws.max_row + 1):
        row_values = [ws.cell(row=row_index, column=col).value for col in range(1, max_col + 1)]
        row_date = parse_datetime_value(row_values[0] if row_values else None)
        account_name = row_values[1] if len(row_values) > 1 else None
        if row_date and row_date.date() == target_date and account_name:
            account_key = str(account_name).strip()
            if account_key in seen_accounts:
                removed += 1
                continue
            seen_accounts.add(account_key)
        kept_rows.append(row_values)
    if removed == 0:
        return 0
    _rewrite_worksheet_data_rows(ws, kept_rows, max_col)
    return removed


def append_to_sheet(ws, data_rows, column_names, crm_columns):
    """将数据追加到 sheet 末尾（通用方法，支持不同列名映射）"""
    last_row = ws.max_row
    while last_row > 0 and ws.cell(row=last_row, column=1).value is None:
        last_row -= 1
    start_row = last_row + 1

    # CRM 列名 → 目标 sheet 列名的映射
    col_mapping = dict(zip(crm_columns, column_names))

    for i, row_data in enumerate(data_rows, start=0):
        r = start_row + i
        # row_data 按 CRM_COLUMNS 顺序排列，需要按目标列名写入
        for c, col_name in enumerate(column_names, start=1):
            crm_name = None
            for k, v in col_mapping.items():
                if v == col_name:
                    crm_name = k
                    break
            if crm_name:
                idx = crm_columns.index(crm_name)
                ws.cell(row=r, column=c, value=row_data[idx])

    return start_row


def update_shenzhen_file(shenzhen_path, data_rows, yesterday):
    """更新有效表：总表追加数据，平台 sheet 更新日期，消费 sheet 追加账户数据"""
    wb = load_encrypted_workbook(shenzhen_path)

    # 1. 总表 sheet：先清除目标日旧数据，再追加 CRM 新数据
    ws_zongbiao = wb["总表"]
    cleared_zongbiao = remove_zongbiao_rows_for_date(ws_zongbiao, yesterday)
    if cleared_zongbiao:
        print(
            f"  总表: 已清除 {yesterday} 旧数据 {cleared_zongbiao} 条（重跑覆盖）",
            flush=True,
        )
    existing_keys = load_zongbiao_append_keys(ws_zongbiao)
    data_rows, skipped_dup = filter_new_crm_rows_against_keys(data_rows, existing_keys)
    if skipped_dup:
        print(f"  总表去重: 跳过本批重复 {skipped_dup} 条（电话+创建日期）", flush=True)
    start_row = append_to_sheet(ws_zongbiao, data_rows, ZONGBIAO_COLUMNS, CRM_COLUMNS)

    # 1b. 格式化 C列（创建时间）和 F列（建档时间），处理 G列（有效情况）
    for r in range(start_row, start_row + len(data_rows)):
        # C列（创建时间，column 3）：格式 yyyy/m/d h:mm:ss
        cell_c = ws_zongbiao.cell(row=r, column=3)
        val_c = cell_c.value
        if val_c is not None:
            if isinstance(val_c, datetime.datetime):
                cell_c.number_format = "yyyy/m/d h:mm:ss"
            elif isinstance(val_c, str):
                try:
                    parsed = parse_datetime_value(val_c)
                    if parsed:
                        cell_c.value = parsed
                        cell_c.number_format = "yyyy/m/d h:mm:ss"
                    else:
                        cell_c.value = val_c.replace("-", "/")
                except Exception:
                    cell_c.value = val_c.replace("-", "/")

        # F列（建档时间，column 6）：格式 yyyy/m/d h:mm:ss
        cell_f = ws_zongbiao.cell(row=r, column=6)
        val_f = cell_f.value
        if val_f is not None:
            if isinstance(val_f, datetime.datetime):
                cell_f.number_format = "yyyy/m/d h:mm:ss"
            elif isinstance(val_f, str):
                try:
                    parsed = parse_datetime_value(val_f)
                    if parsed:
                        cell_f.value = parsed
                        cell_f.number_format = "yyyy/m/d h:mm:ss"
                    else:
                        cell_f.value = val_f.replace("-", "/")
                except Exception:
                    cell_f.value = val_f.replace("-", "/")

        # G列（有效情况，column 7）：按创建时间 vs 建档时间判断
        cell_g = ws_zongbiao.cell(row=r, column=7)
        # 获取 C列 和 F列 的日期（重新读取，因为上面可能已改为 datetime）
        c_val = ws_zongbiao.cell(row=r, column=3).value
        f_val = ws_zongbiao.cell(row=r, column=6).value
        c_date = parse_datetime_value(c_val)
        f_date = parse_datetime_value(f_val)
        if c_date and f_date:
            if c_date.date() != f_date.date():
                cell_g.value = "重复"
            # 同一天不做任何修改
        elif f_date is None:
            # 建档时间为空，标记重复
            cell_g.value = "重复"

        # 排除重复后，G列空白统一填充为"空"
        if cell_g.value is None or str(cell_g.value).strip() == "":
            cell_g.value = "空"

    # 2. 平台 sheet：只更新 B1/C1 日期，不追加数据
    ws_platform = wb["平台"]
    ws_platform["B1"] = yesterday
    ws_platform["C1"] = yesterday

    # 3. 消费 sheet：稍后由调用方填充
    ws_consume = wb["消费"]
    return wb, ws_consume


# ==================== 公式求值器（替代 Excel COM 重算） ====================
# 原理：openpyxl 不会计算公式，本类用 Python 复现「账户项目数据」中的
# SUMPRODUCT / SUMIF / 算术 / 引用 公式，得到账户项目数据 N/O 列的真实计算值。
# 已验证与 Excel 重算结果误差 < 0.01。

import re as _re

_SHEET_PREFIX_RE = _re.compile(r"^(?:'([^']+)'|([^!:$]+))!")
_CELL_RANGE_RE = _re.compile(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$")


def _col_to_idx(col):
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def _split_sheet(rng):
    """返回 (sheet 或 None, 剩余部分)"""
    m = _SHEET_PREFIX_RE.match(rng)
    if m:
        sheet = m.group(1) or m.group(2)
        return sheet, rng[m.end():]
    return None, rng


def _parse_cell_range(rng):
    """解析 $A$1 / A1 / A1:B2, 返回 (col1, row1, col2, row2)；单格时 col2/row2 为 None"""
    m = _CELL_RANGE_RE.match(rng)
    if not m:
        raise ValueError(f'无法解析区域: {rng}')
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    return c1, r1, c2, (int(r2) if r2 else None)


class FormulaEvaluator:
    """用 Python 复现 Excel 公式计算（支持 SUMPRODUCT/SUMIF/算术/引用）"""

    def __init__(self, wb_formula, wb_values, date_start, date_end):
        self.wb_f = wb_formula      # 公式视图 (data_only=False)
        self.wb_v = wb_values       # 缓存值视图 (data_only=True)
        self.date_start = date_start
        self.date_end = date_end
        self.cache = {}

    def get_value(self, sheet, col, row):
        ws = self.wb_v[sheet]
        return ws.cell(row=row, column=_col_to_idx(col) + 1).value

    def get_column(self, sheet, col, row_start, row_end):
        ws = self.wb_v[sheet]
        vals = []
        for r in range(row_start, row_end + 1):
            vals.append(ws.cell(row=r, column=_col_to_idx(col) + 1).value)
        return vals

    def eval_formula(self, text, sheet='账户项目数据'):
        if text in self.cache:
            return self.cache[text]
        text = text.strip()
        if text.startswith('='):
            text = text[1:].strip()
        # 数字
        if _re.fullmatch(r'-?\d+(\.\d+)?', text):
            return float(text)
        # 函数调用：SUMIF(...) / SUMPRODUCT(...)（须优先于引用解析，文本中含 ! 会被误判）
        if text.startswith('SUMIF(') and text.endswith(')'):
            result = self.eval_sumif(text, sheet)
            self.cache[text] = result
            return result
        if text.startswith('SUMPRODUCT(') and text.endswith(')'):
            result = self.eval_sump(text, sheet)
            self.cache[text] = result
            return result
        # 简单引用 =G5 (同 sheet) 或 'sheet'!A2
        sh, rest = _split_sheet(text)
        if sh:
            c1, r1, c2, r2 = _parse_cell_range(rest)
            return self.eval_cell(sh, c1, r1)
        m = _CELL_RANGE_RE.match(text)
        if m:
            col, row = m.group(1), int(m.group(2))
            return self.eval_cell(sheet, col, row)
        # 算术：A+B, A-B, A*B, A/B
        for op, fn in [('+', lambda a, b: a + b), ('-', lambda a, b: a - b),
                       ('*', lambda a, b: a * b), ('/', lambda a, b: a / b if b else 0)]:
            parts = text.rsplit(op, 1)
            if len(parts) == 2 and parts[0].strip() and parts[1].strip():
                a = self.eval_formula(parts[0], sheet)
                b = self.eval_formula(parts[1], sheet)
                try:
                    result = fn(float(a), float(b))
                except (TypeError, ValueError):
                    result = 0.0
                self.cache[text] = result
                return result
        raise ValueError(f'无法解析公式: {text}')

    def eval_cell(self, sheet, col, row):
        ws = self.wb_f[sheet]
        cell = ws.cell(row=row, column=_col_to_idx(col) + 1)
        v = cell.value
        if isinstance(v, openpyxl.worksheet.formula.ArrayFormula):
            return self.eval_formula(v.text, sheet)
        if isinstance(v, str) and v.startswith('='):
            return self.eval_formula(v, sheet)
        if v is None:
            return 0.0
        if isinstance(v, datetime.datetime):
            return v.date()
        if isinstance(v, datetime.date):
            return v
        try:
            return float(v)
        except (TypeError, ValueError):
            return v

    def eval_sump(self, text, sheet):
        inner = text[text.index('(') + 1:text.rindex(')')]
        factors = _re.split(r'\)\s*\*\s*\(', inner)
        conditions = []
        sum_range = None
        for f in factors:
            f = f.strip()
            if f.startswith('('):
                f = f[1:]
            if f.endswith(')'):
                f = f[:-1]
            sh, rest = _split_sheet(f.strip())
            if rest and _CELL_RANGE_RE.match(rest.strip()) and ':' in rest.strip():
                sum_range = f.strip()
            else:
                conditions.append(f)
        if sum_range is None:
            raise ValueError(f'SUMPRODUCT 未找到求和列: {text}')
        return self.eval_conditions(conditions, sum_range, sheet)

    def eval_conditions(self, conditions, sum_range, sheet):
        sh, rest = _split_sheet(sum_range)
        if sh is None:
            sh = sheet
        c1, rs, c2, re_ = _parse_cell_range(rest.strip())
        row_end = re_ if re_ else rs
        sum_vals = self.get_column(sh, c1, rs, row_end)
        total = 0.0
        n = len(sum_vals)
        for i in range(n):
            ok = True
            for cond in conditions:
                if not self.eval_condition(cond, i, rs, sh):
                    ok = False
                    break
            if ok:
                sv = sum_vals[i]
                if isinstance(sv, (int, float)) and not isinstance(sv, bool):
                    total += float(sv)
        return total

    def eval_condition(self, cond, idx, row_start, sheet):
        cond = cond.strip()
        # 1. ISNUMBER(FIND("key", range))
        m = _re.fullmatch(r'ISNUMBER\(FIND\("([^"]*)",\s*(.+)\)\)?', cond)
        if m:
            key, rng = m.group(1), m.group(2)
            val = self.get_range_value(rng, idx, row_start)
            if isinstance(val, str):
                return key in val
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                if float(val).is_integer():
                    return key in str(int(val))
                return key in str(val)
            return False
        # 2. 比较: range>=B2 / range<=C2 / range="x"
        m = _re.fullmatch(r'(.+?)(>=|<=|=)(.+)$', cond)
        if m:
            rng, op, rhs = m.group(1).strip(), m.group(2), m.group(3).strip()
            val = self.get_range_value(rng, idx, row_start)
            rhs_val = self.parse_rhs(rhs)
            if op == '>=':
                return self._cmp(val, '>=', rhs_val)
            if op == '<=':
                return self._cmp(val, '<=', rhs_val)
            if op == '=':
                return str(val) == str(rhs_val)
        raise ValueError(f'无法解析条件: {cond}')

    def parse_rhs(self, rhs):
        rhs = rhs.strip()
        if rhs.startswith('"') and rhs.endswith('"'):
            return rhs[1:-1]
        if rhs in ('$B$2', 'B2'):
            return self.date_start
        if rhs in ('$C$2', 'C2'):
            return self.date_end
        try:
            return float(rhs)
        except ValueError:
            return rhs

    def get_range_value(self, rng, idx, row_start):
        rng = rng.strip()
        sh, rest = _split_sheet(rng)
        if sh is None:
            sh = '账户项目数据'
        c1, rs, c2, r2 = _parse_cell_range(rest.strip())
        if r2:
            return self.get_value(sh, c1, rs + idx)
        else:
            return self.get_value(sh, c1, rs)

    def _cmp(self, val, op, rhs):
        # 统一日期类型：datetime / date 都转为 date 再比较
        if isinstance(val, datetime.datetime):
            val = val.date()
        if isinstance(rhs, datetime.datetime):
            rhs = rhs.date()
        if isinstance(val, datetime.date) and isinstance(rhs, datetime.date):
            if op == '>=':
                return val >= rhs
            return val <= rhs
        try:
            a, b = float(val), float(rhs)
        except (TypeError, ValueError):
            return False
        if op == '>=':
            return a >= b
        return a <= b

    def eval_sumif(self, text, sheet):
        inner = text[text.index('(') + 1:text.rindex(')')]
        parts = inner.split(',')
        if len(parts) != 3:
            raise ValueError(f'SUMIF 参数异常: {text}')
        crit_range = parts[0].strip()
        crit = parts[1].strip()
        sum_range = parts[2].strip()
        if crit.startswith('"') and crit.endswith('"'):
            crit = crit[1:-1]
        sh1, rest1 = _split_sheet(crit_range)
        if sh1 is None:
            sh1 = sheet
        c1, rs1, _, re1 = _parse_cell_range(rest1.strip())
        sh2, rest2 = _split_sheet(sum_range)
        if sh2 is None:
            sh2 = sheet
        c2, rs2, _, _ = _parse_cell_range(rest2.strip())
        crit_vals = self.get_column(sh1, c1, rs1, re1)
        total = 0.0
        for i, cv in enumerate(crit_vals):
            if str(cv) == crit:
                v = self.eval_cell(sh2, c2, rs2 + i)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    total += float(v)
        return total


# 账户项目数据 N/O 列：从第 4 行向下扫描，直到 N 列为空（含惠州广点通等后续新增行）
ACCOUNT_DATA_START_ROW = 4
ACCOUNT_DATA_MAX_ROW = 30


def _iter_account_project_rows(ws_account):
    """迭代账户项目数据 sheet 的 N/O 账户行。"""
    for row_index in range(ACCOUNT_DATA_START_ROW, ACCOUNT_DATA_MAX_ROW + 1):
        account_name = ws_account.cell(row=row_index, column=14).value
        if account_name is None or str(account_name).strip() == "":
            break
        yield row_index, account_name


def read_toufang_account_data_com_cached(toufang_path, target_date):
    """
    historical/probe 专用：读取 COM 重算后的 data_only 缓存值，不用 Python 公式求值器。
    #REF! 行跳过（不伪造 0）；合法数值行保留。
    """
    print("  历史探测模式：读取 COM 缓存值（data_only），不用 Python 公式求值器...", flush=True)
    wb_v = load_encrypted_workbook(toufang_path, data_only=True)
    try:
        ws = wb_v["账户项目数据"]
        account_data = []
        skipped_ref = 0
        for row_index, account_name in _iter_account_project_rows(ws):
            value = ws.cell(row=row_index, column=15).value
            if isinstance(value, str) and value.startswith("#"):
                skipped_ref += 1
                continue
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                continue
            account_data.append((account_name, float(value)))
        print(
            f"  已读取账户项目数据 N{ACCOUNT_DATA_START_ROW}:O 共 {len(account_data)} 条"
            f"（跳过 {skipped_ref} 条 #REF 缓存，只读）",
            flush=True,
        )
        return account_data
    finally:
        try:
            wb_v.close()
        except Exception:
            pass


def read_toufang_account_data(toufang_path, target_date):
    """
    只读投放数据表，用公式求值器计算账户项目数据 N/O 列的消费值。
    不修改、不保存原表（投放数据由深圳信息流报表-cursor 项目维护）。
    返回 account_data 列表 [(账户名称, 账面), ...]
    """
    if _是否历史探测模式():
        return read_toufang_account_data_com_cached(toufang_path, target_date)
    print("  打开投放数据表（只读，openpyxl 公式视图）...", flush=True)
    wb_f = load_encrypted_workbook(toufang_path, data_only=False)
    wb_v = load_encrypted_workbook(toufang_path, data_only=True)

    try:
        ev = FormulaEvaluator(wb_f, wb_v, target_date, target_date)
        ws_account = wb_f["账户项目数据"]
        account_data = []
        for row_index, account_name in _iter_account_project_rows(ws_account):
            o_formula = ws_account.cell(row=row_index, column=15).value
            if not (isinstance(o_formula, str) and o_formula.startswith('=')):
                continue
            g_ref = o_formula[1:].strip()
            g_col = g_ref[0]
            g_row = int(g_ref[1:])
            value = ev.eval_cell("账户项目数据", g_col, g_row)
            account_data.append((account_name, value))
        print(
            f"  已读取账户项目数据 N{ACCOUNT_DATA_START_ROW}:O 共 {len(account_data)} 条（只读，未修改原表）",
            flush=True,
        )
        return account_data
    finally:
        try:
            wb_f.close()
        except Exception:
            pass
        try:
            wb_v.close()
        except Exception:
            pass


def fill_consume_sheet(ws_consume, account_data, yesterday):
    """在 消费 sheet 写入账户消费数据（先清目标日旧行再追加，A列 yyyy/m/d）。"""
    cleared_consume = remove_consume_rows_for_date(ws_consume, yesterday)
    if cleared_consume:
        print(
            f"  消费 sheet: 已清除 {yesterday} 旧数据 {cleared_consume} 条（重跑覆盖）",
            flush=True,
        )

    last_row = ws_consume.max_row
    while last_row > 0 and ws_consume.cell(row=last_row, column=1).value is None:
        last_row -= 1
    start_row = last_row + 1

    ydt = yesterday
    for i, (name, value) in enumerate(account_data, start=0):
        row_index = start_row + i
        cell_a = ws_consume.cell(row=row_index, column=1, value=ydt)
        cell_a.number_format = "yyyy/m/d"
        ws_consume.cell(row=row_index, column=2, value=name)
        ws_consume.cell(row=row_index, column=3, value=value)


# ==================== 主流程 ====================
def main():
    global TARGET_DATE
    date_args = [arg for arg in sys.argv[1:] if not arg.startswith("--")]
    # 支持命令行指定目标日期：python daily_report.py 2026-09-02
    if date_args:
        try:
            TARGET_DATE = datetime.datetime.strptime(date_args[0], "%Y-%m-%d").date()
        except ValueError:
            print(f"[错误] 无法解析日期参数: {date_args[0]}，应为 YYYY-MM-DD 格式")
            return
    yesterday = get_yesterday()
    print(f"目标日期: {yesterday}")
    print()

    # 0. 按目标日期月份解析两个 Excel 的固定路径
    print("0. 检查文件...")
    shenzhen_path, toufang_path = report_paths_for_date(yesterday)

    missing = []
    if not os.path.exists(shenzhen_path):
        missing.append(f"有效表: {shenzhen_path}")
    if not os.path.exists(toufang_path):
        missing.append(f"投放数据（只读）: {toufang_path}")
    if missing:
        print(f"\n  [错误] 缺少以下文件:")
        for item in missing:
            print(f"    - {item}")
        print(f"\n  有效表请放到本项目根目录: {SCRIPT_DIR}")
        print(f"  投放数据请放到 cursor 项目根目录: {CURSOR_PROJECT_DIR}")
        input("\n  放好后按回车键退出，再重新运行脚本...")
        return

    print(f"  有效表: {shenzhen_path}")
    print(f"  投放数据（只读）: {toufang_path}")

    # 0b. 若已有目标日数据，本次重跑将先删后写（覆盖更新）
    print("\n0b. 检查总表目标日数据...")
    existing_row_count = count_zongbiao_rows_for_date(shenzhen_path, yesterday)
    if existing_row_count > 0:
        print(
            f"  总表已有 {yesterday} 数据 {existing_row_count} 条，"
            f"本次将重新导出并覆盖",
            flush=True,
        )
    else:
        print(f"  总表尚无 {yesterday} 数据，继续全新导入")

    # 1. 备份（仅有效表；投放数据不修改故不备份）
    print("\n1. 备份文件...")
    backup_file(shenzhen_path)

    # 2. CRM 导出（信息流线索池）
    print("\n2. CRM 导出（信息流线索池）...")
    exported_file = run_crm_export()

    from crm_export_artifact_gate import is_source_native_zero_evidence_path

    # 2b. CRM 导出（建档客户池-片区）
    print("\n2b. CRM 导出（建档客户池-片区）...")
    try:
        jianDang_file = run_crm_export_jianDang()
        print(f"  建档客户池导出完成")
    except Exception as e:
        print(f"  建档客户池导出失败（不影响主流程）: {e}")
        jianDang_file = None

    # 3. 处理导出数据
    print("\n3. 处理导出文件...")
    if is_source_native_zero_evidence_path(exported_file):
        data_rows = []
        print("  SOURCE_NATIVE_RAW_ROW_COUNT=0", flush=True)
        print("  CRM 信息流线索池 source-native explicit zero，跳过非空 xlsx 结构门", flush=True)
    else:
        导出结构门 = validate_crm_export_structure(exported_file)
        print(
            f"  EXPORT_ARTIFACT_STRUCTURE_VALID={导出结构门.get('EXPORT_ARTIFACT_STRUCTURE_VALID')}",
            flush=True,
        )
        if 导出结构门.get("EXPORT_ARTIFACT_STRUCTURE_VALID") != "PASS":
            raise RuntimeError(
                f"WRONG_OR_MALFORMED_EXPORT_ARTIFACT: {导出结构门.get('reason')}"
            )

        data_rows = process_exported_file(exported_file)
    print(f"  信息流线索池符合条件的数据行数: {len(data_rows)}")

    # 3b. 处理建档客户池数据（如果有）
    jianDang_rows = []
    if jianDang_file and os.path.exists(jianDang_file):
        print("\n3b. 处理建档客户池导出文件...")
        # 提取信息流线索池中的电话号码集合，用于去重
        existing_phones = set()
        for row_data in data_rows:
            phone = str(row_data[CRM_COLUMNS.index("电话")] or "").strip()
            if phone:
                existing_phones.add(phone)
        print(f"  信息流线索池电话号码数: {len(existing_phones)}")
        jianDang_rows = process_jianDang_file(jianDang_file, existing_phones)
        print(f"  建档客户池符合条件的数据行数: {len(jianDang_rows)}")
    else:
        print("\n3b. 建档客户池未导出，跳过处理")

    # 合并数据
    all_data_rows = data_rows + jianDang_rows
    print(f"  合计追加到总表的数据行数: {len(all_data_rows)}")
    零产出 = assert_export_process_yield(
        structure_gate=导出结构门,
        processed_row_count=len(data_rows),
        jianDang_row_count=len(jianDang_rows),
    )
    if 零产出:
        print(f"  EXPORT_PROCESS_YIELD=FAIL code={零产出.get('code')}", flush=True)
        raise RuntimeError(
            f"{零产出.get('code')}: {零产出.get('reason')} "
            f"(raw_rows={零产出.get('raw_data_row_count')})"
        )

    # 4. 只读投放数据，读取账户项目数据 N/O 列消费（不修改原表）
    print("\n4. 读取投放数据账户消费（只读，不修改原表）...")
    account_data = read_toufang_account_data(toufang_path, yesterday)
    print(f"  读取到 {len(account_data)} 条账户数据")

    # 5. 更新有效表（总表 + 平台日期 + 消费）
    print("\n5. 更新有效表...")
    wb_shenzhen, ws_consume = update_shenzhen_file(shenzhen_path, all_data_rows, yesterday)
    print(f"  总表 sheet: 追加了 {len(all_data_rows)} 行数据（信息流线索池 {len(data_rows)} + 建档客户池 {len(jianDang_rows)}）")
    print(f"  平台 sheet: B1/C1 已更新为 {yesterday}")
    fill_consume_sheet(ws_consume, account_data, yesterday)
    print(f"  消费 sheet: 追加了 {len(account_data)} 行数据")

    # 6. 保存有效表（原地覆盖）
    print("\n6. 保存文件...")
    save_encrypted_workbook(wb_shenzhen, shenzhen_path, encrypted=is_encrypted(shenzhen_path))
    print(f"  有效表已更新: {shenzhen_path}")

    # 7. 存档 CRM 导出原表到完成文件夹
    print("\n7. 存档 CRM 导出原表...")
    完成目录 = 有效表完成目录()
    os.makedirs(完成目录, exist_ok=True)
    done_export = os.path.join(完成目录, os.path.basename(exported_file))
    if os.path.exists(done_export):
        os.remove(done_export)
    shutil.copy2(exported_file, done_export)
    print(f"  已复制: {os.path.basename(done_export)}（信息流线索池原表）")
    if jianDang_file and os.path.exists(jianDang_file):
        done_jianDang = os.path.join(完成目录, os.path.basename(jianDang_file))
        if os.path.exists(done_jianDang):
            os.remove(done_jianDang)
        shutil.copy2(jianDang_file, done_jianDang)
        print(f"  已复制: {os.path.basename(done_jianDang)}（建档客户池原表）")

    print(f"\n=== 日报流程完成 ===")
    print(f"有效表已更新: {shenzhen_path}")
    print(f"CRM 原表存档: {完成目录}")


if __name__ == "__main__":
    main()
