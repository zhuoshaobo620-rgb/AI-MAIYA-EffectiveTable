# -*- coding: utf-8 -*-
"""
补写消费 sheet 中缺失的「惠州惠城麦芽口腔医院有限公司」（惠州广点通）消费行。
原因：投放表新增账户后惠州广点通移至 N14，旧逻辑只读 N4:N13。
用法：python fix_huizhou_consume.py [YYYY-MM-DD ...]
不传日期时，自动扫描 9 月有效表并补全所有缺失该账户的日期。
"""
import datetime
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import daily_report as dr

HUIZHOU_GDT_ACCOUNT = "惠州惠城麦芽口腔医院有限公司"


def _消费日期缺少惠州账户(ws_consume, target_date):
    """目标日有消费数据但缺少惠州广点通账户时返回 True。"""
    has_any = False
    has_huizhou = False
    for row_index in range(2, ws_consume.max_row + 1):
        cell_date = dr.parse_datetime_value(ws_consume.cell(row_index, 1).value)
        if not cell_date or cell_date.date() != target_date:
            continue
        has_any = True
        account_name = str(ws_consume.cell(row_index, 2).value or "")
        if HUIZHOU_GDT_ACCOUNT in account_name:
            has_huizhou = True
            break
    return has_any and not has_huizhou


def _扫描九月缺失日期(shenzhen_path):
    workbook = dr.load_encrypted_workbook(shenzhen_path, data_only=True)
    try:
        ws_consume = workbook["消费"]
        missing_dates = []
        seen_dates = set()
        for row_index in range(2, ws_consume.max_row + 1):
            cell_date = dr.parse_datetime_value(ws_consume.cell(row_index, 1).value)
            if not cell_date or cell_date.month != 9:
                continue
            seen_dates.add(cell_date.date())
        for target_date in sorted(seen_dates):
            if _消费日期缺少惠州账户(ws_consume, target_date):
                missing_dates.append(target_date)
        return missing_dates
    finally:
        workbook.close()


def _重写目标日消费(shenzhen_path, toufang_path, target_date):
    dr.TARGET_DATE = target_date
    account_data = dr.read_toufang_account_data(toufang_path, target_date)
    huizhou_rows = [
        (name, value) for name, value in account_data if HUIZHOU_GDT_ACCOUNT in str(name)
    ]
    if not huizhou_rows:
        print(f"  [跳过] {target_date}：投放表未找到 {HUIZHOU_GDT_ACCOUNT}")
        return False

    workbook = dr.load_encrypted_workbook(shenzhen_path)
    try:
        ws_consume = workbook["消费"]
        dr.fill_consume_sheet(ws_consume, account_data, target_date)
        dr.save_encrypted_workbook(
            workbook, shenzhen_path, encrypted=dr.is_encrypted(shenzhen_path)
        )
        print(
            f"  [完成] {target_date}：已重写消费 {len(account_data)} 条，"
            f"惠州广点通={huizhou_rows[0][1]:.2f}"
        )
        return True
    finally:
        try:
            workbook.close()
        except Exception:
            pass


def main():
    if len(sys.argv) > 1:
        target_dates = [
            datetime.datetime.strptime(arg, "%Y-%m-%d").date()
            for arg in sys.argv[1:]
        ]
    else:
        sample_date = datetime.date(2026, 9, 19)
        shenzhen_path, _ = dr.report_paths_for_date(sample_date)
        if not os.path.exists(shenzhen_path):
            print(f"[错误] 找不到有效表: {shenzhen_path}")
            sys.exit(1)
        target_dates = _扫描九月缺失日期(shenzhen_path)
        if not target_dates:
            print("未发现缺失惠州广点通消费的日期")
            return

    print(f"待补写日期: {', '.join(str(d) for d in target_dates)}")
    shenzhen_path, toufang_path = dr.report_paths_for_date(target_dates[0])
    dr.backup_file(shenzhen_path)
    print(f"有效表: {shenzhen_path}")
    print(f"投放表: {toufang_path}")

    fixed_count = 0
    for target_date in target_dates:
        if _重写目标日消费(shenzhen_path, toufang_path, target_date):
            fixed_count += 1

    print(f"共补写 {fixed_count}/{len(target_dates)} 个日期")


if __name__ == "__main__":
    main()
