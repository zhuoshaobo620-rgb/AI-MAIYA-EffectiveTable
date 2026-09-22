# -*- coding: utf-8 -*-
"""
补写总表中缺失的「三方数据（3）-线下」行（历史建档、昨日留资）。
用法：python fix_offline_missing.py [YYYY-MM-DD]
默认补目标日=2026-09-13（完成/ 中须有对应 CRM 导出）。
"""
import datetime
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import daily_report as dr


def main():
    if len(sys.argv) > 1:
        target = datetime.datetime.strptime(sys.argv[1], "%Y-%m-%d").date()
    else:
        target = datetime.date(2026, 9, 13)
    dr.TARGET_DATE = target

    leads_path = os.path.join(SCRIPT_DIR, "完成", "信息流查询.xlsx")
    if not os.path.exists(leads_path):
        print(f"[错误] 找不到 CRM 导出: {leads_path}")
        sys.exit(1)

    shenzhen_path, _ = dr.report_paths_for_date(target)
    print(f"目标日: {target}")
    print(f"有效表: {shenzhen_path}")

    all_rows = dr.process_exported_file(leads_path)
    offline_rows = [
        row
        for row in all_rows
        if "-线下" in str(row[dr.CRM_COLUMNS.index("表单名称")] or "")
    ]
    print(f"修复逻辑下线下行数: {len(offline_rows)}")

    dr.backup_file(shenzhen_path)
    workbook = dr.load_encrypted_workbook(shenzhen_path)
    worksheet = workbook["总表"]
    existing_keys = dr.load_zongbiao_append_keys(worksheet)
    new_rows, skipped = dr.filter_new_crm_rows_against_keys(offline_rows, existing_keys)
    print(f"需补写: {len(new_rows)} 条，已存在跳过: {skipped} 条")

    if not new_rows:
        print("无需补写")
        workbook.close()
        return

    start_row = dr.append_to_sheet(
        worksheet, new_rows, dr.ZONGBIAO_COLUMNS, dr.CRM_COLUMNS
    )
    for row_index in range(start_row, start_row + len(new_rows)):
        cell_create = worksheet.cell(row=row_index, column=3)
        cell_file = worksheet.cell(row=row_index, column=6)
        create_dt = dr.parse_datetime_value(cell_create.value)
        file_dt = dr.parse_datetime_value(cell_file.value)
        if create_dt and file_dt and create_dt.date() != file_dt.date():
            worksheet.cell(row=row_index, column=7, value="重复")
        elif worksheet.cell(row=row_index, column=7).value in (None, ""):
            worksheet.cell(row=row_index, column=7, value="空")

    platform_sheet = workbook["平台"]
    platform_sheet["B1"] = target
    platform_sheet["C1"] = target

    dr.save_encrypted_workbook(
        workbook, shenzhen_path, encrypted=dr.is_encrypted(shenzhen_path)
    )
    print(f"已补写并保存: {shenzhen_path}")


if __name__ == "__main__":
    main()
