# -*- coding: utf-8 -*-
"""
修复有效表总表/消费 sheet 重复行（重跑导致电话+创建日期重复）。
用法：python fix_sanfang_dedup.py
"""
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import daily_report as dr


def main():
    shenzhen_path, _ = dr.report_paths_for_date(dr.get_yesterday())
    if not os.path.exists(shenzhen_path):
        # 默认修当前月有效表
        import datetime
        shenzhen_path = os.path.join(
            SCRIPT_DIR,
            f"{datetime.date.today().month}月-信息流【深圳有效表】.xlsx",
        )
    print(f"修复文件: {shenzhen_path}")
    dr.backup_file(shenzhen_path)
    wb = dr.load_encrypted_workbook(shenzhen_path)
    ws_zongbiao = wb["总表"]
    ws_consume = wb["消费"]
    removed_zongbiao = dr.remove_duplicate_zongbiao_rows(ws_zongbiao)
    print(f"总表删除重复行: {removed_zongbiao}")

    removed_consume_total = 0
    seen_dates = set()
    for row in ws_consume.iter_rows(min_row=2, values_only=True):
        row_date = dr.parse_datetime_value(row[0] if row else None)
        if row_date:
            seen_dates.add(row_date.date())
    for business_date in sorted(seen_dates):
        removed = dr.remove_duplicate_consume_rows(ws_consume, business_date)
        removed_consume_total += removed
    print(f"消费 sheet 删除重复行: {removed_consume_total}")

    dr.save_encrypted_workbook(
        wb, shenzhen_path, encrypted=dr.is_encrypted(shenzhen_path)
    )
    print(f"已保存: {shenzhen_path}")


if __name__ == "__main__":
    main()
