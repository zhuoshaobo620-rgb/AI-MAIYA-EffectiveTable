# -*- coding: utf-8 -*-
"""
补跑脚本：补做 2026-09-02 和 2026-09-03 两天的数据。
流程：
  1. 项目根目录需有当月有效表（如 9月-信息流【深圳有效表】.xlsx）
  2. cursor 项目根目录需有当月投放数据（只读，不修改）
  3. 跑 9-2 → 有效表原地覆盖 → 9-3 基于 9-2 结果继续跑
最终有效表在项目根目录，含 9-1+9-2+9-3 数据。
"""
import os
import sys
import shutil
import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
DONE = os.path.join(BASE, "完成")

sys.path.insert(0, BASE)
import daily_report as dr

SZ_FILE = "9月-信息流【深圳有效表】.xlsx"


def step_run(date_str, label):
    """设置目标日期并执行一次 main"""
    print("\n" + "=" * 60)
    print(f"  补跑 {label} ({date_str})")
    print("=" * 60)
    dr.TARGET_DATE = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    dr.main()


if __name__ == "__main__":
    target = datetime.date(2026, 9, 2)
    shenzhen_path, toufang_path = dr.report_paths_for_date(target)
    if not os.path.exists(shenzhen_path):
        print(f"[错误] 缺少有效表: {shenzhen_path}")
        sys.exit(1)
    if not os.path.exists(toufang_path):
        print(f"[错误] 缺少投放数据（只读）: {toufang_path}")
        sys.exit(1)
    print(f"基础有效表: {os.path.basename(shenzhen_path)}")
    print(f"投放数据（只读）: {toufang_path}")

    # 备份完成文件夹里可能存在的旧 CRM 存档（避免误覆盖后无法追溯）
    for fname in os.listdir(DONE) if os.path.isdir(DONE) else []:
        if fname.endswith(".xlsx") and "查询" in fname:
            src = os.path.join(DONE, fname)
            bak = src + ".polluted.bak"
            if os.path.exists(src):
                shutil.copy2(src, bak)
                print(f"  已备份: {os.path.basename(bak)}")

    step_run("2026-09-02", "9月2日")
    step_run("2026-09-03", "9月3日")

    print("\n" + "=" * 60)
    print(f"  补跑完成！有效表已更新: {os.path.join(BASE, SZ_FILE)}")
    print("=" * 60)
