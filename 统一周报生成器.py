#!/usr/bin/env python3
"""
统一项目周报生成器
==================
独立生成周报 TXT + 图片，用法：

  python3 统一周报生成器.py                        # 生成上周一→本周日的周报
  python3 统一周报生成器.py 2026-6-17              # 以指定日期为截止，自动前推 7 天
  python3 统一周报生成器.py 2026-6-17 2026-6-10   # 手动指定新旧日期

会生成：
  - 统一周报_*.txt
  - 项目A_DATA_周报_*.png  /  项目B_ANOMALY_周报_*.png
"""

import os
import sys

# 把本脚本所在目录加到路径，以便导入统一日报生成器中的函数
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from datetime import datetime, timedelta

# 从统一日报生成器导入所需函数
from 统一日报生成器 import (
    _parse_date_arg, _collect_report_data,
    format_full_report, format_report_image,
    OUTPUT_DIR,
)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]

    if len(args) >= 1:
        target = _parse_date_arg(args[0])
    else:
        now = datetime.now()
        target = f"{now.year}_{now.month}_{now.day}"

    if len(args) >= 2:
        old_d = _parse_date_arg(args[1])
    else:
        parts = target.split("_")
        wt = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
        wo = wt - timedelta(days=7)
        old_d = f"{wo.year}_{wo.month}_{wo.day}"

    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print("═" * 60)
    print("  统一项目周报生成器")
    print(f"  时间：{run_time}")
    print(f"  窗口：{old_d.replace('_','/')} → {target.replace('_','/')}（7 天）")
    print("═" * 60)
    print()

    # 收集数据（7 天窗口）
    all_team, all_dash, all_analysis = _collect_report_data(target, old_d)

    # 生成周报 TXT
    print("📝 生成文本周报...")
    txt_content = format_full_report(all_team, all_dash, all_analysis, run_time, "周报")
    txt_path = os.path.join(OUTPUT_DIR, f"统一周报_{run_time.replace(':', '-').replace(' ', '_')}.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(txt_content)
    print(f"   TXT 周报已保存: {txt_path}")

    # 生成周报图片
    print("🖼️  生成周报图片...")
    format_report_image(all_team, all_dash, all_analysis, run_time, "周报")

    print()
    print("═" * 60)
    print("  周报生成完毕。")
    print("═" * 60)


if __name__ == "__main__":
    main()
