#!/usr/bin/env python3
"""
标注团队绩效日报生成器
用法：双击运行，或 python3 生成日报.py
自动比较最新两份历史全量文件，输出每个标注员的产量/质量波动日报。
"""

import os
import re
import sys
from datetime import datetime, timedelta
import pandas as pd

# ==================== 配置 ====================
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATTERN = re.compile(r"团队绩效明细表_历史全量_(\d{4})_(\d{1,2})_(\d{1,2})\.xls")
WARN_THRESHOLD = 10.0  # 波动率超过 10% 触发警告

# ==================== 辅助函数 ====================

def parse_date_from_filename(filename: str):
    """从文件名解析日期，返回 (date_obj, date_str) 或 None"""
    m = FILE_PATTERN.search(filename)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return (datetime(y, mo, d).date(), f"{y}_{mo}_{d}")


def extract_pct(val) -> float | None:
    """从 '94.4% (对85/阅90)' 格式中提取百分比数字"""
    if pd.isna(val) or str(val).strip() == "-":
        return None
    s = str(val).strip()
    m = re.match(r"([\d.]+)%", s)
    if m:
        return float(m.group(1))
    return None


def read_file(filepath: str) -> pd.DataFrame:
    """读取一个 .xls (HTML格式) 绩效文件，返回清洗后的 DataFrame"""
    tables = pd.read_html(filepath)
    df = tables[0].copy()
    # 过滤掉 "未分配" 行
    df = df[df["标注员"] != "未分配"].copy()
    df["标注员"] = df["标注员"].astype(str).str.strip()
    # 提取百分比数值列
    df["质检通过率_pct"] = df["🛡️质检通过率"].apply(extract_pct)
    df["验收首次通过率_pct"] = df["🎯验收(首次验收通过)"].apply(extract_pct)
    df["验收累积通过率_pct"] = df["🎯验收(累积折损通过)"].apply(extract_pct)
    # 确保已标注是数值
    df["已标注"] = pd.to_numeric(df["已标注"], errors="coerce").fillna(0).astype(int)
    return df


def calc_fluctuation(new_val, old_val):
    """计算波动率 (%)。old_val 为 0 时返回 None"""
    if old_val is None or new_val is None:
        return None
    if old_val == 0:
        return None  # 无法计算波动率
    return round((new_val - old_val) / old_val * 100, 2)


def flag_warn(rate):
    """正向波动超过阈值→优秀，负向波动超过阈值→警告，其余→空"""
    if rate is None:
        return ""
    if rate > WARN_THRESHOLD:
        return "🌟 优秀"
    if rate < -WARN_THRESHOLD:
        return "⚠️ 异常"
    return ""


# ==================== 主流程 ====================

def main():
    # 1. 扫描目录下所有符合条件的文件
    files_info = []
    for fname in os.listdir(DATA_DIR):
        result = parse_date_from_filename(fname)
        if result:
            date_obj, date_str = result
            files_info.append((date_obj, date_str, os.path.join(DATA_DIR, fname)))

    if len(files_info) < 2:
        print(f"[错误] 至少需要 2 个历史文件才能做对比，当前找到 {len(files_info)} 个。")
        sys.exit(1)

    # 按日期排序，取最新两个
    files_info.sort(key=lambda x: x[0])
    old_info = files_info[-2]
    new_info = files_info[-1]

    old_date_obj, old_date_str, old_path = old_info
    new_date_obj, new_date_str, new_path = new_info

    print(f"📊 对比：{new_date_str}（最新） vs {old_date_str}（上期）")
    print()

    # 2. 读取两个文件
    df_old = read_file(old_path)
    df_new = read_file(new_path)

    # 3. 按标注员合并
    merged = df_old[["标注员", "已标注", "质检通过率_pct", "验收首次通过率_pct", "验收累积通过率_pct"]].copy()
    merged.columns = ["标注员", "已标注_old", "质检通过率_old", "验收首次通过率_old", "验收累积通过率_old"]

    df_new_sel = df_new[["标注员", "已标注", "质检通过率_pct", "验收首次通过率_pct", "验收累积通过率_pct"]].copy()
    df_new_sel.columns = ["标注员", "已标注_new", "质检通过率_new", "验收首次通过率_new", "验收累积通过率_new"]

    merged = merged.merge(df_new_sel, on="标注员", how="outer")

    # 4. 计算各项波动
    for metric, label in [
        ("已标注", "产量(已标注)"),
        ("质检通过率", "质检通过率"),
        ("验收首次通过率", "验收首次通过率"),
        ("验收累积通过率", "验收累积通过率"),
    ]:
        old_col = f"{metric}_old"
        new_col = f"{metric}_new"
        diff_col = f"{metric}_差值"
        rate_col = f"{metric}_波动率"
        warn_col = f"{metric}_警告"

        if metric == "已标注":
            merged[old_col] = merged[old_col].fillna(0).astype(int)
            merged[new_col] = merged[new_col].fillna(0).astype(int)
            merged[diff_col] = merged[new_col] - merged[old_col]
            merged[rate_col] = merged.apply(
                lambda r: calc_fluctuation(r[new_col], r[old_col]), axis=1
            )
        else:
            merged[old_col] = merged[old_col]
            merged[new_col] = merged[new_col]
            merged[diff_col] = merged.apply(
                lambda r: (
                    round(r[new_col] - r[old_col], 1)
                    if r[new_col] is not None and r[old_col] is not None
                    else None
                ),
                axis=1,
            )
            merged[rate_col] = merged.apply(
                lambda r: calc_fluctuation(r[new_col], r[old_col]), axis=1
            )

        merged[warn_col] = merged[rate_col].apply(flag_warn)

    # 5. 输出日报

    # -- 文本日报 --
    lines = []
    lines.append("=" * 70)
    lines.append(f"  标注团队绩效日报")
    lines.append(f"  最新数据：{new_date_str.replace('_', '/')}    上期数据：{old_date_str.replace('_', '/')}")
    lines.append(f"  生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("=" * 70)
    lines.append("")

    has_any_warn = False
    has_any_excellent = False

    # 按标注员遍历
    for _, row in merged.iterrows():
        name = row["标注员"]
        # 处理新增/离职标注员
        if pd.isna(row["已标注_old"]):
            lines.append(f"  [{name}] 🆕 新成员，上期无数据")
            lines.append(f"    本期产量: {int(row['已标注_new'])} 条")
            if row["质检通过率_new"] is not None:
                lines.append(f"    质检通过率: {row['质检通过率_new']}%")
            if row["验收首次通过率_new"] is not None:
                lines.append(f"    验收首次通过率: {row['验收首次通过率_new']}%")
            lines.append("")
            continue
        if pd.isna(row["已标注_new"]):
            lines.append(f"  [{name}] ⚪ 上期存在但本期无数据（可能已离职）")
            lines.append("")
            continue

        lines.append(f"  ┌─ [{name}] ─────────────────────────────")

        # 产量
        old_v = int(row["已标注_old"])
        new_v = int(row["已标注_new"])
        diff = int(row["已标注_差值"])
        rate = row["已标注_波动率"]
        flag = row["已标注_警告"]
        if flag == "⚠️ 异常":
            has_any_warn = True
        elif flag == "🌟 优秀":
            has_any_excellent = True
        dir_sign = "+" if diff >= 0 else ""
        rate_str = f"{rate:+.1f}%" if rate is not None else "N/A"
        lines.append(f"  │ 📦 产量(已标注): {old_v} → {new_v}  ({dir_sign}{diff} 条, 波动 {rate_str})  {flag}")

        # 质检通过率
        old_p = row["质检通过率_old"]
        new_p = row["质检通过率_new"]
        diff_p = row["质检通过率_差值"]
        rate_p = row["质检通过率_波动率"]
        flag_p = row["质检通过率_警告"]
        if flag_p == "⚠️ 异常":
            has_any_warn = True
        elif flag_p == "🌟 优秀":
            has_any_excellent = True
        if old_p is not None and new_p is not None:
            dir_p = "+" if diff_p >= 0 else ""
            rate_p_str = f"{rate_p:+.1f}%" if rate_p is not None else "N/A"
            lines.append(f"  │ 🛡️ 质检通过率:    {old_p}% → {new_p}%  ({dir_p}{diff_p}pp, 波动 {rate_p_str})  {flag_p}")
        else:
            lines.append(f"  │ 🛡️ 质检通过率:    数据缺失")

        # 验收首次通过率
        old_a1 = row["验收首次通过率_old"]
        new_a1 = row["验收首次通过率_new"]
        diff_a1 = row["验收首次通过率_差值"]
        rate_a1 = row["验收首次通过率_波动率"]
        flag_a1 = row["验收首次通过率_警告"]
        if flag_a1 == "⚠️ 异常":
            has_any_warn = True
        elif flag_a1 == "🌟 优秀":
            has_any_excellent = True
        if old_a1 is not None and new_a1 is not None:
            dir_a1 = "+" if diff_a1 >= 0 else ""
            rate_a1_str = f"{rate_a1:+.1f}%" if rate_a1 is not None else "N/A"
            lines.append(f"  │ 🎯 首次验收通过率: {old_a1}% → {new_a1}%  ({dir_a1}{diff_a1}pp, 波动 {rate_a1_str})  {flag_a1}")
        else:
            lines.append(f"  │ 🎯 首次验收通过率: 数据缺失")

        # 验收累积通过率
        old_a2 = row["验收累积通过率_old"]
        new_a2 = row["验收累积通过率_new"]
        diff_a2 = row["验收累积通过率_差值"]
        rate_a2 = row["验收累积通过率_波动率"]
        flag_a2 = row["验收累积通过率_警告"]
        if flag_a2 == "⚠️ 异常":
            has_any_warn = True
        elif flag_a2 == "🌟 优秀":
            has_any_excellent = True
        if old_a2 is not None and new_a2 is not None:
            dir_a2 = "+" if diff_a2 >= 0 else ""
            rate_a2_str = f"{rate_a2:+.1f}%" if rate_a2 is not None else "N/A"
            lines.append(f"  │ 🎯 累积验收通过率: {old_a2}% → {new_a2}%  ({dir_a2}{diff_a2}pp, 波动 {rate_a2_str})  {flag_a2}")
        else:
            lines.append(f"  │ 🎯 累积验收通过率: 数据缺失")

        lines.append(f"  └──────────────────────────────────────────")
        lines.append("")

    lines.append("=" * 70)
    if has_any_warn:
        lines.append("⚠️  存在负向波动超 10% 的指标，请关注！")
    if has_any_excellent:
        lines.append("🌟 存在正向波动超 10% 的优秀指标，请继续保持！")
    if not has_any_warn and not has_any_excellent:
        lines.append("✅ 所有指标波动率均在 10% 以内。")
    lines.append("=" * 70)

    report_text = "\n".join(lines)
    print(report_text)

    # -- 保存文本日报 --
    txt_path = os.path.join(DATA_DIR, f"日报_{new_date_str}.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(report_text)
    print(f"\n📄 文本日报已保存: {txt_path}")

    # -- 保存 Excel 日报 --
    xlsx_path = os.path.join(DATA_DIR, f"日报_{new_date_str}.xlsx")
    excel_cols = [
        "标注员",
        "已标注_old", "已标注_new", "已标注_差值", "已标注_波动率", "已标注_警告",
        "质检通过率_old", "质检通过率_new", "质检通过率_差值", "质检通过率_波动率", "质检通过率_警告",
        "验收首次通过率_old", "验收首次通过率_new", "验收首次通过率_差值", "验收首次通过率_波动率", "验收首次通过率_警告",
        "验收累积通过率_old", "验收累积通过率_new", "验收累积通过率_差值", "验收累积通过率_波动率", "验收累积通过率_警告",
    ]
    # 重命名列头为中文
    rename_map = {
        "标注员": "标注员",
        "已标注_old": f"已标注({old_date_str})",
        "已标注_new": f"已标注({new_date_str})",
        "已标注_差值": "产量差值",
        "已标注_波动率": "产量波动率(%)",
        "已标注_警告": "产量警告",
        "质检通过率_old": f"质检通过率({old_date_str})",
        "质检通过率_new": f"质检通过率({new_date_str})",
        "质检通过率_差值": "质检差值(pp)",
        "质检通过率_波动率": "质检波动率(%)",
        "质检通过率_警告": "质检警告",
        "验收首次通过率_old": f"首次验收通过率({old_date_str})",
        "验收首次通过率_new": f"首次验收通过率({new_date_str})",
        "验收首次通过率_差值": "首次验收差值(pp)",
        "验收首次通过率_波动率": "首次验收波动率(%)",
        "验收首次通过率_警告": "首次验收警告",
        "验收累积通过率_old": f"累积验收通过率({old_date_str})",
        "验收累积通过率_new": f"累积验收通过率({new_date_str})",
        "验收累积通过率_差值": "累积验收差值(pp)",
        "验收累积通过率_波动率": "累积验收波动率(%)",
        "验收累积通过率_警告": "累积验收警告",
    }
    out_df = merged[excel_cols].rename(columns=rename_map)

    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        out_df.to_excel(writer, sheet_name="日报", index=False)
        ws = writer.sheets["日报"]
        # 条件格式：红色警告 / 金色优秀 / 绿色正常
        warn_fmt = writer.book.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        excel_fmt = writer.book.add_format({"bg_color": "#FFD700", "font_color": "#7B5800"})
        ok_fmt = writer.book.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        for col_idx, col_name in enumerate(out_df.columns):
            if "警告" in col_name:
                for row_idx in range(1, len(out_df) + 1):
                    val = out_df.iloc[row_idx - 1][col_name]
                    if val == "⚠️ 异常":
                        ws.write(row_idx, col_idx, val, warn_fmt)
                    elif val == "🌟 优秀":
                        ws.write(row_idx, col_idx, val, excel_fmt)
                    else:
                        ws.write(row_idx, col_idx, val, ok_fmt)
        ws.autofit()

    print(f"📊 Excel 日报已保存: {xlsx_path}")

    # 末尾再输出一个简短汇总
    print()
    wc = sum(1 for c in out_df.columns if "警告" in c and (out_df[c] == "⚠️ 异常").any())
    ec = sum(1 for c in out_df.columns if "警告" in c and (out_df[c] == "🌟 优秀").any())
    if wc > 0 and ec > 0:
        print(f"⚠️  共 {wc} 项指标出现负向异常  |  🌟 共 {ec} 项指标表现优秀")
    elif wc > 0:
        print(f"⚠️  共 {wc} 项指标出现负向异常，请关注！")
    elif ec > 0:
        print(f"🌟 共 {ec} 项指标表现优秀！")
    else:
        print("✅ 所有指标波动率均在 10% 以内，团队表现稳定。")

    return merged


if __name__ == "__main__":
    main()
