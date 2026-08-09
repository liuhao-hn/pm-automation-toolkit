#!/usr/bin/env python3
"""
统一项目日报生成器 —— 跨项目、跨供应商自动对比 + 数据分析 + 邮件发送
========================================================================
⚠️ 本文件为脱敏版本（Portfolio Version），项目名称、供应商信息、人员数据已做脱敏处理。
   原始代码已在实际生产环境中稳定运行数月，本版本保留全部工程逻辑供展示。

覆盖项目：项目A_DATA（供应商Alpha/供应商Beta/供应商Gamma）、项目B_ANOMALY（供应商Delta）
覆盖维度：团队绩效（标注员级） + 项目大盘（团队级）
数据模式：全量日报 + 当天日报（各取最新两份对比）

输出：
  1. TXT 文本日报（含数据分析意见）
  2. Excel 报表（多 Sheet，条件格式）
  3. 邮件发送至指定邮箱（可选）

用法：
  python3 统一日报生成器.py                  # 生成 + 发送邮件
  python3 统一日报生成器.py --no-email        # 仅生成，不发邮件
  python3 统一日报生成器.py --dry-run         # 仅打印，不保存文件
  python3 统一日报生成器.py --date 2026-6-15  # 手动指定"新"日期，自动找上期对比
  python3 统一日报生成器.py --date 2026-6-15 --old-date 2026-6-11  # 同时指定新旧日期
"""

import os
import re
import sys
import json
import smtplib
import ssl
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from collections import defaultdict

import math
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import pandas as pd


def vis_width(s: str) -> int:
    """计算字符串的可视宽度（ASCII 占 1 列，其余字符占 2 列）"""
    w = 0
    for c in s:
        w += 1 if ord(c) < 128 else 2
    return w


def vis_ljust(s: str, width: int) -> str:
    """按可视宽度左对齐"""
    return s + " " * (width - vis_width(s))


def vis_center(s: str, width: int) -> str:
    """按可视宽度居中"""
    pad = width - vis_width(s)
    left = pad // 2
    return " " * left + s + " " * (pad - left)

# ============================================================================
# 配置
# ============================================================================

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(ROOT_DIR, "统一日报输出")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 项目 → 供应商 映射
PROJECTS = {
    "项目A_DATA": ["供应商Alpha", "供应商Beta", "供应商Gamma"],
    "项目B_ANOMALY": ["供应商Delta", "供应商Epsilon", "供应商Zeta", "供应商Eta"],
}

# 供应商名 → 目录名（不一致时使用）
SUPPLIER_DIR_MAP = {}

# 阈值
WARN_THRESHOLD = 10.0          # 波动率 ±10% 触发标记
QUALIFY_THRESHOLD_QC = 90.0    # 质检通过率 >= 90% 为合格
QUALIFY_THRESHOLD_AC = 80.0    # 验收通过率 >= 80% 为合格
EXCELLENT_RATE = 95.0          # 通过率 >= 95% 为优秀（绝对值）

# 离职人员配置：标注员 → 最后有效日期（该日期及之后的数据将被排除）
EXCLUDED_ANNOTATORS = {
    "user_a001": datetime(2026, 5, 25).date(),
    "user_b002": datetime(2020, 1, 1).date(),
    "user_c003": datetime(2026, 6, 10).date(),
    "user_d004": datetime(2026, 6, 10).date(),
    "user_e005": datetime(2026, 6, 10).date(),
    "user_f006": datetime(2026, 6, 10).date(),
    "user_g007": datetime(2026, 6, 10).date(),
    "user_h008": datetime(2026, 6, 10).date(),
}

# 邮件配置
EMAIL_CONFIG = {
    "smtp_host": "smtp.example.com",
    "smtp_port": 465,
    "sender": "your-email@example.com",
    "password": "",          # 优先级最低：可直接填密码（不推荐提交到 git）
    "to": ["your-email@example.com"],
}

# 本地配置文件路径（优先级最高，推荐使用）
_EMAIL_CONFIG_FILE = os.path.join(ROOT_DIR, "email_config.json")

def _load_email_password() -> str:
    """按优先级加载邮箱密码：1. 本地配置文件  2. 环境变量  3. 脚本硬编码"""
    # 1. 本地配置文件
    if os.path.exists(_EMAIL_CONFIG_FILE):
        try:
            with open(_EMAIL_CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            pwd = cfg.get("password", "").strip()
            if pwd:
                return pwd
        except Exception:
            pass

    # 2. 环境变量
    pwd = os.environ.get("EMAIL_PASSWORD", "").strip()
    if pwd:
        return pwd

    # 3. 脚本硬编码
    return EMAIL_CONFIG.get("password", "").strip()

# ============================================================================
# 辅助函数
# ============================================================================

def parse_date_from_filename(filename: str, base_name: str):
    """从文件名解析日期和类型，返回 (ftype, date_obj, date_str) 或 None"""
    pattern = re.compile(
        re.escape(base_name) + r"_(历史全量|当天实时)_(\d{4})_(\d{1,2})_(\d{1,2})\.xls"
    )
    m = pattern.search(filename)
    if not m:
        return None
    ftype = m.group(1)
    y, mo, d = int(m.group(2)), int(m.group(3)), int(m.group(4))
    return (ftype, datetime(y, mo, d).date(), f"{y}_{mo}_{d}")


def extract_pct(val):
    """从百分比字符串提取数值"""
    if pd.isna(val) or str(val).strip() == "-":
        return None
    m = re.match(r"([\d.]+)%", str(val).strip())
    return float(m.group(1)) if m else None


def parse_team_quality_cell(val):
    """从 '94.4% (对85/阅90)' 等格式提取 (通过数, 总量)"""
    if pd.isna(val) or str(val).strip() == "-":
        return (0, 0)
    s = str(val).strip()
    # 尝试多种格式: "94.4% (对85/阅90)", "94.4% (85/90)", "94.4%(85/90)"
    m = re.search(r"\([^\d]*(\d+)\s*/[^\d]*(\d+)\)", s)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return (0, 0)


def parse_quality_cell(val):
    """从 '90.9% (10/11)' 提取 (通过数, 总量)"""
    if pd.isna(val) or str(val).strip() == "-":
        return (0, 0)
    m = re.match(r"[\d.]+%\s*\((\d+)/(\d+)\)", str(val).strip())
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return (0, 0)


def calc_rate(pass_cnt, total_cnt):
    if total_cnt == 0:
        return None
    return round(pass_cnt / total_cnt * 100, 1)


def _metric_to_pass_col(metric: str) -> str:
    """团队绩效指标名 → pass列名"""
    m = {"质检通过率": "质检_pass", "验收首次通过率": "验收首次_pass", "验收累积通过率": "验收累积_pass"}
    return m.get(metric, "")


def _metric_to_total_col(metric: str) -> str:
    """团队绩效指标名 → total列名"""
    m = {"质检通过率": "质检_total", "验收首次通过率": "验收首次_total", "验收累积通过率": "验收累积_total"}
    return m.get(metric, "")


def calc_fluctuation(new_val, old_val):
    if old_val is None or new_val is None or old_val == 0:
        return None
    return round((new_val - old_val) / old_val * 100, 2)


def flag_warn(rate):
    """波动标记（仅用于产量等无合格阈值的指标）"""
    if not _valid(rate):
        return ""
    if rate > WARN_THRESHOLD:
        return "🌟 优秀"
    if rate < -WARN_THRESHOLD:
        return "⚠️ 异常"
    return ""


def _valid(v) -> bool:
    """值有效（非 None 且非 NaN）"""
    if v is None:
        return False
    if isinstance(v, float) and math.isnan(v):
        return False
    return True


def flag_combined(fluct_rate, pass_rate, threshold):
    """
    综合标记（用于质量指标）：
      负向波动>10% → 异常（最高优先级）
      正向波动>10% 且 通过率>=阈值 → 优秀
      通过率>=阈值 但波动不显著 → 合格
      通过率<阈值 → 不合格
      无数据（None/NaN） → 空，不评价
    质检阈值 90%，验收阈值 80%
    """
    if _valid(fluct_rate) and fluct_rate < -WARN_THRESHOLD:
        return "⚠️ 异常"
    if _valid(fluct_rate) and fluct_rate > WARN_THRESHOLD and _valid(pass_rate) and pass_rate >= threshold:
        return "🌟 优秀"
    if _valid(pass_rate) and pass_rate >= threshold:
        return "✅ 合格"
    if _valid(pass_rate) and pass_rate < threshold:
        return "❌ 不合格"
    return ""


# ============================================================================
# 团队绩效处理器（标注员维度）
# ============================================================================

def process_team_performance(team_dir: str, target_date: str = "", old_date: str = "") -> dict:
    """
    扫描团队目录，生成全量和当天两份对比报告。
    target_date: "YYYY_M_D" 格式，指定"新"日期；为空则自动取最新
    old_date:    同格式，指定"旧"日期；为空则自动取 target_date 之前的最新
    返回 {"全量": report_dict, "当天": report_dict} 或 {}
    """
    base_name = "标注团队绩效明细表"
    files_by_type = defaultdict(list)

    for fname in os.listdir(team_dir):
        result = parse_date_from_filename(fname, base_name)
        if result:
            ftype, date_obj, date_str = result
            files_by_type[ftype].append((date_obj, date_str, os.path.join(team_dir, fname)))

    results = {}
    for ftype in ["历史全量", "当天实时"]:
        flist = sorted(files_by_type.get(ftype, []), key=lambda x: x[0])
        if len(flist) < 2:
            continue
        if target_date:
            # 查找匹配 target_date 的文件作为"新"文件
            new_info = None
            for info in flist:
                if info[1] == target_date:
                    new_info = info
                    break
            if new_info is None:
                print(f"   ⚠️ 未找到日期 {target_date} 的 {ftype} 文件，跳过")
                continue
            # 确定"旧"文件：优先用 old_date，否则找 target_date 之前最新的
            old_info = None
            if old_date:
                for info in flist:
                    if info[1] == old_date:
                        old_info = info
                        break
                if old_info is None:
                    print(f"   ⚠️ 未找到旧日期 {old_date} 的 {ftype} 文件，跳过")
                    continue
            else:
                for info in reversed(flist):
                    if info[0] < new_info[0]:
                        old_info = info
                        break
                if old_info is None:
                    print(f"   ⚠️ {target_date} 之前无可用 {ftype} 文件作为对比基准，跳过")
                    continue
            results[ftype] = _generate_team_comparison(old_info, new_info, base_name)
        else:
            old_info, new_info = flist[-2], flist[-1]
            results[ftype] = _generate_team_comparison(old_info, new_info, base_name)

    return results


def _generate_team_comparison(old_info, new_info, base_name):
    """生成单份团队绩效对比"""
    old_date_obj, old_date_str, old_path = old_info
    new_date_obj, new_date_str, new_path = new_info

    def read_team_file(filepath):
        tables = pd.read_html(filepath)
        df = tables[0].copy()
        df = df[df["标注员"] != "未分配"].copy()
        df["标注员"] = df["标注员"].astype(str).str.strip()
        df["质检通过率_pct"] = df["🛡️质检通过率"].apply(extract_pct)
        df["验收首次通过率_pct"] = df["🎯验收(首次验收通过)"].apply(extract_pct)
        df["验收累积通过率_pct"] = df["🎯验收(累积折损通过)"].apply(extract_pct)
        # 提取 通过数 / 总量
        qc_parsed = df["🛡️质检通过率"].apply(parse_team_quality_cell)
        ac1_parsed = df["🎯验收(首次验收通过)"].apply(parse_team_quality_cell)
        ac2_parsed = df["🎯验收(累积折损通过)"].apply(parse_team_quality_cell)
        df["质检_pass"] = qc_parsed.apply(lambda x: x[0])
        df["质检_total"] = qc_parsed.apply(lambda x: x[1])
        df["验收首次_pass"] = ac1_parsed.apply(lambda x: x[0])
        df["验收首次_total"] = ac1_parsed.apply(lambda x: x[1])
        df["验收累积_pass"] = ac2_parsed.apply(lambda x: x[0])
        df["验收累积_total"] = ac2_parsed.apply(lambda x: x[1])
        df["已标注"] = pd.to_numeric(df["已标注"], errors="coerce").fillna(0).astype(int)
        return df

    df_old = read_team_file(old_path)
    df_new = read_team_file(new_path)

    # 过滤已离职标注员
    if EXCLUDED_ANNOTATORS:
        for name, cutoff in EXCLUDED_ANNOTATORS.items():
            if new_date_obj >= cutoff:
                df_new = df_new[df_new["标注员"] != name]
            if old_date_obj >= cutoff:
                df_old = df_old[df_old["标注员"] != name]

    old_cols = ["标注员", "已标注", "质检通过率_pct", "验收首次通过率_pct", "验收累积通过率_pct",
                "质检_pass", "质检_total", "验收首次_pass", "验收首次_total", "验收累积_pass", "验收累积_total"]
    merged = df_old[old_cols].copy()
    merged.columns = ["标注员", "已标注_old", "质检通过率_old", "验收首次通过率_old", "验收累积通过率_old",
                      "质检_pass_old", "质检_total_old", "验收首次_pass_old", "验收首次_total_old", "验收累积_pass_old", "验收累积_total_old"]

    new_cols = ["标注员", "已标注", "质检通过率_pct", "验收首次通过率_pct", "验收累积通过率_pct",
                "质检_pass", "质检_total", "验收首次_pass", "验收首次_total", "验收累积_pass", "验收累积_total"]
    df_new_sel = df_new[new_cols].copy()
    df_new_sel.columns = ["标注员", "已标注_new", "质检通过率_new", "验收首次通过率_new", "验收累积通过率_new",
                          "质检_pass_new", "质检_total_new", "验收首次_pass_new", "验收首次_total_new", "验收累积_pass_new", "验收累积_total_new"]

    merged = merged.merge(df_new_sel, on="标注员", how="outer")

    for metric in ["已标注", "质检通过率", "验收首次通过率", "验收累积通过率"]:
        old_col = f"{metric}_old"
        new_col = f"{metric}_new"
        diff_col = f"{metric}_差值"
        rate_col = f"{metric}_波动率"
        warn_col = f"{metric}_标记"

        if metric == "已标注":
            merged[old_col] = merged[old_col].fillna(0).astype(int)
            merged[new_col] = merged[new_col].fillna(0).astype(int)
            merged[diff_col] = merged[new_col] - merged[old_col]
        else:
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
        if metric == "已标注":
            merged[warn_col] = merged[rate_col].apply(flag_warn)
        else:
            thr = QUALIFY_THRESHOLD_QC if "质检" in metric else QUALIFY_THRESHOLD_AC
            merged[warn_col] = merged.apply(
                lambda r: flag_combined(r[rate_col], r[new_col], thr), axis=1
            )

    return {
        "old_date": old_date_str,
        "new_date": new_date_str,
        "data": merged,
    }


# ============================================================================
# 项目大盘处理器（团队维度）
# ============================================================================

def process_project_dashboard(dashboard_dir: str, target_date: str = "", old_date: str = "") -> dict:
    """
    扫描整体目录，生成全量和当天两份对比报告。
    target_date: "YYYY_M_D" 格式，指定"新"日期；为空则自动取最新
    old_date:    同格式，指定"旧"日期；为空则自动取 target_date 之前的最新
    返回 {"全量": report_dict, "当天": report_dict} 或 {}
    """
    base_name = "项目大盘进度与质量统计表"
    files_by_type = defaultdict(list)

    for fname in os.listdir(dashboard_dir):
        result = parse_date_from_filename(fname, base_name)
        if result:
            ftype, date_obj, date_str = result
            files_by_type[ftype].append((date_obj, date_str, os.path.join(dashboard_dir, fname)))

    results = {}
    for ftype in ["历史全量", "当天实时"]:
        flist = sorted(files_by_type.get(ftype, []), key=lambda x: x[0])
        if len(flist) < 2:
            continue
        if target_date:
            new_info = None
            for info in flist:
                if info[1] == target_date:
                    new_info = info
                    break
            if new_info is None:
                print(f"   ⚠️ 未找到日期 {target_date} 的 {ftype} 文件，跳过")
                continue
            old_info = None
            if old_date:
                for info in flist:
                    if info[1] == old_date:
                        old_info = info
                        break
                if old_info is None:
                    print(f"   ⚠️ 未找到旧日期 {old_date} 的 {ftype} 文件，跳过")
                    continue
            else:
                for info in reversed(flist):
                    if info[0] < new_info[0]:
                        old_info = info
                        break
                if old_info is None:
                    print(f"   ⚠️ {target_date} 之前无可用 {ftype} 文件作为对比基准，跳过")
                    continue
            results[ftype] = _generate_dashboard_comparison(old_info, new_info, base_name)
        else:
            old_info, new_info = flist[-2], flist[-1]
            results[ftype] = _generate_dashboard_comparison(old_info, new_info, base_name)

    return results


def _generate_dashboard_comparison(old_info, new_info, base_name):
    """生成单份项目大盘对比"""
    old_date_obj, old_date_str, old_path = old_info
    new_date_obj, new_date_str, new_path = new_info

    def read_dashboard(filepath):
        tables = pd.read_html(filepath)
        # Table 1: 团队进度
        team_df = tables[1].copy()
        team_df["归属团队名称"] = team_df["归属团队名称"].astype(str).str.strip()
        team_df = team_df[team_df["归属团队名称"] != "内部团队（已脱敏）"].copy()
        for col in ["总量", "已标注", "质检中", "验收中", "已驳回", "已完成"]:
            team_df[col] = pd.to_numeric(team_df[col], errors="coerce").fillna(0).astype(int)

        # Table 0: 批次明细 → 按团队聚合质量
        batch_df = tables[0].copy()
        # 兼容新旧列名: "批次名称" vs "批次名称/标注员"
        batch_name_col = "批次名称/标注员" if "批次名称/标注员" in batch_df.columns else "批次名称"
        batch_df["团队"] = batch_df[batch_name_col].astype(str).str.extract(r"\[([^\]]+)\]")
        batch_df["团队"] = batch_df["团队"].fillna("未知团队")

        qc_parsed = batch_df["🛡️首次质检"].apply(parse_quality_cell)
        ac_parsed = batch_df["🎯首次验收"].apply(parse_quality_cell)
        batch_df["质检_pass"] = qc_parsed.apply(lambda x: x[0])
        batch_df["质检_total"] = qc_parsed.apply(lambda x: x[1])
        batch_df["验收_pass"] = ac_parsed.apply(lambda x: x[0])
        batch_df["验收_total"] = ac_parsed.apply(lambda x: x[1])

        quality_by_team = {}
        for team, grp in batch_df.groupby("团队"):
            if "未分配" in team:
                continue
            quality_by_team[team] = {
                "质检_pass": int(grp["质检_pass"].sum()),
                "质检_total": int(grp["质检_total"].sum()),
                "验收_pass": int(grp["验收_pass"].sum()),
                "验收_total": int(grp["验收_total"].sum()),
            }
        # 保留批次明细
        detail_cols = [batch_name_col, "团队", "质检_pass", "质检_total", "验收_pass", "验收_total"]
        batch_detail = batch_df[[c for c in detail_cols if c in batch_df.columns]].copy()
        batch_detail = batch_detail.rename(columns={batch_name_col: "批次名称"})
        return team_df, quality_by_team, batch_detail

    team_old, qc_old, batch_old = read_dashboard(old_path)
    team_new, qc_new, batch_new = read_dashboard(new_path)

    # 进度对比
    progress_cols = ["总量", "已标注", "质检中", "验收中", "已驳回", "已完成"]
    p_old = team_old[["归属团队名称"] + progress_cols].copy()
    p_old.columns = ["团队"] + [f"{c}_old" for c in progress_cols]
    p_new = team_new[["归属团队名称"] + progress_cols].copy()
    p_new.columns = ["团队"] + [f"{c}_new" for c in progress_cols]
    progress = p_old.merge(p_new, on="团队", how="outer")

    for col in progress_cols:
        progress[f"{col}_old"] = progress[f"{col}_old"].fillna(0).astype(int)
        progress[f"{col}_new"] = progress[f"{col}_new"].fillna(0).astype(int)
        progress[f"{col}_差值"] = progress[f"{col}_new"] - progress[f"{col}_old"]
        progress[f"{col}_波动率"] = progress.apply(
            lambda r, c=col: calc_fluctuation(r[f"{c}_new"], r[f"{c}_old"]), axis=1
        )
        progress[f"{col}_标记"] = progress[f"{col}_波动率"].apply(flag_warn)

    # 质量对比
    all_teams = sorted(set(list(qc_old.keys()) + list(qc_new.keys())))
    quality_rows = []
    for team in all_teams:
        o = qc_old.get(team, {"质检_pass": 0, "质检_total": 0, "验收_pass": 0, "验收_total": 0})
        n = qc_new.get(team, {"质检_pass": 0, "质检_total": 0, "验收_pass": 0, "验收_total": 0})

        qc_rate_old = calc_rate(o["质检_pass"], o["质检_total"])
        qc_rate_new = calc_rate(n["质检_pass"], n["质检_total"])
        ac_rate_old = calc_rate(o["验收_pass"], o["验收_total"])
        ac_rate_new = calc_rate(n["验收_pass"], n["验收_total"])

        qc_diff = round(qc_rate_new - qc_rate_old, 1) if qc_rate_old is not None and qc_rate_new is not None else None
        ac_diff = round(ac_rate_new - ac_rate_old, 1) if ac_rate_old is not None and ac_rate_new is not None else None
        qc_fluct = calc_fluctuation(qc_rate_new, qc_rate_old)
        ac_fluct = calc_fluctuation(ac_rate_new, ac_rate_old)

        quality_rows.append({
            "团队": team,
            "质检_pass_old": o["质检_pass"], "质检_total_old": o["质检_total"],
            "质检_pass_new": n["质检_pass"], "质检_total_new": n["质检_total"],
            "质检通过率_old": qc_rate_old, "质检通过率_new": qc_rate_new,
            "质检差值": qc_diff, "质检波动率": qc_fluct, "质检标记": flag_combined(qc_fluct, qc_rate_new, QUALIFY_THRESHOLD_QC),
            "验收_pass_old": o["验收_pass"], "验收_total_old": o["验收_total"],
            "验收_pass_new": n["验收_pass"], "验收_total_new": n["验收_total"],
            "验收通过率_old": ac_rate_old, "验收通过率_new": ac_rate_new,
            "验收差值": ac_diff, "验收波动率": ac_fluct, "验收标记": flag_combined(ac_fluct, ac_rate_new, QUALIFY_THRESHOLD_AC),
        })

    quality = pd.DataFrame(quality_rows)

    return {
        "old_date": old_date_str,
        "new_date": new_date_str,
        "progress": progress,
        "quality": quality,
        "batch_new": batch_new,
        "batch_old": batch_old,
    }


# ============================================================================
# 数据分析与意见生成
# ============================================================================

def generate_analysis(team_results: dict, dash_results: dict, supplier_name: str) -> list:
    """
    基于团队绩效和项目大盘数据，生成数据分析意见。
    返回意见字符串列表。
    """
    comments = []

    # --- 团队绩效分析 ---
    for ftype_key, label in [("历史全量", "全量"), ("当天实时", "当天")]:
        if ftype_key not in team_results:
            continue
        r = team_results[ftype_key]
        df = r["data"]

        # 只分析本期有数据的标注员
        active = df[df["已标注_new"].notna() & (df["已标注_new"] > 0)].copy()

        if active.empty:
            continue

        # 1. 合格率统计（标记列含 "✅ 合格" 或 "🌟 优秀" 均为达标）
        for metric, display, thr in [("质检通过率", "质检", QUALIFY_THRESHOLD_QC),
                                      ("验收首次通过率", "首次验收", QUALIFY_THRESHOLD_AC),
                                      ("验收累积通过率", "累积验收", QUALIFY_THRESHOLD_AC)]:
            mark_col = f"{metric}_标记"
            if mark_col not in df.columns:
                continue
            total = len(active)
            qualified = active[active[mark_col].isin(["✅ 合格", "🌟 优秀"])]
            rate = round(len(qualified) / total * 100, 1) if total > 0 else 0
            if rate < 80:
                comments.append(
                    f"⚠️ [{supplier_name}][{label}][{display}] 合格率仅 {rate}% ({len(qualified)}/{total})，"
                    f"需关注未达标标注员（阈值≥{thr}%）"
                )

        # 2. 异常波动标注员
        for metric, display in [("质检通过率", "质检"), ("验收首次通过率", "首次验收"), ("验收累积通过率", "累积验收")]:
            warn_col = f"{metric}_标记"
            if warn_col not in df.columns:
                continue
            warned = active[active[warn_col] == "⚠️ 异常"]
            if len(warned) > 0:
                names = "、".join(warned["标注员"].tolist())
                comments.append(
                    f"⚠️ [{supplier_name}][{label}][{display}] 负向波动>10%: {names}"
                )

        # 3. 优秀波动标注员
        for metric, display in [("质检通过率", "质检"), ("验收首次通过率", "首次验收"), ("验收累积通过率", "累积验收")]:
            warn_col = f"{metric}_标记"
            if warn_col not in df.columns:
                continue
            excelled = active[active[warn_col] == "🌟 优秀"]
            if len(excelled) > 0:
                names = "、".join(excelled["标注员"].tolist())
                comments.append(
                    f"🌟 [{supplier_name}][{label}][{display}] 正向波动>10%: {names}"
                )

        # 4. 极低通过率 (< 60%)
        for metric, display in [("质检通过率", "质检"), ("验收首次通过率", "首次验收")]:
            new_col = f"{metric}_new"
            if new_col not in df.columns:
                continue
            low = active[(active[new_col].notna()) & (active[new_col] < 60)]
            if len(low) > 0:
                names = "、".join(low["标注员"].tolist())
                comments.append(
                    f"🔴 [{supplier_name}][{label}][{display}] 通过率<60%: {names}"
                )

        # 5. 高产量但低质量
        if "质检通过率_new" in df.columns and "已标注_new" in df.columns:
            avg_output = active["已标注_new"].mean()
            low_quality = active[
                (active["已标注_new"] > avg_output * 1.2) &
                (active["质检通过率_new"].notna()) &
                (active["质检通过率_new"] < QUALIFY_THRESHOLD_QC)
            ]
            if len(low_quality) > 0:
                names = "、".join(low_quality["标注员"].tolist())
                comments.append(
                    f"💡 [{supplier_name}][{label}] 产量高但质检不达标: {names}，建议优先复查"
                )

    # --- 项目大盘分析 ---
    for ftype_key, label in [("历史全量", "全量"), ("当天实时", "当天")]:
        if ftype_key not in dash_results:
            continue
        r = dash_results[ftype_key]
        quality = r["quality"]
        progress = r["progress"]

        if quality.empty:
            continue

        # 团队质量合格率（标记列含 "✅ 合格" 或 "🌟 优秀" 均为达标）
        for qtype, display in [("质检", "质检"), ("验收", "验收")]:
            mcol = f"{qtype}标记"
            if mcol in quality.columns:
                total = len(quality)
                qualified = quality[quality[mcol].isin(["✅ 合格", "🌟 优秀"])]
                rate = round(len(qualified) / total * 100, 1) if total > 0 else 0
                comments.append(
                    f"📊 [{supplier_name}][{label}][大盘-{display}] 团队合格率: {rate}% ({len(qualified)}/{total})"
                )

        # 异常团队
        for qtype, display in [("质检", "质检"), ("验收", "验收")]:
            mcol = f"{qtype}标记"
            if mcol in quality.columns:
                warned = quality[quality[mcol] == "⚠️ 异常"]
                if len(warned) > 0:
                    teams = "、".join(warned["团队"].tolist())
                    comments.append(
                        f"⚠️ [{supplier_name}][{label}][大盘-{display}] 通过率下滑: {teams}"
                    )

        # 进度停滞
        if not progress.empty:
            stagnant = progress[
                (progress["已完成_new"] == progress["已完成_old"]) &
                (progress["已标注_new"] == progress["已标注_old"]) &
                (progress["已标注_old"] > 0)
            ]
            if len(stagnant) > 0:
                teams = "、".join(stagnant["团队"].tolist())
                comments.append(
                    f"⏸️ [{supplier_name}][{label}][大盘] 进度无变化团队: {teams}"
                )

    return comments


# ============================================================================
# 报告格式化
# ============================================================================

def _fmt_num(n) -> str:
    """格式化数字，带千分位"""
    if n is None:
        return "-"
    return f"{int(n):,}"


def _fmt_pct(n, digits=1) -> str:
    """格式化百分比"""
    if n is None or (isinstance(n, float) and math.isnan(n)):
        return "-"
    return f"{round(float(n), digits)}%"


def _safe_int(v) -> int:
    """安全转 int"""
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return 0


def _get_week_info(date_str: str):
    """从 '2026_6_15' 格式推断周次和星期"""
    parts = date_str.split("_")
    if len(parts) != 3:
        return "?", "?"
    d = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
    week_num = d.isocalendar()[1]
    # 项目起始周为第19周(5月初)的偏移
    week_label = f"第{week_num - 18}周" if week_num >= 19 else f"第{week_num}周"
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    wd = weekdays[d.weekday()]
    return week_label, wd


def format_full_report(all_team, all_dash, all_analysis, run_time: str, report_type: str = "日报") -> str:
    """生成完整 TXT 日报/周报（新版模板格式）"""
    is_weekly = (report_type == "周报")
    rp = "周" if is_weekly else "日"
    lines = []
    W = 90  # 内容区宽

    for project_name, suppliers in PROJECTS.items():
        # ── 聚合项目数据 ──
        team_daily = {}
        team_full = {}
        dash_daily = {}
        dash_full = {}

        for sup in suppliers:
            td = all_team.get(project_name, {}).get(sup, {})
            dd = all_dash.get(project_name, {}).get(sup, {})
            if "当天实时" in td:
                team_daily[sup] = td["当天实时"]
            if "历史全量" in td:
                team_full[sup] = td["历史全量"]
            if "当天实时" in dd:
                dash_daily[sup] = dd["当天实时"]
            if "历史全量" in dd:
                dash_full[sup] = dd["历史全量"]

        if not dash_daily and not dash_full:
            continue

        # 确定日期和对比基准（从历史全量取，当天实时可能被重置）
        sample_rd = next(iter(dash_full.values())) if dash_full else \
                    next(iter(dash_daily.values()))
        new_date = sample_rd.get("new_date_str", sample_rd.get("new_date", "?"))
        old_date = sample_rd.get("old_date_str", sample_rd.get("old_date", "?"))
        week_label, weekday = _get_week_info(new_date)

        # ── 计算 KPI（日增量用历史全量，不用当天实时，当天实时可能因数据重置而不准）──
        today_anno = 0      # 今日标注量（日增量）
        today_deliv = 0     # 今日交付量（日增量）
        cumul_anno = 0      # 累计标注量
        cumul_deliv = 0     # 累计交付量
        total_anno_plan = 0 # 总计划量（总量）

        # 标注员统计
        active_count = 0
        total_count = 0
        supplier_active = {}

        for sup in suppliers:
            # 日增量：从历史全量计算（已标注_new - 已标注_old = 日增量）
            if sup in dash_full:
                rf = dash_full[sup]
                prog = rf.get("progress")
                if prog is not None and not prog.empty:
                    for _, prow in prog.iterrows():
                        today_anno += _safe_int(prow.get("已标注_差值", 0))
                        today_deliv += _safe_int(prow.get("已完成_差值", 0))
                        cumul_anno += _safe_int(prow.get("已标注_new", 0))
                        cumul_deliv += _safe_int(prow.get("已完成_new", 0))
                        total_anno_plan += _safe_int(prow.get("总量_new", 0))
            elif sup in dash_daily:
                # fallback: 无历史全量时用当天实时
                rd = dash_daily[sup]
                prog = rd.get("progress")
                if prog is not None and not prog.empty:
                    for _, prow in prog.iterrows():
                        today_anno += _safe_int(prow.get("已标注_差值", 0))
                        today_deliv += _safe_int(prow.get("已完成_差值", 0))

            # 标注员人数：用历史全量统计（已标注_差值 > 0 才是今天有产出的）
            if sup in team_full:
                tf = team_full[sup]
                df = tf.get("data")
                if df is not None and not df.empty:
                    for _, row in df.iterrows():
                        inc = _safe_int(row.get("已标注_差值", 0))
                        if inc > 0:
                            active_count += 1
                        total_count += 1
                    supplier_active[sup] = sum(
                        1 for _, row in df.iterrows()
                        if _safe_int(row.get("已标注_差值", 0)) > 0
                    )
            elif sup in team_daily:
                # fallback: 无历史全量时用当天实时（已标注_new > 0 即认为当天活跃）
                td = team_daily[sup]
                df = td.get("data")
                if df is not None and not df.empty and sup not in supplier_active:
                    supplier_active[sup] = sum(
                        1 for _, row in df.iterrows()
                        if _safe_int(row.get("已标注_new", 0)) > 0
                    )
                    active_count += supplier_active[sup]
                    total_count += len(df)

        attendance_pct = round(active_count / total_count * 100, 1) if total_count > 0 else 0

        # ── 头部 ──
        safe_new_date = new_date.replace("_", "/")
        lines.append("═" * W)
        report_en = "Weekly Report" if is_weekly else "Daily Report"
        lines.append(vis_center(f"{project_name}  {report_en} | {week_label} | {weekday} | 正常产出", W))
        lines.append(vis_center(safe_new_date, W))
        lines.append("═" * W)
        lines.append("")

        # ── KPI 指标栏 ──
        if is_weekly:
            kpi_labels = ["本周累计标注量", "标注周增量", "本周累计已完成量", "已完成周增量", "出勤率"]
            kpi_sub = [safe_new_date, "对比上期", safe_new_date, "对比上期", f"{active_count}/{total_count} 人在岗"]
        else:
            kpi_labels = ["今日累计标注量", "标注日增量", "今日累计已完成量", "已完成日增量", "出勤率"]
            kpi_sub = [safe_new_date, "对比上期", safe_new_date, "对比上期", f"{active_count}/{total_count} 人在岗"]
        kpi_values = [
            _fmt_num(cumul_anno),
            _fmt_num(today_anno),
            _fmt_num(cumul_deliv),
            _fmt_num(today_deliv),
            f"{attendance_pct}%",
        ]

        # KPI box
        total_kpi_w = W - 2
        col_w = total_kpi_w // 5
        gaps = total_kpi_w - col_w * 5

        kpi_row1 = ""
        kpi_row2 = ""
        kpi_row3 = ""
        for i in range(5):
            w = col_w + (1 if i < gaps else 0)
            kpi_row1 += vis_center(kpi_labels[i], w)
            kpi_row2 += vis_center(kpi_values[i], w)
            kpi_row3 += vis_center(kpi_sub[i], w)

        lines.append("┌" + "─" * total_kpi_w + "┐")
        lines.append("│" + kpi_row1 + "│")
        lines.append("│" + kpi_row2 + "│")
        lines.append("│" + kpi_row3 + "│")
        lines.append("└" + "─" * total_kpi_w + "┘")
        lines.append("")

        # 累计完成率
        cumul_anno_pct = round(cumul_anno / total_anno_plan * 100, 1) if total_anno_plan > 0 else 0
        cumul_deliv_pct = round(cumul_deliv / cumul_anno * 100, 1) if cumul_anno > 0 else 0
        # 计算对比窗口天数
        try:
            od_parts = old_date.split("_")
            nd_parts = new_date.split("_")
            od_dt = datetime(int(od_parts[0]), int(od_parts[1]), int(od_parts[2]))
            nd_dt = datetime(int(nd_parts[0]), int(nd_parts[1]), int(nd_parts[2]))
            win_days = (nd_dt - od_dt).days
            win_str = f"对比窗口：{old_date.replace('_', '/')} → {new_date.replace('_', '/')}（{win_days} 天）"
        except:
            win_str = f"对比窗口：{old_date.replace('_', '/')} → {new_date.replace('_', '/')}"
        lines.append(f"  较计划总量 {_fmt_num(total_anno_plan)} 完成 {cumul_anno_pct}%  |  已标注 {_fmt_num(cumul_anno)} 条，其中已完成 {_fmt_num(cumul_deliv)} 条（{cumul_deliv_pct}%）")
        lines.append(f"  {win_str}")
        lines.append("")

        # ── 人员在岗情况及重点工作 ──
        prefix_t = "本周" if is_weekly else "今日"
        lines.append("━" * W)
        lines.append(f"  ○ 人员在岗情况及{prefix_t}重点工作")
        lines.append("━" * W)
        lines.append("")
        lines.append("  一、人员在岗情况")
        lines.append(f"  应出勤 {total_count} 人（账号），{prefix_t}有产出 {active_count} 人，出勤率 {attendance_pct}%。")
        for sup in suppliers:
            sa = supplier_active.get(sup, 0)
            sup_total = 0
            if sup in team_full:
                df = team_full[sup].get("data")
                if df is not None and not df.empty:
                    sup_total = len(df)
            if sup_total == 0 and sup in team_daily:
                df = team_daily[sup].get("data")
                if df is not None and not df.empty:
                    sup_total = len(df)
            lines.append(f"    {sup}：{sa}/{sup_total} 人有产出")
        lines.append("")

        proj_comments = [c for c in all_analysis if any(f"[{s}]" in c or f"[{s}][" in c for s in suppliers)]
        warn_comments = [c for c in proj_comments if c.startswith("⚠️") or c.startswith("🔴")]
        excel_comments = [c for c in proj_comments if c.startswith("🌟")]
        low_qual_comments = [c for c in proj_comments if "合格率仅" in c or "通过率<60%" in c]
        stagnant_comments = [c for c in proj_comments if "进度无变化" in c]

        lines.append(f"  二、{prefix_t}重点工作")
        if warn_comments:
            lines.append(f"  1. 关注异常预警：共 {len(warn_comments)} 项，需对波动超阈值的标注员/团队进行复查。")
            for c in warn_comments[:3]:
                lines.append(f"     - {c}")
        if low_qual_comments:
            lines.append("  2. 低质量标注员需重点复核，合格率/通过率未达标的条目需逐一排查。")
        if excel_comments:
            lines.append(f"  3. 正向优秀：{len(excel_comments)} 项表现突出，可总结经验推广。")
        if stagnant_comments:
            lines.append(f"  4. 进度停滞团队需催促推进：{len(stagnant_comments)} 项。")
        if not warn_comments and not low_qual_comments:
            lines.append(f"  1. {prefix_t}各项指标正常，整体节奏稳定。")
        lines.append("")

        # ── 计划 vs 实际产出 ──
        plan_title = "本周计划 vs 实际产出" if is_weekly else "今日计划 vs 实际产出"
        lines.append("━" * W)
        lines.append(f"  ○ {plan_title}")
        lines.append("━" * W)
        lines.append("")

        act_label = "本周实际" if is_weekly else "今日实际"
        plan_header = f"{'维度':<12} {act_label:>8} {'人效':>8} {'累计':>10} {'累计完成':>8}"
        lines.append(f"  {plan_header}")
        lines.append(f"  {'-' * 56}")

        for sup in suppliers:
            sup_day_anno = 0
            sup_day_deliv = 0
            sup_cum_anno = 0
            sup_cum_deliv = 0
            sup_total = 0

            if sup in dash_full:
                prog = dash_full[sup].get("progress")
                if prog is not None and not prog.empty:
                    sup_day_anno = sum(_safe_int(prow.get("已标注_差值", 0)) for _, prow in prog.iterrows())
                    sup_day_deliv = sum(_safe_int(prow.get("已完成_差值", 0)) for _, prow in prog.iterrows())
                    sup_cum_anno = sum(_safe_int(prow.get("已标注_new", 0)) for _, prow in prog.iterrows())
                    sup_cum_deliv = sum(_safe_int(prow.get("已完成_new", 0)) for _, prow in prog.iterrows())
                    sup_total = sum(_safe_int(prow.get("总量_new", 0)) for _, prow in prog.iterrows())
            elif sup in dash_daily:
                prog = dash_daily[sup].get("progress")
                if prog is not None and not prog.empty:
                    sup_day_anno = sum(_safe_int(prow.get("已标注_差值", 0)) for _, prow in prog.iterrows())
                    sup_day_deliv = sum(_safe_int(prow.get("已完成_差值", 0)) for _, prow in prog.iterrows())

            sup_cum_pct = f"{round(sup_cum_anno / sup_total * 100, 1)}%" if sup_total > 0 else "-"
            sup_eff = f"{round(sup_day_anno / sa, 2)}" if (sa := supplier_active.get(sup, 0)) > 0 else "-"
            lines.append(f"  {sup:<12} {_fmt_num(sup_day_anno):>8} {sup_eff:>8} {_fmt_num(sup_cum_anno):>10} {sup_cum_pct:>8}")

        lines.append("")
        inc_label = "周" if is_weekly else "日"
        lines.append(f"  * 标注{inc_label}增量 {_fmt_num(today_anno)}，已完成{inc_label}增量 {_fmt_num(today_deliv)}。")
        if cumul_anno_pct >= 90:
            lines.append(f"  * 累计标注达 {cumul_anno_pct}%，接近收尾，请聚焦质量验收。")
        elif cumul_anno_pct >= 70:
            lines.append(f"  * 累计标注 {cumul_anno_pct}%，进度过半，保持当前节奏。")
        else:
            lines.append(f"  * 累计标注 {cumul_anno_pct}%，尚在前期，需注意提速。")
        lines.append("")

        # ── 整体批次完成情况 ──
        lines.append("━" * W)
        lines.append("  ○ 整体批次完成情况")
        lines.append("━" * W)
        lines.append("")

        # 按供应商收集批次（只展示主批次，过滤掉拆分包/修正包）
        all_batches = []
        for sup in suppliers:
            rd = dash_full.get(sup) or dash_daily.get(sup)
            if rd is None:
                continue
            batch_df = rd.get("batch_new")
            if batch_df is None or batch_df.empty:
                continue
            for _, brow in batch_df.iterrows():
                bn = str(brow.get("批次名称", ""))
                team = str(brow.get("团队", ""))
                if "未分配" in team:
                    continue
                # 只展示主批次（非拆分包/非修正包）
                is_fix = "拆分" in bn or "修正" in bn
                if is_fix:
                    continue
                qp = _safe_int(brow.get("质检_pass", 0))
                qt = _safe_int(brow.get("质检_total", 0))
                ap = _safe_int(brow.get("验收_pass", 0))
                at = _safe_int(brow.get("验收_total", 0))
                qr = round(qp / qt * 100, 1) if qt > 0 else None
                ar = round(ap / at * 100, 1) if at > 0 else None

                # 从批次名提取批次号
                import re
                bm = re.search(r'(batch-\d+)', bn)
                short_name = bm.group(1) if bm else bn[:40]

                # 判断状态
                if ar is not None and ar >= 95 and qr is not None and qr >= 95:
                    status = "已完成"
                elif at > 0:
                    status = "验收中"
                elif qt > 0:
                    status = "质检中"
                else:
                    status = "标注中"

                all_batches.append({
                    "batch": short_name,
                    "team": team[:20],
                    "qc_rate": qr,
                    "ac_rate": ar,
                    "status": status,
                })

        if all_batches:
            # 按供应商分组展示批次
            lines.append(f"  {'批次':<18} {'供应商':<22} {'质检率':>8} {'验收率':>8} {'状态':<10}")
            lines.append(f"  {'-' * 72}")
            batches_to_show = all_batches[:40] if len(all_batches) > 40 else all_batches
            for b in batches_to_show:
                qc_s = f"{b['qc_rate']}%" if b['qc_rate'] is not None else "-"
                ac_s = f"{b['ac_rate']}%" if b['ac_rate'] is not None else "-"
                lines.append(f"  {b['batch']:<18} {b['team']:<22} {qc_s:>8} {ac_s:>8} {b['status']:<10}")
            if len(all_batches) > 40:
                lines.append(f"  ... 共 {len(all_batches)} 个主批次，以上为前40个")
            else:
                lines.append(f"  （共 {len(all_batches)} 个主批次，不含质检拆分/修正子包）")
        else:
            lines.append("  （暂无批次明细数据）")
        lines.append("")

        # ── 业务方验收明细 ──
        lines.append("━" * W)
        lines.append("  ○ 业务方验收明细")
        lines.append("━" * W)
        lines.append("")

        has_acc_data = False
        for sup in suppliers:
            # 全量累计质量
            rd_full = dash_full.get(sup)
            quality_full = rd_full.get("quality") if rd_full else pd.DataFrame()

            # 当天质量（用于日通过率）
            rd_daily = dash_daily.get(sup)
            quality_daily = rd_daily.get("quality") if rd_daily else pd.DataFrame()

            if quality_full.empty and quality_daily.empty:
                continue

            for _, qrow in quality_full.iterrows():
                team = qrow.get("团队", "")
                # 累计质检
                qc_new = qrow.get("质检通过率_new")
                qc_old = qrow.get("质检通过率_old")
                qc_pass = _safe_int(qrow.get("质检_pass_new", 0))
                qc_tot = _safe_int(qrow.get("质检_total_new", 0))
                if pd.notna(qc_new) and pd.notna(qc_old):
                    qc_diff = round(float(qc_new) - float(qc_old), 1)
                    qc_ds = f"+{qc_diff}pp" if qc_diff >= 0 else f"{qc_diff}pp"
                else:
                    qc_ds = "-"
                qc_s = f"{qc_new}% ({qc_pass}/{qc_tot}) {qc_ds}" if pd.notna(qc_new) else "-"

                # 累计验收
                ac_new = qrow.get("验收通过率_new")
                ac_old = qrow.get("验收通过率_old")
                ac_pass = _safe_int(qrow.get("验收_pass_new", 0))
                ac_tot = _safe_int(qrow.get("验收_total_new", 0))
                if pd.notna(ac_new) and pd.notna(ac_old):
                    ac_diff = round(float(ac_new) - float(ac_old), 1)
                    ac_ds = f"+{ac_diff}pp" if ac_diff >= 0 else f"{ac_diff}pp"
                else:
                    ac_ds = "-"
                ac_s = f"{ac_new}% ({ac_pass}/{ac_tot}) {ac_ds}" if pd.notna(ac_new) else "-"

                acc_status = "待验收"
                if pd.notna(ac_new) and float(ac_new) >= QUALIFY_THRESHOLD_AC:
                    acc_status = "V 已通过验收"
                elif pd.notna(qc_new) and float(qc_new) >= QUALIFY_THRESHOLD_QC:
                    acc_status = "质检通过 待验收"

                # 日通过率（从当天质量数据取）
                day_qc_s = "-"
                day_ac_s = "-"
                if not quality_daily.empty:
                    dq = quality_daily[quality_daily["团队"] == team]
                    if len(dq) > 0:
                        drow = dq.iloc[0]
                        dq_new = drow.get("质检通过率_new")
                        dq_old = drow.get("质检通过率_old")
                        da_new = drow.get("验收通过率_new")
                        da_old = drow.get("验收通过率_old")
                        dq_pass = _safe_int(drow.get("质检_pass_new", 0))
                        dq_tot = _safe_int(drow.get("质检_total_new", 0))
                        da_pass = _safe_int(drow.get("验收_pass_new", 0))
                        da_tot = _safe_int(drow.get("验收_total_new", 0))

                        if pd.notna(dq_new) and pd.notna(dq_old):
                            dq_diff = round(float(dq_new) - float(dq_old), 1)
                            dqd_s = f"+{dq_diff}pp" if dq_diff >= 0 else f"{dq_diff}pp"
                        else:
                            dqd_s = ""
                        day_qc_s = f"{dq_new}% ({dq_pass}/{dq_tot})" if pd.notna(dq_new) else "-"

                        if pd.notna(da_new) and pd.notna(da_old):
                            da_diff = round(float(da_new) - float(da_old), 1)
                            dad_s = f"+{da_diff}pp" if da_diff >= 0 else f"{da_diff}pp"
                        else:
                            dad_s = ""
                        day_ac_s = f"{da_new}% ({da_pass}/{da_tot})" if pd.notna(da_new) else "-"

                has_acc_data = True
                lines.append(f"  {team}")
                lines.append(f"    累计: 质检 {qc_s}  |  验收 {ac_s}  |  {acc_status}")
                lines.append(f"    当日: 质检 {day_qc_s}  |  验收 {day_ac_s}")
                lines.append("")

        if not has_acc_data:
            lines.append("  （暂无验收数据）")
        lines.append("")

        # ── 明日/下周工作安排 ──
        next_label = "下周" if is_weekly else "明日"
        lines.append("━" * W)
        lines.append(f"  ○ {next_label}工作安排")
        lines.append("━" * W)
        lines.append("")
        idx = 1
        if warn_comments:
            lines.append(f"  {idx}. 跟进异常标注员复查，确保质量回升至基线。"); idx += 1
        if low_qual_comments:
            lines.append(f"  {idx}. 对低质量标注员完成逐题复盘与专项培训。"); idx += 1
        if stagnant_comments:
            lines.append(f"  {idx}. 催促进度停滞团队恢复产出节奏。"); idx += 1
        lines.append(f"  {idx}. 跟进待验收批次的业务方反馈。"); idx += 1
        if cumul_anno_pct < 95:
            lines.append(f"  {idx}. 推进剩余队列标注进度，向收尾冲刺。"); idx += 1
        lines.append(f"  {idx}. 聚焦内部质检与质量校准。")
        lines.append("")

        # 项目间分隔（仅当不是最后一个项目时）
        proj_names = list(PROJECTS.keys())
        if project_name != proj_names[-1]:
            lines.append("")
            lines.append("═" * W)
            lines.append("")

    return "\n".join(lines)


def save_excel_report(all_team, all_dash, all_analysis, run_time: str) -> list:
    """生成按项目拆分的 Excel 报表，返回两个文件路径的列表"""
    TITLE_LABELS = {"历史全量": "全量", "当天实时": "当天"}
    paths = []

    for project_name, suppliers in PROJECTS.items():
        xlsx_path = os.path.join(OUTPUT_DIR, f"{project_name}验收进度_日报_{run_time.replace(':', '-').replace(' ', '_')}.xlsx")

        with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
            wb = writer.book
            header_fmt = wb.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "#FFFFFF", "align": "center", "valign": "vcenter", "text_wrap": True})
            warn_fmt = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "align": "center"})
            excel_fmt = wb.add_format({"bg_color": "#FFD700", "font_color": "#7B5800", "align": "center"})
            qualify_fmt = wb.add_format({"bg_color": "#BDD7EE", "font_color": "#1F4E79", "align": "center"})
            ok_fmt = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "align": "center"})
            normal_fmt = wb.add_format({"align": "center"})
            new_fmt = wb.add_format({"bg_color": "#E2EFDA", "font_color": "#375623", "align": "center"})

            for supplier in suppliers:
                team_data = all_team.get(project_name, {}).get(supplier, {})
                dash_data = all_dash.get(project_name, {}).get(supplier, {})

                # ── Sheet 1: {供应商} — 项目大盘对比（合并全量+当天）──
                dash_rows = []
                for ftype_key in ["历史全量", "当天实时"]:
                    if ftype_key not in dash_data:
                        continue
                    r = dash_data[ftype_key]
                    quality = r["quality"].copy()
                    progress = r["progress"].copy()
                    if quality.empty:
                        continue
                    lbl = TITLE_LABELS[ftype_key]
                    for _, qrow in quality.iterrows():
                        team = qrow["团队"]
                        prow = progress[progress["团队"] == team]
                        row = {"团队": team, "对比日期": r["old_date_str"].replace("_", "/"),
                               "数据日期": r["new_date_str"].replace("_", "/"),
                               "数据类型": lbl}
                        for col, label in [("总量", "已下发量"), ("已标注", "已标注量"),
                                           ("质检中", "质检中"), ("验收中", "验收中"),
                                           ("已驳回", "已驳回"), ("已完成", "已完成")]:
                            for suf, lab in [("_new", ""), ("_old", "(上期)")]:
                                k = f"{col}{suf}"
                                row[f"{label}{lab}"] = int(prow.iloc[0][k]) if len(prow) > 0 and k in prow.columns else ""
                            dk = f"{col}_差值"
                            row[f"{label}增量"] = int(prow.iloc[0][dk]) if len(prow) > 0 and dk in prow.columns else ""
                        for qt, label in [("质检", "质检通过率"), ("验收", "验收通过率")]:
                            row[label] = qrow.get(f"{qt}通过率_new", "")
                            row[f"{label}(上期)"] = qrow.get(f"{qt}通过率_old", "")
                            row[f"{label}波动率"] = qrow.get(f"{qt}波动率", "")
                            row[f"{label}标记"] = qrow.get(f"{qt}标记", "")
                            row[f"{qt}通过数"] = int(qrow.get(f"{qt}_pass_new", 0) or 0)
                            row[f"{qt}总量"] = int(qrow.get(f"{qt}_total_new", 0) or 0)
                        dash_rows.append(row)
                if dash_rows:
                    ddf = pd.DataFrame(dash_rows)
                    sheet_name = f"{supplier}"[:31]
                    ddf.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
                    _excel_write_headers_and_format(writer.sheets[sheet_name], ddf, header_fmt, warn_fmt, excel_fmt, qualify_fmt, ok_fmt, normal_fmt, new_fmt)

                # ── Sheet 2: {供应商}人力 — 标注员绩效对比（合并全量+当天）──
                hr_rows = []
                for ftype_key in ["历史全量", "当天实时"]:
                    if ftype_key not in team_data:
                        continue
                    r = team_data[ftype_key]
                    df = r["data"].copy()
                    lbl = TITLE_LABELS[ftype_key]
                    for idx, row in df.iterrows():
                        name = row["标注员"]
                        if pd.isna(row.get("已标注_new")) or (isinstance(row.get("已标注_new"), float) and math.isnan(row.get("已标注_new"))):
                            continue
                        rec = {
                            "标注员": name,
                            "对比日期": r["old_date_str"].replace("_", "/"),
                            "数据日期": r["new_date_str"].replace("_", "/"),
                            "数据类型": lbl,
                            "已标注量": int(row.get("已标注_new", 0) or 0),
                            "已标注量(上期)": int(row.get("已标注_old", 0) or 0),
                            "已标注增量": int(row.get("已标注_差值", 0) or 0),
                        }
                        for m, label, thr in [
                            ("质检通过率", "质检通过率", QUALIFY_THRESHOLD_QC),
                            ("验收首次通过率", "首次验收通过率", QUALIFY_THRESHOLD_AC),
                            ("验收累积通过率", "累积验收通过率", QUALIFY_THRESHOLD_AC),
                        ]:
                            rec[label] = row.get(f"{m}_new", "")
                            rec[f"{label}(上期)"] = row.get(f"{m}_old", "")
                            rec[f"{label}波动率"] = row.get(f"{m}_波动率", "")
                            rec[f"{label}标记"] = row.get(f"{m}_标记", "")
                            pcol = _metric_to_pass_col(m)
                            tcol = _metric_to_total_col(m)
                            pval = row.get(f"{pcol}_new", 0)
                            tval = row.get(f"{tcol}_new", 0)
                            rec[f"{label}通过数"] = int(pval) if _valid(pval) else 0
                            rec[f"{label}总量"] = int(tval) if _valid(tval) else 0
                        hr_rows.append(rec)
                if hr_rows:
                    hdf = pd.DataFrame(hr_rows)
                    sheet_name = f"{supplier}人力"[:31]
                    hdf.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
                    _excel_write_headers_and_format(writer.sheets[sheet_name], hdf, header_fmt, warn_fmt, excel_fmt, qualify_fmt, ok_fmt, normal_fmt, new_fmt)

                # ── Sheet 3: {供应商}批次明细（合并全量+当天）──
                batch_rows = []
                for ftype_key in ["历史全量", "当天实时"]:
                    if ftype_key not in dash_data:
                        continue
                    r = dash_data[ftype_key]
                    batch_new = r.get("batch_new")
                    if batch_new is None or batch_new.empty:
                        continue
                    lbl = TITLE_LABELS[ftype_key]
                    for _, brow in batch_new.iterrows():
                        qc_r = round(brow["质检_pass"] / brow["质检_total"] * 100, 1) if brow["质检_total"] > 0 else ""
                        ac_r = round(brow["验收_pass"] / brow["验收_total"] * 100, 1) if brow["验收_total"] > 0 else ""
                        batch_rows.append({
                            "批次名称": brow["批次名称"],
                            "团队": brow["团队"],
                            "数据日期": r["new_date_str"].replace("_", "/"),
                            "数据类型": lbl,
                            "质检通过数": int(brow["质检_pass"]),
                            "质检总量": int(brow["质检_total"]),
                            "质检通过率": qc_r,
                            "验收通过数": int(brow["验收_pass"]),
                            "验收总量": int(brow["验收_total"]),
                            "验收通过率": ac_r,
                        })
                if batch_rows:
                    bdf = pd.DataFrame(batch_rows)
                    sheet_name = f"{supplier}批次明细"[:31]
                    bdf.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
                    _excel_write_headers_and_format(writer.sheets[sheet_name], bdf, header_fmt, warn_fmt, excel_fmt, qualify_fmt, ok_fmt, normal_fmt, new_fmt)

            # ── Sheet: 数据分析意见 ──
            proj_comments = [c for c in all_analysis if any(s in c for s in suppliers)]
            if proj_comments:
                adf = pd.DataFrame({"序号": range(1, len(proj_comments) + 1), "分析意见": proj_comments})
                adf.to_excel(writer, sheet_name="数据分析意见", index=False, startrow=1, header=False)
                ws = writer.sheets["数据分析意见"]
                for ci, cn in enumerate(adf.columns):
                    ws.write(0, ci, cn, header_fmt)
                for row_idx in range(len(adf)):
                    ws.write(row_idx + 1, 0, row_idx + 1, normal_fmt)
                    ws.write(row_idx + 1, 1, adf.iloc[row_idx, 1], normal_fmt)
                ws.set_column(0, 0, 6)
                ws.set_column(1, 1, 90)

        paths.append(xlsx_path)
        print(f"   [{project_name}] Excel 已保存: {xlsx_path}")

    return paths


def _excel_write_headers_and_format(ws, df, header_fmt, warn_fmt, excel_fmt, qualify_fmt, ok_fmt, normal_fmt, new_fmt=None):
    """写入表头（第0行）+ 标记列条件格式 + 本期数据绿色高亮，自动列宽"""
    _SKIP_GREEN = ("标记", "(上期)", "增量", "波动率", "对比日期", "数据日期", "数据类型",
                   "团队", "标注员", "批次名称", "序号", "分析意见", "项目名称", "供应商")
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, col_name, header_fmt)
    for row_idx in range(len(df)):
        for col_idx, col_name in enumerate(df.columns):
            val = df.iloc[row_idx, col_idx]
            if isinstance(val, float) and pd.isna(val):
                val = ""
            if "标记" in col_name:
                if val == "⚠️ 异常":
                    ws.write(row_idx + 1, col_idx, val, warn_fmt)
                elif val == "❌ 不合格":
                    ws.write(row_idx + 1, col_idx, val, warn_fmt)
                elif val == "🌟 优秀":
                    ws.write(row_idx + 1, col_idx, val, excel_fmt)
                elif val == "✅ 合格":
                    ws.write(row_idx + 1, col_idx, val, qualify_fmt)
                else:
                    ws.write(row_idx + 1, col_idx, val, ok_fmt)
            elif new_fmt is not None and not any(k in col_name for k in _SKIP_GREEN):
                ws.write(row_idx + 1, col_idx, val, new_fmt)
            else:
                ws.write(row_idx + 1, col_idx, val, normal_fmt)
    # 自动列宽
    for col_idx, col_name in enumerate(df.columns):
        max_w = vis_width(str(col_name)) + 4
        for row_idx in range(len(df)):
            v = str(df.iloc[row_idx, col_idx]) if not (isinstance(df.iloc[row_idx, col_idx], float) and pd.isna(df.iloc[row_idx, col_idx])) else ""
            max_w = max(max_w, vis_width(v) + 4)
        ws.set_column(col_idx, col_idx, min(max_w, 40))


# ============================================================================
# 邮件发送
# ============================================================================

def send_email(txt_path: str, xlsx_paths: list, run_time: str):
    """发送邮件"""
    password = _load_email_password()
    if not password:
        print("\n⚠️ 未配置邮箱密码/授权码，跳过邮件发送。")
        print("  配置方法（任选其一）：")
        print(f"    1. [推荐] 创建文件 {_EMAIL_CONFIG_FILE}")
        print('       内容: {"password": "你的邮箱授权码"}')
        print("    2. 设置环境变量: export EMAIL_PASSWORD='你的邮箱授权码'")
        print("    3. 直接在脚本 EMAIL_CONFIG 中填写 password 字段")
        return False

    msg = MIMEMultipart()
    msg["From"] = EMAIL_CONFIG["sender"]
    msg["To"] = ", ".join(EMAIL_CONFIG["to"])
    msg["Subject"] = f"项目管理统一日报 - {run_time}"

    # 正文
    body = f"""各位好，

附件为项目标注质量统一日报，覆盖以下项目与供应商：

  • 项目A_DATA：供应商Alpha、供应商Beta、供应商Gamma
  • 项目B_ANOMALY：供应商Delta

日报包含团队绩效波动（标注员级）和项目大盘质量波动（团队级），
分为全量日报和当天日报两个维度。

详细数据请见附件 TXT 和 Excel。

---
自动生成于 {run_time}
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # 附件
    all_attachments = [txt_path] + list(xlsx_paths)
    for path in all_attachments:
        if os.path.exists(path):
            with open(path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f'attachment; filename="{os.path.basename(path)}"'
                )
                msg.attach(part)

    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL(EMAIL_CONFIG["smtp_host"], EMAIL_CONFIG["smtp_port"], context=ctx) as server:
            server.login(EMAIL_CONFIG["sender"], password)
            server.sendmail(EMAIL_CONFIG["sender"], EMAIL_CONFIG["to"], msg.as_string())
        print(f"\n📧 邮件已发送至: {', '.join(EMAIL_CONFIG['to'])}")
        return True
    except smtplib.SMTPAuthenticationError:
        print("\n❌ 邮件认证失败，请检查邮箱密码/授权码。")
        return False
    except Exception as e:
        print(f"\n❌ 邮件发送失败: {e}")
        return False


# ============================================================================
# 日报图片渲染
# ============================================================================

def _get_font(size: int, bold: bool = False):
    """获取中文字体"""
    import PIL.ImageFont
    try:
        # Hiragino Sans GB.ttc 是 TrueType Collection，需要指定索引
        # 索引 0 = W3 (常规), 索引 1 = W6 (粗体)
        idx = 1 if bold else 0
        return PIL.ImageFont.truetype("/System/Library/Fonts/Hiragino Sans GB.ttc", size, index=idx)
    except Exception:
        # fallback: 尝试按名称加载
        try:
            return PIL.ImageFont.truetype("Hiragino Sans GB", size) if not bold else \
                   PIL.ImageFont.truetype("Hiragino Sans GB W6", size)
        except Exception:
            return PIL.ImageFont.load_default()


_SMALL_FONT = lambda: _get_font(14)
_NORMAL_FONT = lambda: _get_font(16)
_BOLD_FONT = lambda: _get_font(16, True)
_TITLE_FONT = lambda: _get_font(22, True)
_H1_FONT = lambda: _get_font(20, True)
_KPI_VAL_FONT = lambda: _get_font(28, True)


def _draw_card(draw, canvas, x, y, w, h, bg="#FFFFFF", border="#DDE2E8", radius=8):
    """绘制圆角卡片"""
    from PIL import ImageDraw as _ID
    _idr = _ID.Draw(canvas)
    _idr.rounded_rectangle([x, y, x + w, y + h], radius=radius, fill=bg, outline=border, width=1)
    return _ID.Draw(canvas)


def _draw_section_header(draw, canvas, x, y, w, text, bg="#2C3E50"):
    """绘制章节标题栏"""
    from PIL import ImageDraw as _ID
    _idr = _ID.Draw(canvas)
    _idr.rectangle([x, y, x + w, y + 36], fill=bg)
    _idr.text((x + 10, y + 7), text, fill="#FFFFFF", font=_BOLD_FONT())
    return _ID.Draw(canvas)


def _text_width(text, font):
    """计算文字像素宽"""
    try:
        return int(font.getlength(text))
    except Exception:
        return len(text) * int(font.size * 0.6)


def _draw_text(draw, x, y, text, fill="#333", font=None):
    """便捷绘制文字"""
    if font is None:
        font = _NORMAL_FONT()
    draw.text((x, y), text, fill=fill, font=font)


def _draw_table(draw, canvas, x, y, col_w, header, rows, header_bg="#3B5998", row_bg=("#FFFFFF", "#F5F7FA")):
    """绘制表格，返回表格总高度"""
    from PIL import ImageDraw as _ID
    _idr = _ID.Draw(canvas)
    row_h = 32
    ncols = len(col_w)
    total_w = sum(col_w)

    # header
    _idr.rectangle([x, y, x + total_w, y + row_h], fill=header_bg)
    cx = x
    for i, h in enumerate(header):
        _idr.text((cx + 6, y + 6), h, fill="#FFFFFF", font=_SMALL_FONT())
        cx += col_w[i]
    # grid lines
    for i in range(ncols):
        lx = x + sum(col_w[:i])
        _idr.line([lx, y, lx, y + row_h + len(rows) * row_h], fill="#D0D5DD", width=1)
    _idr.line([x, y + row_h, x + total_w, y + row_h], fill="#D0D5DD", width=1)

    # data rows
    for ri, row_data in enumerate(rows):
        ry = y + row_h + ri * row_h
        bg = row_bg[ri % 2]
        _idr.rectangle([x, ry, x + total_w, ry + row_h], fill=bg)
        cx = x
        for ci, val in enumerate(row_data):
            _idr.text((cx + 6, ry + 6), str(val), fill="#333333", font=_SMALL_FONT())
            cx += col_w[ci]
        _idr.line([x, ry + row_h, x + total_w, ry + row_h], fill="#D0D5DD", width=1)

    return row_h + len(rows) * row_h


def format_report_image(all_team, all_dash, all_analysis, run_time: str, report_type: str = "日报") -> list:
    """生成日报/周报 PNG 图片，返回文件路径列表"""
    from PIL import Image, ImageDraw

    is_weekly = (report_type == "周报")
    rp_label = "周报" if is_weekly else "日报"
    prefix_t = "本周" if is_weekly else "今日"

    paths = []
    W = 1240
    PAD = 30
    CARD_GAP = 16
    C_BLUE = "#1B3A5B"
    C_ACCENT = "#2D7DD2"
    C_GREEN = "#1B8C4A"
    C_RED = "#D14343"
    C_ORANGE = "#E67E22"
    C_BG = "#EEF1F5"
    C_CARD = "#FFFFFF"
    C_TEXT = "#2C3E50"
    C_SUB = "#6B7B8D"

    for project_name, suppliers in PROJECTS.items():
        # ── 聚合项目数据（与 format_full_report 相同逻辑）──
        dash_daily, dash_full, team_daily, team_full = {}, {}, {}, {}
        for sup in suppliers:
            td = all_team.get(project_name, {}).get(sup, {})
            dd = all_dash.get(project_name, {}).get(sup, {})
            if "当天实时" in td:
                team_daily[sup] = td["当天实时"]
            if "历史全量" in td:
                team_full[sup] = td["历史全量"]
            if "当天实时" in dd:
                dash_daily[sup] = dd["当天实时"]
            if "历史全量" in dd:
                dash_full[sup] = dd["历史全量"]

        if not dash_daily and not dash_full:
            continue

        sample_rd = next(iter(dash_full.values())) if dash_full else \
                    next(iter(dash_daily.values()))
        new_date = sample_rd.get("new_date_str", sample_rd.get("new_date", "?"))
        old_date = sample_rd.get("old_date_str", sample_rd.get("old_date", "?"))
        week_label, weekday = _get_week_info(new_date)
        safe_date = new_date.replace("_", "/")

        # ── KPI（日增量从历史全量算，当天实时可能被重置）──
        today_anno = today_deliv = cumul_anno = cumul_deliv = total_plan = 0
        active_count = total_count = 0
        supplier_active = {}
        sup_data = {}

        for sup in suppliers:
            sup_day_anno = sup_day_deliv = sup_cum_anno = sup_cum_deliv = sup_total = 0

            if sup in dash_full:
                prog = dash_full[sup].get("progress")
                if prog is not None and not prog.empty:
                    for _, prow in prog.iterrows():
                        sup_day_anno += _safe_int(prow.get("已标注_差值", 0))
                        sup_day_deliv += _safe_int(prow.get("已完成_差值", 0))
                        today_anno += _safe_int(prow.get("已标注_差值", 0))
                        today_deliv += _safe_int(prow.get("已完成_差值", 0))
                        sup_cum_anno += _safe_int(prow.get("已标注_new", 0))
                        sup_cum_deliv += _safe_int(prow.get("已完成_new", 0))
                        sup_total += _safe_int(prow.get("总量_new", 0))
                        cumul_anno += _safe_int(prow.get("已标注_new", 0))
                        cumul_deliv += _safe_int(prow.get("已完成_new", 0))
                        total_plan += _safe_int(prow.get("总量_new", 0))
            elif sup in dash_daily:
                prog = dash_daily[sup].get("progress")
                if prog is not None and not prog.empty:
                    for _, prow in prog.iterrows():
                        sup_day_anno += _safe_int(prow.get("已标注_差值", 0))
                        sup_day_deliv += _safe_int(prow.get("已完成_差值", 0))
                        today_anno += _safe_int(prow.get("已标注_差值", 0))
                        today_deliv += _safe_int(prow.get("已完成_差值", 0))

            # 活跃人数：从历史全量统计（已标注_差值 > 0 才是今天有产出）
            if sup in team_full:
                tf = team_full[sup]
                df = tf.get("data")
                if df is not None and not df.empty:
                    sa = sum(1 for _, row in df.iterrows() if _safe_int(row.get("已标注_差值", 0)) > 0)
                    supplier_active[sup] = sa
                    active_count += sa
                    total_count += len(df)
            elif sup in team_daily:
                td = team_daily[sup]
                df = td.get("data")
                if df is not None and not df.empty:
                    sa = sum(1 for _, row in df.iterrows() if _safe_int(row.get("已标注_new", 0)) > 0)
                    supplier_active[sup] = sa
                    active_count += sa
                    total_count += len(df)
            sup_data[sup] = {
                "day_anno": sup_day_anno, "day_deliv": sup_day_deliv,
                "cum_anno": sup_cum_anno, "cum_deliv": sup_cum_deliv,
                "total": sup_total
            }

        att_pct = round(active_count / total_count * 100, 1) if total_count > 0 else 0
        cumul_anno_pct = round(cumul_anno / total_plan * 100, 1) if total_plan > 0 else 0

        # ── 画布高度预估 ──
        n_teams = sum(1 for sup in suppliers for rd in [dash_full.get(sup)] if rd and not rd.get("quality", pd.DataFrame()).empty)
        n_batches = sum(1 for sup in suppliers for rd in [dash_full.get(sup)]
                        if rd and rd.get("batch_new") is not None and not rd["batch_new"].empty
                        for _, row in rd["batch_new"].iterrows()
                        if "未分配" not in str(row.get("团队", "")) and "修正" not in str(row.get("批次名称", "")) and "拆分" not in str(row.get("批次名称", "")))
        img_h = 200 + 150 + 32 * (len(suppliers) + 2) + 32 * (n_teams + 2) + 32 * (min(n_batches, 20) + 2) + 500
        img_h = max(img_h, 2000)
        img_h = img_h + 2000  # 留足余量，末尾 crop 会裁掉多余空白

        img = Image.new("RGB", (W, img_h), C_BG)
        draw = ImageDraw.Draw(img)
        y = 0

        # ═══ 1. 顶部标题栏 ═══
        draw.rectangle([0, y, W, y + 64], fill=C_BLUE)
        title = f"{project_name}  {rp_label}  |  {week_label}  |  {weekday}  |  正常产出"
        draw.text((PAD + 8, y + 12), title, fill="#FFFFFF", font=_H1_FONT())
        draw.text((PAD + 8, y + 38), safe_date, fill="#B8C5D0", font=_SMALL_FONT())
        y += 64 + CARD_GAP

        # ═══ 2. KPI 卡片 ═══
        if is_weekly:
            kpi_items = [
                ("本周累计标注量", _fmt_num(cumul_anno), safe_date),
                ("标注周增量", _fmt_num(today_anno), "对比上期"),
                ("本周累计已完成量", _fmt_num(cumul_deliv), safe_date),
                ("已完成周增量", _fmt_num(today_deliv), "对比上期"),
                ("出勤率", f"{att_pct}%", f"{active_count}/{total_count} 人在岗"),
            ]
        else:
            kpi_items = [
                ("今日累计标注量", _fmt_num(cumul_anno), safe_date),
                ("标注日增量", _fmt_num(today_anno), "对比上期"),
                ("今日累计已完成量", _fmt_num(cumul_deliv), safe_date),
                ("已完成日增量", _fmt_num(today_deliv), "对比上期"),
                ("出勤率", f"{att_pct}%", f"{active_count}/{total_count} 人在岗"),
            ]
        n_kpi = len(kpi_items)
        card_w = (W - PAD * 2 - CARD_GAP * (n_kpi - 1)) // n_kpi
        card_h = 88

        for i, (label, value, sub) in enumerate(kpi_items):
            cx = PAD + i * (card_w + CARD_GAP)
            # card bg with top accent bar
            idr = ImageDraw.Draw(img)
            idr.rounded_rectangle([cx, y, cx + card_w, y + card_h], radius=8, fill=C_CARD, outline="#D0D8E4", width=1)
            idr.rounded_rectangle([cx + 1, y + 1, cx + card_w - 1, y + 7], radius=3, fill=C_ACCENT)
            # label
            draw.text((cx + card_w // 2 - _text_width(label, _SMALL_FONT()) // 2, y + 14), label, fill=C_SUB, font=_SMALL_FONT())
            # value
            draw.text((cx + card_w // 2 - _text_width(value, _KPI_VAL_FONT()) // 2, y + 34), value, fill=C_TEXT, font=_KPI_VAL_FONT())
            # sub
            draw.text((cx + card_w // 2 - _text_width(sub, _SMALL_FONT()) // 2, y + 66), sub, fill=C_SUB, font=_SMALL_FONT())

        y += card_h + CARD_GAP

        # 累计统计 + 对比窗口
        cumul_text = f"较计划总量 {_fmt_num(total_plan)} 完成 {cumul_anno_pct}%   |   已标注 {_fmt_num(cumul_anno)} 条   |   已完成 {_fmt_num(cumul_deliv)} 条"
        try:
            od2 = old_date.split("_")
            nd2 = new_date.split("_")
            od_dt2 = datetime(int(od2[0]), int(od2[1]), int(od2[2]))
            nd_dt2 = datetime(int(nd2[0]), int(nd2[1]), int(nd2[2]))
            win_d = (nd_dt2 - od_dt2).days
            win_text = f"对比窗口：{old_date.replace('_', '/')} → {new_date.replace('_', '/')}（{win_d} 天）"
        except:
            win_text = f"对比窗口：{old_date.replace('_', '/')} → {new_date.replace('_', '/')}"
        idr = ImageDraw.Draw(img)
        idr.rounded_rectangle([PAD, y, W - PAD, y + 56], radius=6, fill=C_CARD, outline="#D0D8E4", width=1)
        draw.text((PAD + 16, y + 8), cumul_text, fill=C_TEXT, font=_NORMAL_FONT())
        draw.text((PAD + 16, y + 30), win_text, fill=C_SUB, font=_SMALL_FONT())
        y += 56 + CARD_GAP + 6

        # ═══ 3. 人员在岗情况及重点工作 ═══
        draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, f"   人员在岗情况及{prefix_t}重点工作")
        y += 36 + 10

        # 人员在岗卡片
        att_lines = [f"应出勤 {total_count} 人（账号），{prefix_t}有产出 {active_count} 人，出勤率 {att_pct}%。"]
        for sup in suppliers:
            sa = supplier_active.get(sup, 0)
            st = sup_data[sup].get("total_p", 0) or (len(team_daily[sup]["data"]) if sup in team_daily else 0)
            if sup in team_daily:
                st = len(team_daily[sup]["data"])
            att_lines.append(f"  {sup}：{sa}/{st} 人有产出")
        att_h = 12 + len(att_lines) * 22 + 12

        idr = ImageDraw.Draw(img)
        idr.rounded_rectangle([PAD, y, W - PAD, y + att_h], radius=8, fill=C_CARD, outline="#D0D8E4", width=1)
        for li, line in enumerate(att_lines):
            draw.text((PAD + 20, y + 10 + li * 22), line, fill=C_TEXT, font=_NORMAL_FONT())
        y += att_h + CARD_GAP

        # 今日重点
        proj_comments = [c for c in all_analysis if any(f"[{s}]" in c or f"[{s}][" in c for s in suppliers)]
        warn_comments = [c for c in proj_comments if c.startswith("⚠️") or c.startswith("🔴")]
        lowqual = [c for c in proj_comments if "合格率仅" in c or "通过率<60%" in c]
        foci = []
        if warn_comments:
            foci.append(f"关注异常预警：共 {len(warn_comments)} 项，需复查波动超阈值的标注员/团队。")
        if lowqual:
            foci.append("低质量标注员需重点复核，未达标条目逐一排查。")
        if not foci:
            foci.append(f"{prefix_t}各项指标正常，整体节奏稳定。")
        foci_h = 12 + len(foci) * 22 + 12

        idr = ImageDraw.Draw(img)
        idr.rounded_rectangle([PAD, y, W - PAD, y + foci_h], radius=8, fill="#FFF8E7", outline="#F0D090", width=1)
        for li, line in enumerate(foci):
            draw.text((PAD + 20, y + 10 + li * 22), f"• {line}", fill=C_TEXT, font=_NORMAL_FONT())
        y += foci_h + CARD_GAP + 4

        # ═══ 4. 计划 vs 实际产出 ═══
        section_plan = "   本周计划 vs 实际产出" if is_weekly else "   今日计划 vs 实际产出"
        draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, section_plan)
        y += 36 + 10

        plan_cols = [160, 120, 90, 130, 130, 110]
        col_anno = "本周标注" if is_weekly else "今日标注"
        col_done = "本周已完成" if is_weekly else "今日已完成"
        plan_header = ["供应商", col_anno, "人效", col_done, "累计标注", "完成率"]
        plan_rows = []
        for sup in suppliers:
            sd = sup_data[sup]
            cp = f"{round(sd['cum_anno'] / sd['total'] * 100, 1)}%" if sd["total"] > 0 else "-"
            sa = supplier_active.get(sup, 0)
            sup_eff = f"{round(sd['day_anno'] / sa, 2)}" if sa > 0 else "-"
            plan_rows.append([sup, _fmt_num(sd["day_anno"]), sup_eff, _fmt_num(sd["day_deliv"]),
                              _fmt_num(sd["cum_anno"]), cp])
        th = _draw_table(draw, img, PAD, y, plan_cols, plan_header, plan_rows)
        y += th + CARD_GAP

        # ═══ 5. 整体批次完成情况 ═══
        draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, "   整体批次完成情况")
        y += 36 + 10

        all_batches = []
        for sup in suppliers:
            rd = dash_full.get(sup) or dash_daily.get(sup)
            if rd is None:
                continue
            bdf = rd.get("batch_new")
            if bdf is None or bdf.empty:
                continue
            for _, brow in bdf.iterrows():
                bn = str(brow.get("批次名称", ""))
                team = str(brow.get("团队", ""))
                if "未分配" in team:
                    continue
                if "拆分" in bn or "修正" in bn:
                    continue
                import re as _re
                bm = _re.search(r'(batch-\d+)', bn)
                short = bm.group(1) if bm else bn[:30]
                qp = _safe_int(brow.get("质检_pass", 0))
                qt = _safe_int(brow.get("质检_total", 0))
                ap = _safe_int(brow.get("验收_pass", 0))
                at = _safe_int(brow.get("验收_total", 0))
                qr = round(qp / qt * 100, 1) if qt > 0 else None
                ar = round(ap / at * 100, 1) if at > 0 else None
                if ar is not None and ar >= 95 and qr is not None and qr >= 95:
                    st = "已完成"
                elif at > 0:
                    st = "验收中"
                elif qt > 0:
                    st = "质检中"
                else:
                    st = "标注中"
                all_batches.append([short, team[:16], f"{qr}%" if qr is not None else "-",
                                    f"{ar}%" if ar is not None else "-", st])

        if all_batches:
            bt_cols = [140, 200, 90, 90, 90]
            bt_header = ["批次", "供应商", "质检率", "验收率", "状态"]
            bt_rows = all_batches[:20] if len(all_batches) > 20 else all_batches
            th = _draw_table(draw, img, PAD, y, bt_cols, bt_header, bt_rows,
                             header_bg="#3B5998" if len(all_batches) <= 20 else "#3B5998")
            y += th + 6
            suffix = f"  （共 {len(all_batches)} 个主批次，展示前20个，不含质检拆分/修正子包）" if len(all_batches) > 20 else f"  （共 {len(all_batches)} 个主批次）"
            draw.text((PAD + 4, y), suffix, fill=C_SUB, font=_SMALL_FONT())
            y += 22 + CARD_GAP
        else:
            draw.text((PAD + 16, y), "（暂无批次数据）", fill=C_SUB, font=_NORMAL_FONT())
            y += 30 + CARD_GAP

        # ═══ 6. 验收明细 ═══
        draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, "   业务方验收明细")
        y += 36 + 10

        acc_rows = []
        for sup in suppliers:
            rd_full = dash_full.get(sup)
            quality_full = rd_full.get("quality") if rd_full else pd.DataFrame()
            rd_daily = dash_daily.get(sup)
            quality_daily = rd_daily.get("quality") if rd_daily else pd.DataFrame()

            if quality_full.empty and quality_daily.empty:
                continue
            for _, qrow in quality_full.iterrows():
                team = qrow.get("团队", "")
                # 累计
                qn = qrow.get("质检通过率_new")
                qo = qrow.get("质检通过率_old")
                an = qrow.get("验收通过率_new")
                ao = qrow.get("验收通过率_old")
                qp = _safe_int(qrow.get("质检_pass_new", 0))
                qt = _safe_int(qrow.get("质检_total_new", 0))
                ap = _safe_int(qrow.get("验收_pass_new", 0))
                at = _safe_int(qrow.get("验收_total_new", 0))
                if pd.notna(qn) and pd.notna(qo):
                    qd = round(float(qn) - float(qo), 1)
                    qds = f"+{qd}pp" if qd >= 0 else f"{qd}pp"
                    qs = f"{qn}%({qp}/{qt}){qds}"
                else:
                    qs = f"{qn}%({qp}/{qt})" if pd.notna(qn) else "-"
                if pd.notna(an) and pd.notna(ao):
                    ad = round(float(an) - float(ao), 1)
                    ads = f"+{ad}pp" if ad >= 0 else f"{ad}pp"
                    acs = f"{an}%({ap}/{at}){ads}"
                else:
                    acs = f"{an}%({ap}/{at})" if pd.notna(an) else "-"
                ast = "已通过验收" if pd.notna(an) and float(an) >= QUALIFY_THRESHOLD_AC else ("待验收" if pd.notna(qn) else "质检中")

                # 日通过率
                day_qc = "-"
                day_ac = "-"
                if not quality_daily.empty:
                    dq = quality_daily[quality_daily["团队"] == team]
                    if len(dq) > 0:
                        dr = dq.iloc[0]
                        dqn = dr.get("质检通过率_new")
                        don = dr.get("验收通过率_new")
                        dqp = _safe_int(dr.get("质检_pass_new", 0))
                        dqt = _safe_int(dr.get("质检_total_new", 0))
                        dap = _safe_int(dr.get("验收_pass_new", 0))
                        dat = _safe_int(dr.get("验收_total_new", 0))
                        if pd.notna(dqn):
                            day_qc = f"{dqn}%({dqp}/{dqt})" if dqt > 0 else f"{dqn}%"
                        if pd.notna(don):
                            day_ac = f"{don}%({dap}/{dat})" if dat > 0 else f"{don}%"

                acc_rows.append([team[:20], qs, acs, day_qc, day_ac, ast])

        if acc_rows:
            ac_cols = [200, 210, 210, 160, 160, 100]
            ac_header = ["团队", "累计质检(波动)", "累计验收(波动)", "日质检", "日验收", "状态"]
            th = _draw_table(draw, img, PAD, y, ac_cols, ac_header, acc_rows)
            y += th + CARD_GAP
        else:
            draw.text((PAD + 16, y), "（暂无验收数据）", fill=C_SUB, font=_NORMAL_FONT())
            y += 30 + CARD_GAP

        # ═══ 7. 明日/下周工作安排 ═══
        next_sec = "   下周工作安排" if is_weekly else "   明日工作安排"
        draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, next_sec)
        y += 36 + 10

        tasks = []
        if warn_comments:
            tasks.append("跟进异常标注员复查，确保质量回升至基线。")
        if lowqual:
            tasks.append("对低质量标注员完成逐题复盘与专项培训。")
        tasks.append("跟进待验收批次的业务方反馈。")
        if cumul_anno_pct < 95:
            tasks.append("推进剩余队列标注进度，向收尾冲刺。")
        tasks.append("聚焦内部质检与质量校准。")

        tasks_h = 12 + len(tasks) * 24 + 12
        idr = ImageDraw.Draw(img)
        idr.rounded_rectangle([PAD, y, W - PAD, y + tasks_h], radius=8, fill=C_CARD, outline="#D0D8E4", width=1)
        for ti, task in enumerate(tasks):
            draw.text((PAD + 20, y + 10 + ti * 24), f"{ti + 1}. {task}", fill=C_TEXT, font=_NORMAL_FONT())
        y += tasks_h + CARD_GAP + 10

        # ── 裁剪底部空白 ──
        final_h = y + PAD
        img = img.crop((0, 0, W, final_h))

        out_path = os.path.join(OUTPUT_DIR, f"{project_name}_{rp_label}_{run_time.replace(':', '-').replace(' ', '_')}.png")
        img.save(out_path, "PNG", optimize=True)
        paths.append(out_path)
        print(f"   [{project_name}] {rp_label}图片已保存: {out_path}")

    return paths


# ============================================================================
# 主流程
# ============================================================================

def _parse_date_arg(arg: str) -> str:
    """解析用户输入的日期参数，支持 '2026-6-15'、'2026-06-15'、'2026_6_15'、'2026_06_15' 等格式，
    统一转换为文件名匹配格式 'YYYY_M_D'"""
    m = re.match(r"(\d{4})[_-]?(\d{1,2})[_-]?(\d{1,2})", arg)
    if not m:
        print(f"❌ 日期格式无效: {arg}，支持格式: 2026-6-15 / 2026_6_15 / 20260615 等")
        sys.exit(1)
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    datetime(y, mo, d)  # 校验日期合法性
    return f"{y}_{mo}_{d}"


def _collect_report_data(target_date_str: str, old_date_str: str):
    """收集所有项目的日报数据，返回 (all_team, all_dash, all_analysis)"""
    all_team = defaultdict(lambda: defaultdict(dict))
    all_dash = defaultdict(lambda: defaultdict(dict))
    all_analysis = []

    for project_name, suppliers in PROJECTS.items():
        project_dir = os.path.join(ROOT_DIR, project_name)
        if not os.path.isdir(project_dir):
            print(f"[跳过] 项目目录不存在: {project_dir}")
            continue

        for supplier in suppliers:
            sdir = SUPPLIER_DIR_MAP.get(supplier, supplier)
            supplier_dir = os.path.join(project_dir, sdir)
            if not os.path.isdir(supplier_dir):
                print(f"[跳过] 供应商目录不存在: {supplier_dir}")
                continue

            team_dir = os.path.join(supplier_dir, "团队")
            dash_dir = os.path.join(supplier_dir, "整体")

            print(f"📂 处理: {project_name} / {supplier}")

            team_results = {}
            dash_results = {}

            if os.path.isdir(team_dir):
                team_results = process_team_performance(team_dir, target_date_str, old_date_str)
                for ftype, r in team_results.items():
                    r["new_date_str"] = r["new_date"]
                    r["old_date_str"] = r["old_date"]
                    all_team[project_name][supplier][ftype] = r
                    print(f"   ├─ 团队绩效 [{ftype}]: {r['new_date']} vs {r['old_date']} ({len(r['data'])} 人)")

            if os.path.isdir(dash_dir):
                dash_results = process_project_dashboard(dash_dir, target_date_str, old_date_str)
                for ftype, r in dash_results.items():
                    r["new_date_str"] = r["new_date"]
                    r["old_date_str"] = r["old_date"]
                    all_dash[project_name][supplier][ftype] = r
                    n_teams = len(r["quality"])
                    print(f"   ├─ 项目大盘 [{ftype}]: {r['new_date']} vs {r['old_date']} ({n_teams} 团队)")

            comments = generate_analysis(team_results, dash_results, supplier)
            all_analysis.extend(comments)
            for c in comments:
                print(f"   └─ {c}")

    return all_team, all_dash, all_analysis


def _generate_reports(all_team, all_dash, all_analysis, run_time, report_type="日报") -> str:
    """生成日报或周报的 TXT、图片，返回 TXT 路径"""
    # TXT
    print(f"\n📝 生成文本{report_type}...")
    txt_content = format_full_report(all_team, all_dash, all_analysis, run_time, report_type)
    txt_path = os.path.join(OUTPUT_DIR, f"统一{report_type}_{run_time.replace(':', '-').replace(' ', '_')}.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(txt_content)
    print(f"   TXT {report_type}已保存: {txt_path}")

    # 图片
    print(f"🖼️  生成{report_type}图片...")
    format_report_image(all_team, all_dash, all_analysis, run_time, report_type)

    return txt_path


def main():
    dry_run = "--dry-run" in sys.argv
    no_email = "--no-email" in sys.argv
    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 解析 --date, --old-date, --weekly 参数
    target_date_str = ""
    old_date_str = ""
    weekly_target = ""
    weekly_old = ""
    for i, arg in enumerate(sys.argv):
        if arg == "--date" and i + 1 < len(sys.argv):
            target_date_str = _parse_date_arg(sys.argv[i + 1])
        elif arg == "--old-date" and i + 1 < len(sys.argv):
            old_date_str = _parse_date_arg(sys.argv[i + 1])
        elif arg == "--weekly" and i + 1 < len(sys.argv):
            weekly_target = _parse_date_arg(sys.argv[i + 1])
        elif arg == "--weekly-old" and i + 1 < len(sys.argv):
            weekly_old = _parse_date_arg(sys.argv[i + 1])

    if target_date_str and not old_date_str:
        print(f"📅 手动指定日期: 新={target_date_str}, 旧=自动查找")
    elif target_date_str and old_date_str:
        print(f"📅 手动指定日期: 新={target_date_str}, 旧={old_date_str}")

    if weekly_target:
        if not weekly_old:
            parts = weekly_target.split("_")
            wt = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
            wo = wt - timedelta(days=7)
            weekly_old = f"{wo.year}_{wo.month}_{wo.day}"
            print(f"📅 周报: 新={weekly_target}, 旧={weekly_old}（自动前推7天）")
        else:
            print(f"📅 周报: 新={weekly_target}, 旧={weekly_old}")

    BW = 60
    print("═" * BW)
    print(vis_center("统一项目日报生成器", BW))
    print(vis_center(f"启动时间：{run_time}", BW))
    print("═" * BW)
    print()

    # ── 日报数据（日报始终使用独立数据，不受 --weekly 影响）──
    if not target_date_str:
        now = datetime.now()
        target_date_str = f"{now.year}_{now.month}_{now.day}"

    all_team, all_dash, all_analysis = _collect_report_data(target_date_str, old_date_str)

    if dry_run:
        print("\n[Dry-run 模式] 不保存文件，不发送邮件。")
        return

    # Excel 报表
    print("📊 生成 Excel 报表...")
    xlsx_paths = save_excel_report(all_team, all_dash, all_analysis, run_time)

    # 日报 TXT + 图片
    txt_path = _generate_reports(all_team, all_dash, all_analysis, run_time, "日报")

    # 进度追踪表
    print("\n📋 生成进度追踪表...")
    generate_tracking_files(all_team, all_dash, run_time)

    # ── 周报 ──
    if weekly_target:
        print("\n" + "═" * BW)
        # 重新收集周报数据
        weekly_all_team, weekly_all_dash, weekly_all_analysis = \
            _collect_report_data(weekly_target, weekly_old)
        _generate_reports(weekly_all_team, weekly_all_dash, weekly_all_analysis, run_time, "周报")

    # 发送邮件
    if not no_email:
        print("\n📧 发送邮件...")
        send_email(txt_path, xlsx_paths, run_time)
    else:
        print("\n[--no-email] 跳过邮件发送。")

    print()
    print("=" * 60)
    print("  全部处理完毕。")
    print("=" * 60)


# ============================================================================
# 进度追踪文件生成（格式匹配参考 Excel，直接可用）
# ============================================================================

EXCEL_EPOCH = datetime(1899, 12, 30)


def _date_str_to_serial(date_str: str) -> int:
    """ '2026_5_28' → Excel 日期序列号 """
    parts = date_str.split("_")
    d = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
    return (d - EXCEL_EPOCH).days


def _read_ref_xlsx(path: str) -> dict:
    """读取参考 Excel，返回 {sheet_name: {'title': A1值, 'headers': [...], 'rows': [[...], ...]}} """
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    result = {}
    try:
        with ZipFile(path) as z:
            ss = []
            if 'xl/sharedStrings.xml' in z.namelist():
                tree = ET.parse(z.open('xl/sharedStrings.xml'))
                for si in tree.findall('.//s:si', ns):
                    text = ''.join(t.text or '' for t in si.findall('.//s:t', ns))
                    ss.append(text)
            wb_tree = ET.parse(z.open('xl/workbook.xml'))
            sheets = [(s.get('name'), s.get('sheetId')) for s in wb_tree.findall('.//s:sheet', ns)]
            for name, sid in sheets:
                sheet_tree = ET.parse(z.open(f'xl/worksheets/sheet{sid}.xml'))
                all_rows = []
                for row in sheet_tree.findall('.//s:row', ns):
                    cells = {}
                    for c in row.findall('s:c', ns):
                        v = c.find('s:v', ns)
                        val = v.text if v is not None else ''
                        if c.get('t') == 's' and val:
                            val = ss[int(val)] if int(val) < len(ss) else val
                        cells[c.get('r')] = val
                    row_vals = [cells.get(k, '') for k in sorted(cells.keys(), key=lambda x: (len(''.join(filter(str.isalpha, x))), x))]
                    all_rows.append(row_vals)
                title = ""
                headers = []
                data_rows = []
                if all_rows:
                    # 检测表头行：第一行有多个非空单元格且含中文 → 是表头
                    # 第一行只有一个非空单元格（如"供应商Alpha"）→ 是标题，第二行为表头或数据
                    row0 = all_rows[0]
                    non_empty_0 = sum(1 for v in row0 if v and str(v).strip())
                    first_cell = str(row0[0]) if row0 else ""
                    if non_empty_0 >= 2 and re.search(r'[一-鿿]', first_cell):
                        # 多列表头行（如"作业日期, 序号, ..."）
                        headers = row0
                        data_rows = all_rows[1:]
                    elif non_empty_0 <= 1 and len(all_rows) > 1:
                        # 单单元格 → 可能是标题，第二行作为表头
                        row1 = all_rows[1]
                        non_empty_1 = sum(1 for v in row1 if v and str(v).strip())
                        if non_empty_1 >= 2:
                            headers = row1
                            data_rows = all_rows[2:]
                        else:
                            data_rows = all_rows[1:]  # 没有表头，第二行起都是数据
                    else:
                        data_rows = all_rows
                result[name] = {"title": title, "headers": headers, "rows": data_rows}
    except Exception as e:
        print(f"   [警告] 读取参考文件失败 {path}: {e}")
    return result


def generate_tracking_files(all_team, all_dash, run_time: str) -> list:
    """
    生成可直接使用的进度追踪 Excel，格式匹配参考文件。
    进度列 → 当天历史全量数据，质量列 → 当天实时数据。
    """
    ref_files = [
        (os.path.join(ROOT_DIR, "项目A_DATA验收进度.xlsx"), "项目A_DATA",
         ["供应商Alpha", "供应商Beta", "供应商Gamma"]),
        (os.path.join(ROOT_DIR, "项目B_ANOMALY验收进度.xlsx"), "项目B_ANOMALY",
         ["供应商Delta", "供应商Epsilon", "供应商Zeta", "供应商Eta"]),
    ]
    paths = []
    today_str = datetime.now().strftime("%Y_%m_%d")
    # 目录名 → 参考文件 Sheet 名映射
    DIR_TO_REF = {}

    for ref_path, project_name, proj_suppliers in ref_files:
        if not os.path.exists(ref_path):
            print(f"   [跳过] 参考文件不存在: {ref_path}")
            continue

        ref_data = _read_ref_xlsx(ref_path)
        if not ref_data:
            continue

        out_path = os.path.join(ROOT_DIR, f"{project_name}验收进度.xlsx")

        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            wb = writer.book
            header_fmt = wb.add_format({"bold": True, "align": "center", "bg_color": "#D9E1F2"})
            normal_fmt = wb.add_format({"align": "center"})

            proj_dash_data = []  # 收集各供应商大盘数据用于汇总
            proj_hr_data = {}    # 收集各供应商人力数据用于人效计算
            all_hr_rows = []
            hr_headers = []
            for supplier in proj_suppliers:
                sdir = SUPPLIER_DIR_MAP.get(supplier, supplier)
                team_dir = os.path.join(ROOT_DIR, project_name, sdir, "团队")
                dash_dir = os.path.join(ROOT_DIR, project_name, sdir, "整体")

                # ── 供应商 sheet ──
                dash_sheet = DIR_TO_REF.get(supplier, supplier)
                dash_old = ref_data.get(dash_sheet, {"headers": [], "rows": []})
                dash_headers = [h for h in dash_old["headers"] if h and "累积折损" not in str(h)]
                if not dash_headers:
                    dash_headers = ["日期", "已下发量", "已标注量", "待标注", "质检中", "验收中", "已交付", "已完成",
                                    "质检通过率", "首次验收通过率", "已标注增量"]
                trim_idx = next((i + 1 for i, h in enumerate(dash_headers) if "已标注增量" in str(h)), 0)
                if trim_idx:
                    dash_headers = dash_headers[:trim_idx]

                # 从所有可用源文件全量重建大盘数据，只保留参考中无可再生源文件的旧行
                pg_files = _find_files_after(dash_dir, "项目大盘进度与质量统计表", "历史全量", EXCEL_EPOCH)
                ql_files = _find_files_after(dash_dir, "项目大盘进度与质量统计表", "当天实时", EXCEL_EPOCH)
                ql_by_date = {d: p for d, p in ql_files}
                # 可用源文件的最早日期
                earliest_src_date = pg_files[0][0] if pg_files else datetime.now()
                # 只保留早于最早源文件的旧行（无可再生源文件的历史数据）
                all_dash_rows = []
                for r in dash_old["rows"]:
                    rr = r[:len(dash_headers)]
                    d_val = _safe_int(rr[0]) if len(rr) > 0 else 0
                    if d_val > 0:
                        rd = EXCEL_EPOCH + timedelta(days=d_val)
                        if rd < earliest_src_date:
                            all_dash_rows.append(rr)
                # 修正可能的年份偏移错误
                for rr in all_dash_rows:
                    d_val = _safe_int(rr[0]) if len(rr) > 0 else 0
                    if d_val > 0 and d_val < 46000:
                        rr[0] = str(d_val + 365)
                # 按日期排序
                all_dash_rows.sort(key=lambda r: _safe_int(r[0]) if len(r) > 0 else 0)
                # 统一补齐所有行到完整列数
                ncols_d = len(dash_headers)
                for rr in all_dash_rows:
                    while len(rr) < ncols_d:
                        rr.append("")
                for date_obj, pg_path in pg_files:
                    serial = _date_str_to_serial(date_obj.strftime("%Y_%m_%d"))
                    pg = _read_dash_by_path(pg_path)[0]  # 进度：历史全量
                    ql_dict = {}
                    if date_obj in ql_by_date:
                        _, ql_dict, _ = _read_dash_by_path(ql_by_date[date_obj])  # 质量：当天实时
                    ql_rows = []
                    for team, v in ql_dict.items():
                        qc_r = round(v["质检_pass"] / v["质检_total"] * 100, 1) if v["质检_total"] > 0 else None
                        ac_r = round(v["验收_pass"] / v["验收_total"] * 100, 1) if v["验收_total"] > 0 else None
                        ql_rows.append({"团队": team, "质检通过率": qc_r, "首次验收通过率": ac_r,
                                        "质检_pass": v["质检_pass"], "质检_total": v["质检_total"],
                                        "验收_pass": v["验收_pass"], "验收_total": v["验收_total"]})
                    ql_df = pd.DataFrame(ql_rows)
                    new_row = _build_dash_tracking_row(pg, ql_df, serial, all_dash_rows, dash_headers)
                    if new_row:
                        new_row = new_row[:len(dash_headers)]
                        all_dash_rows.append(new_row)
                # 重算全部行的已标注增量
                inc_col = _find_header_index(dash_headers, "已标注增量")
                done_col = _find_header_index(dash_headers, "已标注量")
                for i in range(len(all_dash_rows)):
                    r = all_dash_rows[i]
                    cur = _safe_int(r[done_col]) if len(r) > done_col else 0
                    prev = 0
                    if i > 0:
                        pr = all_dash_rows[i - 1]
                        prev = _safe_int(pr[done_col]) if len(pr) > done_col else 0
                    if len(r) > inc_col:
                        r[inc_col] = str(cur - prev)
                _append_and_write_sheet(writer, dash_sheet, dash_headers,
                                        all_dash_rows, [], header_fmt, normal_fmt)

                # 收集大盘数据用于汇总（在 dashboard if 块内，确保所有供应商都收集到）
                proj_dash_data.append((supplier, dash_headers, all_dash_rows))
                # ── 人力 sheet ──
                hr_sheet = f"{DIR_TO_REF.get(supplier, supplier)}人力"
                hr_old = ref_data.get(hr_sheet, {"headers": [], "rows": []})
                hr_headers = [h for h in hr_old["headers"] if h]
                if not hr_headers:
                    hr_headers = ["作业日期", "序号", "标注员_英文名", "标注员_中文名", "认领总量", "已标注量",
                                  "质检通过量", "质检驳回量", "质检总量", "验收通过量", "验收驳回量", "验收总量",
                                  "质检通过率", "首次验收通过率", "累积折损验收通过率", "已标注增量",
                                  "供应商", "项目名称", "上班卡时间", "下班卡时间"]
                cn_idx = next((i for i, h in enumerate(hr_headers) if "中文名" in str(h)), 2)
                en_idx = next((i for i, h in enumerate(hr_headers) if "英文名" in str(h)), 3)
                name_map = _build_name_map(hr_old["rows"], cn_idx, en_idx)
                if project_name == "项目A_DATA":
                    trim_idx = next((i + 1 for i, h in enumerate(hr_headers) if "已标注增量" in str(h)), len(hr_headers))
                else:
                    trim_idx = next((i for i, h in enumerate(hr_headers) if "供应商" in str(h)), len(hr_headers))
                hr_headers = hr_headers[:trim_idx]

                # 从所有可用源文件全量重建人力数据，只保留参考中无可再生源文件的旧行
                pg_files = _find_files_after(team_dir, "标注团队绩效明细表", "历史全量", EXCEL_EPOCH)
                ql_files = _find_files_after(team_dir, "标注团队绩效明细表", "当天实时", EXCEL_EPOCH)
                ql_by_date = {d: p for d, p in ql_files}
                # 可用源文件的最早日期
                earliest_src_date = pg_files[0][0] if pg_files else datetime.now()
                # 只保留早于最早源文件的旧行（无可再生源文件的历史数据）
                old_hr_cutoff = earliest_src_date - timedelta(days=1)
                all_hr_rows = []
                for r in hr_old["rows"]:
                    rr = r[:len(hr_headers)]
                    d_val = _safe_int(rr[0]) if len(rr) > 0 else 0
                    if d_val > 0:
                        # 修正历史数据中可能的年份错误（如 2025 误写为 2026 等，差 365 天视为年份偏移）
                        rd = EXCEL_EPOCH + timedelta(days=d_val)
                        if rd < earliest_src_date:
                            all_hr_rows.append(rr)
                # 修正历史数据中可能的年份偏移错误（serial 差 365 天 → 加回一年）
                for rr in all_hr_rows:
                    d_val = _safe_int(rr[0]) if len(rr) > 0 else 0
                    if d_val > 0 and d_val < 46000:  # 2026年之前 → 可能是年份错误
                        rr[0] = str(d_val + 365)
                # 按 (日期, 序号) 排序，保持与源文件一致的排列顺序
                def _hr_sort_key(r):
                    dv = _safe_int(r[0]) if len(r) > 0 else 0
                    sq = _safe_int(r[1]) if len(r) > 1 else 0
                    return (dv, sq)
                all_hr_rows.sort(key=_hr_sort_key)
                # 统一补齐所有行到完整的列数，避免历史行缺列导致增量计算跳过
                ncols = len(hr_headers)
                for rr in all_hr_rows:
                    while len(rr) < ncols:
                        rr.append("")
                # 从源文件重建所有行
                for date_obj, pg_path in pg_files:
                    serial = _date_str_to_serial(date_obj.strftime("%Y_%m_%d"))
                    progress = _read_team_raw_by_path(pg_path, file_date=date_obj)
                    quality = _read_team_raw_by_path(ql_by_date[date_obj], file_date=date_obj) if date_obj in ql_by_date else pd.DataFrame()
                    new_rows = _build_hr_tracking_rows(
                        progress, quality, serial, hr_headers, name_map,
                        project_name, all_hr_rows, supplier
                    )
                    if new_rows:
                        new_rows = [r[:len(hr_headers)] for r in new_rows]
                        all_hr_rows.extend(new_rows)
                # 重算全部人力行的已标注增量（按标注员匹配上期）
                hr_inc_col = _find_header_index(hr_headers, "已标注增量")
                hr_done_col = _find_header_index(hr_headers, "已标注量")
                hr_en_col = _find_header_index(hr_headers, "英文名")
                # 按标注员分组，每组内按期排序计算增量
                annotator_prev = {}
                for i in range(len(all_hr_rows)):
                    r = all_hr_rows[i]
                    en = str(r[hr_en_col]).strip() if len(r) > hr_en_col else ""
                    cur = _safe_int(r[hr_done_col]) if len(r) > hr_done_col else 0
                    prev = annotator_prev.get(en, 0)
                    if len(r) > hr_inc_col:
                        r[hr_inc_col] = str(cur - prev)
                    annotator_prev[en] = cur
                _append_and_write_sheet(writer, hr_sheet, hr_headers,
                                        all_hr_rows, [], header_fmt, normal_fmt)

                # 收集人力数据用于汇总的人效计算
                proj_hr_data[supplier] = (all_hr_rows, hr_headers)

                # ── 周报 sheet（周五展示日）──
                weekly_headers, weekly_rows = _build_weekly_sheet(
                    all_dash_rows, dash_headers, all_hr_rows, hr_headers)
                if weekly_rows:
                    _append_and_write_sheet(writer, f"{supplier}周报", weekly_headers,
                                            weekly_rows, [], header_fmt, normal_fmt)

                # ── 人员质量 sheet ──
                qual_headers, qual_rows = _build_personnel_quality_sheet(all_hr_rows, hr_headers)
                if qual_rows:
                    _append_and_write_sheet(writer, f"{supplier}人员质量", qual_headers,
                                            qual_rows, [], header_fmt, normal_fmt)

            # ── 汇总 sheet ──
            summary_headers, summary_rows = _build_summary_sheet(proj_dash_data)
            if summary_rows:
                _append_and_write_sheet(writer, "汇总", summary_headers,
                                        summary_rows, [], header_fmt, normal_fmt)

            # ── 质量效率汇总明细 sheet ──
            qe_headers, qe_rows = _build_quality_efficiency_sheet(proj_dash_data, proj_hr_data)
            if qe_rows:
                _append_and_write_sheet(writer, "质量效率汇总明细", qe_headers,
                                        qe_rows, [], header_fmt, normal_fmt)

            # ── 保留参考文件中未处理的供应商数据 ──
            handled = set()
            for s in proj_suppliers:
                rn = DIR_TO_REF.get(s, s)
                handled.update([rn, f"{rn}人力", f"{rn}周报", f"{rn}人员质量"])
            handled.update(["腾游", "腾游人力", "汇总", "质量效率汇总明细"])  # 旧名称，不再使用
            for sname, sdata in ref_data.items():
                if "明细" in sname:
                    continue
                if sname not in handled:
                    # 角色项目人力：裁掉供应商~下班卡
                    if project_name != "项目A_DATA" and "人力" in sname:
                        hdrs = [h for h in sdata["headers"] if h]
                        trim = next((i for i, h in enumerate(hdrs) if "供应商" in str(h)), len(hdrs))
                        sdata["headers"] = hdrs[:trim]
                        sdata["rows"] = [r[:len(sdata["headers"])] for r in sdata["rows"]]
                    _append_and_write_sheet(writer, sname, sdata["headers"],
                                            sdata["rows"], [], header_fmt, normal_fmt)

        paths.append(out_path)
        print(f"   [{project_name}] 进度追踪表已保存: {out_path}")

    return paths


# ── 读取当天原始数据 ──

def _find_latest_file(directory: str, base_name: str, ftype: str) -> str | None:
    """找目录下最新的匹配文件"""
    if not os.path.isdir(directory):
        return None
    pattern = re.compile(re.escape(base_name) + r"_" + re.escape(ftype) + r"_(\d{4})_(\d{1,2})_(\d{1,2})\.xls")
    files = []
    for fname in os.listdir(directory):
        m = pattern.search(fname)
        if m:
            d = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            files.append((d, os.path.join(directory, fname)))
    files.sort(key=lambda x: x[0])
    return files[-1][1] if files else None


def _find_files_after(directory: str, base_name: str, ftype: str, after_date: datetime) -> list:
    """找目录下所有日期 > after_date 的文件，按日期排序返回 [(date, path), ...]"""
    if not os.path.isdir(directory):
        return []
    pattern = re.compile(re.escape(base_name) + r"_" + re.escape(ftype) + r"_(\d{4})_(\d{1,2})_(\d{1,2})\.xls")
    files = []
    for fname in os.listdir(directory):
        m = pattern.search(fname)
        if m:
            d = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if d >= after_date:
                files.append((d, os.path.join(directory, fname)))
    files.sort(key=lambda x: x[0])
    return files


def _read_team_raw_by_path(path: str, file_date=None) -> pd.DataFrame:
    """读取单个团队绩效明细表，file_date 用于过滤已离职标注员"""
    tables = pd.read_html(path)
    df = tables[0].copy()
    df = df[df["标注员"] != "未分配"].copy()
    df["标注员"] = df["标注员"].astype(str).str.strip()
    for col in ["总量", "已标注", "质检中", "验收中", "已驳回", "已完成"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    for src_col, pct_col, pass_col, total_col in [
        ("🛡️质检通过率", "质检通过率_pct", "质检_pass", "质检_total"),
        ("🎯验收(首次验收通过)", "首次验收通过率_pct", "首次验收_pass", "首次验收_total"),
        ("🎯验收(累积折损通过)", "累积验收通过率_pct", "累积验收_pass", "累积验收_total"),
    ]:
        if src_col in df.columns:
            df[pct_col] = df[src_col].apply(extract_pct)
            parsed = df[src_col].apply(parse_team_quality_cell)
            df[pass_col] = parsed.apply(lambda x: x[0])
            df[total_col] = parsed.apply(lambda x: x[1])
    # 过滤已离职标注员
    if file_date and EXCLUDED_ANNOTATORS:
        fd = file_date.date() if hasattr(file_date, 'date') else file_date
        for name, cutoff in EXCLUDED_ANNOTATORS.items():
            if fd >= cutoff:
                df = df[df["标注员"] != name]
    return df


def _read_dash_by_path(path: str):
    """读取单个项目大盘文件，返回 (team_df, quality_by_team, batch_detail)"""
    tables = pd.read_html(path)
    team_df = tables[1].copy()
    team_df["归属团队名称"] = team_df["归属团队名称"].astype(str).str.strip()
    team_df = team_df[team_df["归属团队名称"] != "内部团队（已脱敏）"].copy()
    for col in ["总量", "已标注", "质检中", "验收中", "已驳回", "已完成"]:
        team_df[col] = pd.to_numeric(team_df[col], errors="coerce").fillna(0).astype(int)
    batch_df = tables[0].copy()
    batch_name_col = "批次名称/标注员" if "批次名称/标注员" in batch_df.columns else "批次名称"
    batch_df["团队"] = batch_df[batch_name_col].astype(str).str.extract(r"\[([^\]]+)\]").fillna("未知团队")
    qc_p = batch_df["🛡️首次质检"].apply(parse_quality_cell)
    ac_p = batch_df["🎯首次验收"].apply(parse_quality_cell)
    batch_df["质检_pass"] = qc_p.apply(lambda x: x[0])
    batch_df["质检_total"] = qc_p.apply(lambda x: x[1])
    batch_df["验收_pass"] = ac_p.apply(lambda x: x[0])
    batch_df["验收_total"] = ac_p.apply(lambda x: x[1])
    quality_by_team = {}
    for team, grp in batch_df.groupby("团队"):
        if "未分配" in team:
            continue
        quality_by_team[team] = {
            "质检_pass": int(grp["质检_pass"].sum()),
            "质检_total": int(grp["质检_total"].sum()),
            "验收_pass": int(grp["验收_pass"].sum()),
            "验收_total": int(grp["验收_total"].sum()),
        }
    detail_cols = [batch_name_col, "团队", "质检_pass", "质检_total", "验收_pass", "验收_total"]
    batch_detail = batch_df[[c for c in detail_cols if c in batch_df.columns]].copy()
    batch_detail = batch_detail.rename(columns={batch_name_col: "批次名称"})
    return team_df, quality_by_team, batch_detail


def _get_last_ref_date(old_rows: list, date_col_idx: int = 0) -> datetime | None:
    """从参考文件行中提取最后一个有效日期"""
    for r in reversed(old_rows):
        if len(r) > date_col_idx and r[date_col_idx]:
            try:
                serial = int(float(str(r[date_col_idx])))
                if serial > 40000:  # 合理的 Excel 日期
                    return EXCEL_EPOCH + timedelta(days=serial)
            except (ValueError, TypeError):
                continue
    return None


def _read_latest_team_raw(team_dir: str, ftype: str) -> pd.DataFrame:
    """读取最新团队绩效明细表原始数据（标注员级），含通过数/总量"""
    path = _find_latest_file(team_dir, "标注团队绩效明细表", ftype)
    if not path:
        return pd.DataFrame()
    tables = pd.read_html(path)
    df = tables[0].copy()
    df = df[df["标注员"] != "未分配"].copy()
    df["标注员"] = df["标注员"].astype(str).str.strip()
    for col in ["总量", "已标注", "质检中", "验收中", "已驳回", "已完成"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    for src_col, pct_col, pass_col, total_col in [
        ("🛡️质检通过率", "质检通过率_pct", "质检_pass", "质检_total"),
        ("🎯验收(首次验收通过)", "首次验收通过率_pct", "首次验收_pass", "首次验收_total"),
        ("🎯验收(累积折损通过)", "累积验收通过率_pct", "累积验收_pass", "累积验收_total"),
    ]:
        if src_col in df.columns:
            df[pct_col] = df[src_col].apply(extract_pct)
            parsed = df[src_col].apply(parse_team_quality_cell)
            df[pass_col] = parsed.apply(lambda x: x[0])
            df[total_col] = parsed.apply(lambda x: x[1])
    return df


def _read_latest_dash_progress(dash_dir: str, ftype: str) -> pd.DataFrame:
    """读取最新项目大盘 team 进度表（Table1）"""
    path = _find_latest_file(dash_dir, "项目大盘进度与质量统计表", ftype)
    if not path:
        return pd.DataFrame()
    tables = pd.read_html(path)
    df = tables[1].copy()
    df["归属团队名称"] = df["归属团队名称"].astype(str).str.strip()
    df = df[df["归属团队名称"] != "内部团队（已脱敏）"].copy()
    for col in ["总量", "已标注", "质检中", "验收中", "已驳回", "已完成"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    return df


def _read_latest_dash_quality(dash_dir: str, ftype: str) -> pd.DataFrame:
    """读取最新项目大盘批次明细（Table0），聚合团队质量"""
    path = _find_latest_file(dash_dir, "项目大盘进度与质量统计表", ftype)
    if not path:
        return pd.DataFrame()
    tables = pd.read_html(path)
    batch_df = tables[0].copy()
    batch_name_col = "批次名称/标注员" if "批次名称/标注员" in batch_df.columns else "批次名称"
    batch_df["团队"] = batch_df[batch_name_col].astype(str).str.extract(r"\[([^\]]+)\]").fillna("未知团队")
    qc_p = batch_df["🛡️首次质检"].apply(parse_quality_cell)
    ac_p = batch_df["🎯首次验收"].apply(parse_quality_cell)
    batch_df["质检_pass"] = qc_p.apply(lambda x: x[0])
    batch_df["质检_total"] = qc_p.apply(lambda x: x[1])
    batch_df["验收_pass"] = ac_p.apply(lambda x: x[0])
    batch_df["验收_total"] = ac_p.apply(lambda x: x[1])
    # 团队聚合
    rows = []
    for team, grp in batch_df.groupby("团队"):
        if "未分配" in team:
            continue
        qc_pass = int(grp["质检_pass"].sum())
        qc_tot = int(grp["质检_total"].sum())
        ac_pass = int(grp["验收_pass"].sum())
        ac_tot = int(grp["验收_total"].sum())
        rows.append({
            "团队": team,
            "质检通过率": round(qc_pass / qc_tot * 100, 1) if qc_tot > 0 else None,
            "首次验收通过率": round(ac_pass / ac_tot * 100, 1) if ac_tot > 0 else None,
            "质检_pass": qc_pass, "质检_total": qc_tot,
            "验收_pass": ac_pass, "验收_total": ac_tot,
            "_batch_df": grp,
        })
    return pd.DataFrame(rows)


# ── 辅助：名称映射 & 通过率格式化 ──

def _safe_int(val) -> int:
    """安全转整数，处理 '-' / 'nan' / 空等非数字"""
    s = str(val).strip() if val else ""
    if not s or s in ("-", "nan", "None"):
        return 0
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def _find_header_index(headers: list, keyword: str) -> int:
    for i, h in enumerate(headers):
        if keyword in str(h):
            return i
    return -1


def _fmt_rate_val(pct_val, pass_cnt=None, total_cnt=None) -> str:
    """百分比数值 → '81.9%(对186/阅227)' 格式"""
    if pct_val is None or not pd.notna(pct_val):
        return "-"
    pct_str = f"{pct_val}%"
    if pass_cnt is not None and total_cnt is not None and total_cnt > 0:
        return f"{pct_val}%(对{int(pass_cnt)}/阅{int(total_cnt)})"
    return pct_str


def _build_name_map(ref_rows: list, cn_col: int, en_col: int) -> dict:
    """从参考文件人力数据中提取 {英文名: 中文名} 映射"""
    name_map = {}
    for r in ref_rows:
        if len(r) > max(cn_col, en_col):
            cn = str(r[cn_col]).strip() if cn_col < len(r) else ""
            en = str(r[en_col]).strip() if en_col < len(r) else ""
            if en and cn and cn != "-":
                name_map[en] = cn
    return name_map


# ── 构建追踪行 ──

def _build_dash_tracking_row(pg: pd.DataFrame, ql: pd.DataFrame, serial: int,
                              old_rows: list, headers: list) -> list | None:
    """构建供应商 sheet 的一行今日数据"""
    if pg.empty:
        return None
    for _, prow in pg.iterrows():
        team = prow["归属团队名称"]
        qrow = ql[ql["团队"] == team] if not ql.empty else pd.DataFrame()
        total_n = int(prow["总量"])
        done_n = int(prow["已标注"])
        qc_n = int(prow["质检中"])
        ac_n = int(prow["验收中"])
        row_data = {"日期": str(serial),
                    "已下发量": str(total_n),
                    "已标注量": str(done_n),
                    "质检中": str(qc_n),
                    "验收中": str(ac_n),
                    "已驳回": str(int(prow["已驳回"])),
                    "已完成": str(int(prow["已完成"]))}
        # 公式列
        row_data["待标注"] = str(total_n - done_n)
        row_data["已交付"] = str(qc_n + ac_n)
        # 已标注增量 = 本期 - 上期
        prev_done = 0
        ridx = _find_header_index(headers, "已标注量")
        for old_r in reversed(old_rows):
            if len(old_r) > ridx and old_r[ridx]:
                try:
                    prev_done = int(float(str(old_r[ridx])))
                except (ValueError, TypeError):
                    pass
                break
        row_data["已标注增量"] = str(done_n - prev_done)
        # 通过率（带明细）
        if len(qrow) > 0:
            qr = qrow.iloc[0]
            row_data["质检通过率"] = _fmt_rate_val(qr.get("质检通过率"), qr.get("质检_pass"), qr.get("质检_total"))
            row_data["首次验收通过率"] = _fmt_rate_val(qr.get("首次验收通过率"), qr.get("验收_pass"), qr.get("验收_total"))
        row_data.setdefault("质检通过率", "-")
        row_data.setdefault("首次验收通过率", "-")
        return _row_dict_to_list(row_data, headers)
    return None


def _build_weekly_sheet(dash_rows: list, dash_headers: list, hr_rows: list = None,
                         hr_headers: list = None) -> tuple:
    """从大盘数据生成周报：以周五为展示日，计算每周已完成增量和已标注增量 + 产效"""
    date_idx = _find_header_index(dash_headers, "日期")
    done_idx = _find_header_index(dash_headers, "已完成")
    labeled_idx = _find_header_index(dash_headers, "已标注量")
    if date_idx < 0 or done_idx < 0:
        return [], []

    # 收集每日数据
    daily = []
    for r in dash_rows:
        d_val = _safe_int(r[date_idx]) if len(r) > date_idx else 0
        if d_val <= 0:
            continue
        dt = EXCEL_EPOCH + timedelta(days=d_val)
        done = _safe_int(r[done_idx]) if len(r) > done_idx else 0
        labeled = _safe_int(r[labeled_idx]) if len(r) > labeled_idx and labeled_idx >= 0 else 0
        daily.append((dt, done, labeled))

    if not daily:
        return [], []

    daily.sort(key=lambda x: x[0])

    weeks = {}
    for dt, done, labeled in daily:
        iso_year, iso_week, _ = dt.isocalendar()
        week_key = (iso_year, iso_week)
        if week_key not in weeks:
            weeks[week_key] = []
        weeks[week_key].append((dt, done, labeled))

    headers = ["周次", "周五日期", "本周已完成", "累计已完成", "周已完成增量",
               "本周已标注", "累计已标注", "周标注增量",
               "投入人数", "人均标注/天"]
    rows = []
    prev_cum_done = None
    prev_cum_labeled = None

    for (y, w), entries in sorted(weeks.items()):
        entries.sort(key=lambda x: x[0])
        fri_data = None
        for dt, done, labeled in entries:
            if dt.weekday() == 4:
                fri_data = (dt, done, labeled)
        if fri_data is None:
            continue  # 非周五不统计

        fri_date, cum_done, cum_labeled = fri_data
        mon_data = entries[0]
        _, week_start_done, week_start_labeled = mon_data
        week_done = cum_done - week_start_done
        week_labeled = cum_labeled - week_start_labeled

        inc_done = cum_done - prev_cum_done if prev_cum_done is not None else 0
        inc_labeled = cum_labeled - prev_cum_labeled if prev_cum_labeled is not None else 0
        prev_cum_done = cum_done
        prev_cum_labeled = cum_labeled

        # 本周投入人数：本周内有标注记录的标注员数
        week_persons = 0
        if hr_rows and hr_headers:
            hr_date_idx = _find_header_index(hr_headers, "作业日期")
            hr_done_idx = _find_header_index(hr_headers, "已标注量")
            hr_en_idx = _find_header_index(hr_headers, "英文名")
            if hr_date_idx >= 0 and hr_done_idx >= 0 and hr_en_idx >= 0:
                persons = set()
                for r in hr_rows:
                    d_val = _safe_int(r[hr_date_idx]) if len(r) > hr_date_idx else 0
                    if d_val <= 0:
                        continue
                    rd = EXCEL_EPOCH + timedelta(days=d_val)
                    if rd.isocalendar()[:2] == (y, w):
                        done = _safe_int(r[hr_done_idx]) if len(r) > hr_done_idx else 0
                        if done > 0:
                            en = str(r[hr_en_idx]).strip() if len(r) > hr_en_idx else ""
                            if en and en != "-":
                                persons.add(en)
                week_persons = len(persons)
        per_day = round(week_labeled / week_persons / 5, 1) if week_persons > 0 and week_labeled > 0 else 0

        rows.append([
            f"W{y - 2000:02d}{w:02d}",
            fri_date.strftime("%Y-%m-%d"),
            str(week_done),
            str(cum_done),
            str(inc_done),
            str(week_labeled),
            str(cum_labeled),
            str(inc_labeled),
            str(week_persons),
            str(per_day),
        ])

    return headers, rows


def _build_personnel_quality_sheet(hr_rows: list, hr_headers: list) -> tuple:
    """从人力数据生成人员质量周报：每人每周的质量数据，以周五为展示日"""
    en_idx = _find_header_index(hr_headers, "英文名")
    cn_idx = _find_header_index(hr_headers, "中文名")
    date_idx = _find_header_index(hr_headers, "作业日期")
    done_idx = _find_header_index(hr_headers, "已标注量")
    comp_idx = _find_header_index(hr_headers, "已完成")
    qc_rate_idx = _find_header_index(hr_headers, "质检通过率")
    ac_rate_idx = _find_header_index(hr_headers, "首次验收通过率")
    ac_cum_rate_idx = _find_header_index(hr_headers, "累积折损验收通过率")
    qc_pass_idx = _find_header_index(hr_headers, "质检通过量")
    qc_total_idx = _find_header_index(hr_headers, "质检总量")
    ac_pass_idx = _find_header_index(hr_headers, "验收通过量")
    ac_total_idx = _find_header_index(hr_headers, "验收总量")
    has_count_cols = qc_pass_idx >= 0 and qc_total_idx >= 0

    # 按人、按周分组，取每周五（或最近一天）的数据
    person_weekly = defaultdict(lambda: defaultdict(list))
    name_map = {}
    for r in hr_rows:
        d_val = _safe_int(r[date_idx]) if len(r) > date_idx and date_idx >= 0 else 0
        if d_val <= 0:
            continue
        en = str(r[en_idx]).strip() if len(r) > en_idx and en_idx >= 0 else ""
        if not en or en == "-":
            continue
        dt = EXCEL_EPOCH + timedelta(days=d_val)
        cn = str(r[cn_idx]).strip() if len(r) > cn_idx and cn_idx >= 0 and r[cn_idx] and str(r[cn_idx]) != "-" else ""
        if en and cn:
            name_map[en] = cn
        iso_year, iso_week, _ = dt.isocalendar()
        person_weekly[en][(iso_year, iso_week)].append((dt, r))

    if not person_weekly:
        return [], []

    # 收集所有周次
    all_weeks = set()
    for pdata in person_weekly.values():
        all_weeks.update(pdata.keys())
    all_weeks = sorted(all_weeks)

    # 提取计数的辅助函数
    def _extract_count(rate_str):
        m = re.search(r'\([^\d]*(\d+)/[^\d]*(\d+)\)', str(rate_str))
        return (m.group(1), m.group(2)) if m else ("-", "-")

    headers = ["周次", "周五日期", "标注员_英文名", "标注员_中文名",
               "本周已标注", "累计已标注量", "已完成量",
               "质检通过量", "质检总量", "质检通过率",
               "验收通过量", "验收总量", "首次验收通过率",
               "累积折损验收通过率"]
    rows = []

    for en in sorted(person_weekly.keys()):
        cn = name_map.get(en, "-")
        pdata = person_weekly[en]
        for week_key in all_weeks:
            if week_key not in pdata:
                continue
            entries = sorted(pdata[week_key], key=lambda x: x[0])
            # 只取周五数据，非周五跳过
            fri = None
            for dt, r in entries:
                if dt.weekday() == 4:
                    fri = (dt, r)
            if fri is None:
                continue

            fri_date, r = fri
            # 本周已标注 = 本周五累计 - 本周一累计
            _, mon_r = entries[0]
            week_labeled = (_safe_int(r[done_idx]) if len(r) > done_idx and done_idx >= 0 else 0) - \
                           (_safe_int(mon_r[done_idx]) if len(mon_r) > done_idx and done_idx >= 0 else 0)
            cum_done = str(r[done_idx]) if len(r) > done_idx and done_idx >= 0 else "0"
            comp = str(r[comp_idx]) if len(r) > comp_idx and comp_idx >= 0 else "0"

            if has_count_cols:
                qc_p = str(r[qc_pass_idx]) if len(r) > qc_pass_idx and r[qc_pass_idx] else "0"
                qc_t = str(r[qc_total_idx]) if len(r) > qc_total_idx and r[qc_total_idx] else "0"
                ac_p = str(r[ac_pass_idx]) if len(r) > ac_pass_idx and r[ac_pass_idx] else "0"
                ac_t = str(r[ac_total_idx]) if len(r) > ac_total_idx and r[ac_total_idx] else "0"
            else:
                qc_rate_str = str(r[qc_rate_idx]) if len(r) > qc_rate_idx and qc_rate_idx >= 0 else ""
                ac_rate_str = str(r[ac_rate_idx]) if len(r) > ac_rate_idx and ac_rate_idx >= 0 else ""
                qc_p, qc_t = _extract_count(qc_rate_str)
                ac_p, ac_t = _extract_count(ac_rate_str)

            qc_r = str(r[qc_rate_idx]) if len(r) > qc_rate_idx and qc_rate_idx >= 0 else "-"
            ac_r = str(r[ac_rate_idx]) if len(r) > ac_rate_idx and ac_rate_idx >= 0 else "-"
            ac_cr = str(r[ac_cum_rate_idx]) if len(r) > ac_cum_rate_idx and ac_cum_rate_idx >= 0 else "-"

            y, w = week_key
            rows.append([
                f"W{y - 2000:02d}{w:02d}",
                fri_date.strftime("%Y-%m-%d"),
                en, cn,
                str(week_labeled), cum_done, comp,
                qc_p, qc_t, qc_r,
                ac_p, ac_t, ac_r, ac_cr,
            ])

    return headers, rows


def _build_summary_sheet(proj_dash_data: list) -> tuple:
    """跨供应商汇总：每日和每周（周五）的已标注量、已完成量"""
    if not proj_dash_data:
        return [], []

    suppliers = []
    all_dates = set()
    supplier_daily = {}

    for supplier, dash_headers, dash_rows in proj_dash_data:
        suppliers.append(supplier)
        date_idx = _find_header_index(dash_headers, "日期")
        labeled_idx = _find_header_index(dash_headers, "已标注量")
        done_idx = _find_header_index(dash_headers, "已完成")
        if date_idx < 0 or labeled_idx < 0 or done_idx < 0:
            continue

        daily = {}
        for r in dash_rows:
            d_val = _safe_int(r[date_idx]) if len(r) > date_idx else 0
            if d_val <= 0:
                continue
            dt = EXCEL_EPOCH + timedelta(days=d_val)
            labeled = _safe_int(r[labeled_idx]) if len(r) > labeled_idx else 0
            done = _safe_int(r[done_idx]) if len(r) > done_idx else 0
            daily[dt] = (labeled, done)
            all_dates.add(dt)
        supplier_daily[supplier] = daily

    if not all_dates:
        return [], []

    all_dates = sorted(all_dates)

    headers = ["日期/周次", "类型"]
    for s in suppliers:
        headers.append(f"{s}_已标注量")
        headers.append(f"{s}_已完成量")
    headers.append("已标注汇总")
    headers.append("已完成汇总")

    rows = []
    for dt in all_dates:
        row = [dt.strftime("%Y-%m-%d"), "日"]
        total_labeled = 0
        total_done = 0
        for s in suppliers:
            sd = supplier_daily.get(s, {})
            labeled, done = sd.get(dt, ("-", "-"))
            row.append(str(labeled) if labeled != "-" else "-")
            row.append(str(done) if done != "-" else "-")
            if isinstance(labeled, (int, float)):
                total_labeled += labeled
            if isinstance(done, (int, float)):
                total_done += done
        row.append(str(total_labeled))
        row.append(str(total_done))
        rows.append(row)

    # 周汇总
    weekly_data = defaultdict(lambda: defaultdict(list))
    for dt in all_dates:
        iso_year, iso_week, _ = dt.isocalendar()
        weekly_data[(iso_year, iso_week)][dt].append(dt)

    for week_key in sorted(weekly_data.keys()):
        y, w = week_key
        week_dates = sorted(weekly_data[week_key].keys())
        fri_date = None
        for dt in week_dates:
            if dt.weekday() == 4:
                fri_date = dt
        if fri_date is None:
            continue

        row = [f"W{y - 2000:02d}{w:02d} ({fri_date.strftime('%m-%d')})", "周"]
        total_labeled = 0
        total_done = 0
        for s in suppliers:
            sd = supplier_daily.get(s, {})
            if fri_date in sd:
                labeled, done = sd[fri_date]
                row.append(str(labeled))
                row.append(str(done))
                if isinstance(labeled, (int, float)):
                    total_labeled += labeled
                if isinstance(done, (int, float)):
                    total_done += done
            else:
                row.append("-")
                row.append("-")
        row.append(str(total_labeled))
        row.append(str(total_done))
        rows.append(row)

    return headers, rows


def _build_quality_efficiency_sheet(proj_dash_data: list, proj_hr_data: dict) -> tuple:
    """质量效率汇总明细：累计良率、当日良率、产能占比、投入人数、速度人效、有效人效"""
    if not proj_dash_data:
        return [], []

    suppliers = []
    all_dates = set()
    supplier_daily = {}

    for supplier, dash_headers, dash_rows in proj_dash_data:
        suppliers.append(supplier)
        date_idx = _find_header_index(dash_headers, "日期")
        labeled_idx = _find_header_index(dash_headers, "已标注量")
        done_idx = _find_header_index(dash_headers, "已完成")
        if date_idx < 0 or labeled_idx < 0 or done_idx < 0:
            continue
        daily = {}
        for r in dash_rows:
            d_val = _safe_int(r[date_idx]) if len(r) > date_idx else 0
            if d_val <= 0:
                continue
            dt = EXCEL_EPOCH + timedelta(days=d_val)
            labeled = _safe_int(r[labeled_idx]) if len(r) > labeled_idx else 0
            done = _safe_int(r[done_idx]) if len(r) > done_idx else 0
            daily[dt] = (labeled, done)
            all_dates.add(dt)
        supplier_daily[supplier] = daily

    # 每日投入人数
    daily_persons = {}
    for supplier in suppliers:
        hr_rows, hr_headers = proj_hr_data.get(supplier, ([], []))
        if not hr_rows:
            daily_persons[supplier] = {}
            continue
        hr_date_idx = _find_header_index(hr_headers, "作业日期")
        hr_done_idx = _find_header_index(hr_headers, "已标注量")
        hr_en_idx = _find_header_index(hr_headers, "英文名")
        if hr_date_idx < 0 or hr_done_idx < 0:
            daily_persons[supplier] = {}
            continue
        dp = defaultdict(set)
        for r in hr_rows:
            d_val = _safe_int(r[hr_date_idx]) if len(r) > hr_date_idx else 0
            if d_val <= 0:
                continue
            dt = EXCEL_EPOCH + timedelta(days=d_val)
            done = _safe_int(r[hr_done_idx]) if len(r) > hr_done_idx else 0
            if done > 0:
                en = str(r[hr_en_idx]).strip() if len(r) > hr_en_idx else ""
                if en and en != "-":
                    dp[dt].add(en)
        daily_persons[supplier] = {dt: len(ps) for dt, ps in dp.items()}

    if not all_dates:
        return [], []

    all_dates = sorted(all_dates)

    headers = ["日期/周次", "类型"]
    for s in suppliers:
        headers.extend([f"{s}_累计良率", f"{s}_当日良率", f"{s}_产能占比",
                        f"{s}_投入人数", f"{s}_速度人效", f"{s}_有效人效"])

    rows = []
    prev_day = {}
    for dt in all_dates:
        day_data = {}
        total_day_labeled = 0
        for s in suppliers:
            sd = supplier_daily.get(s, {})
            labeled, done = sd.get(dt, ("-", "-"))
            day_data[s] = (labeled, done)
            prev_l, _ = prev_day.get(s, (None, None))
            if prev_l is not None and isinstance(labeled, (int, float)):
                total_day_labeled += max(0, labeled - prev_l)

        row = [dt.strftime("%Y-%m-%d"), "日"]
        for s in suppliers:
            labeled, done = day_data[s]
            prev_l, prev_d = prev_day.get(s, (None, None))
            dl = labeled - prev_l if (prev_l is not None and isinstance(labeled, (int, float))) else 0
            dd = done - prev_d if (prev_d is not None and isinstance(done, (int, float))) else 0

            # 累计良率
            cr = f"{round(done / labeled * 100, 1)}%" if (isinstance(labeled, (int, float)) and labeled > 0 and isinstance(done, (int, float))) else "-"
            # 当日良率
            dr = f"{round(dd / dl * 100, 1)}%" if dl > 0 else "-"
            # 产能占比
            sr = f"{round(dl / total_day_labeled * 100, 1)}%" if (dl > 0 and total_day_labeled > 0) else "-"
            # 投入人数
            pc = daily_persons.get(s, {}).get(dt, 0)
            # 速度人效
            se = str(round(dl / pc, 1)) if (pc > 0 and dl > 0) else "-"
            # 有效人效
            ee = str(round(dd / pc, 1)) if (pc > 0 and dd > 0) else "-"

            row.extend([cr, dr, sr, str(pc), se, ee])

            if isinstance(labeled, (int, float)):
                prev_day[s] = (labeled, done)

        rows.append(row)

    # 周汇总
    weekly_data = defaultdict(lambda: defaultdict(list))
    for dt in all_dates:
        weekly_data[dt.isocalendar()[:2]][dt].append(dt)

    for week_key in sorted(weekly_data.keys()):
        y, w = week_key
        week_dates = sorted(weekly_data[week_key].keys())
        fri_date = None
        for dt in week_dates:
            if dt.weekday() == 4:
                fri_date = dt
        if fri_date is None:
            continue

        row = [f"W{y - 2000:02d}{w:02d} ({fri_date.strftime('%m-%d')})", "周"]

        for s in suppliers:
            sd = supplier_daily.get(s, {})
            fd = sd.get(fri_date)
            if fd is None:
                row.extend(["-"] * 6)
                continue
            labeled, done = fd

            # 本周第一天累计值
            week_start = None
            for dt2 in week_dates:
                if dt2 in sd:
                    week_start = dt2
                    break
            wl = labeled - sd[week_start][0] if (week_start and isinstance(labeled, (int, float))) else 0
            wd = done - sd[week_start][1] if (week_start and isinstance(done, (int, float))) else 0

            pc = daily_persons.get(s, {}).get(fri_date, 0)
            cr = f"{round(done / labeled * 100, 1)}%" if (isinstance(labeled, (int, float)) and labeled > 0 and isinstance(done, (int, float))) else "-"
            wr = f"{round(wd / wl * 100, 1)}%" if wl > 0 else "-"
            sr = "-"
            for s2 in suppliers:
                s2_sd = supplier_daily.get(s2, {})
                if fri_date in s2_sd:
                    s2l = s2_sd[fri_date][0]
                    for dt2 in week_dates:
                        if dt2 in s2_sd:
                            s2_sl = s2_sd[dt2][0]
                            break
                    else:
                        continue
                    s2_wl = s2l - s2_sl if isinstance(s2l, (int, float)) and isinstance(s2_sl, (int, float)) else 0
                    if s2_wl > 0:
                        total_week_l = sum(
                            max(0, (s2_sd.get(fri_date, (0,))[0] if isinstance(s2_sd.get(fri_date, (0,))[0], (int, float)) else 0) -
                            (s2_sd.get(wd2, (0,))[0] if isinstance(s2_sd.get(wd2, (0,))[0], (int, float)) else 0))
                            for wd2 in week_dates if wd2 in s2_sd
                        )
                        break
            else:
                total_week_l = 0
            sr = f"{round(wl / total_week_l * 100, 1)}%" if (wl > 0 and total_week_l > 0) else "-"
            se = str(round(wl / pc / 5, 1)) if (pc > 0 and wl > 0) else "-"
            ee = str(round(wd / pc / 5, 1)) if (pc > 0 and wd > 0) else "-"

            row.extend([cr, wr, sr, str(pc), se, ee])

        rows.append(row)

    return headers, rows


def _build_hr_tracking_rows(progress: pd.DataFrame, quality: pd.DataFrame, serial: int,
                             hr_headers: list, name_map: dict, project: str,
                             old_rows: list, supplier: str = "") -> list:
    """构建人力 sheet 的今日数据行，区分项目格式"""
    if progress.empty:
        return []
    # 从旧数据提取默认供应商/项目名
    default_supplier = ""
    default_project = ""
    sidx = _find_header_index(hr_headers, "供应商")
    pidx = _find_header_index(hr_headers, "项目名称")
    for old_r in reversed(old_rows):
        if len(old_r) > sidx and str(old_r[sidx]).strip():
            default_supplier = str(old_r[sidx]).strip()
        if len(old_r) > pidx and str(old_r[pidx]).strip():
            default_project = str(old_r[pidx]).strip()
        if default_supplier and default_project:
            break

    rows = []
    seq = 1
    for _, prow in progress.iterrows():
        name = prow["标注员"]
        cn_name = name_map.get(name, "")
        qrow = quality[quality["标注员"] == name] if not quality.empty else pd.DataFrame()
        out_val = int(prow.get("已标注", 0) or 0)
        row_data = {"作业日期": str(serial), "序号": str(seq), "标注员_英文名": name, "标注员_中文名": cn_name,
                    "认领总量": str(int(prow.get("总量", 0) or 0)),
                    "已标注量": str(out_val),
                    "质检中": str(int(prow.get("质检中", 0) or 0)),
                    "验收中": str(int(prow.get("验收中", 0) or 0)),
                    "已驳回": str(int(prow.get("已驳回", 0) or 0)),
                    "已完成": str(int(prow.get("已完成", 0) or 0))}
        # 通过率（带明细）
        if len(qrow) > 0:
            qr = qrow.iloc[0]
            row_data["质检通过率"] = _fmt_rate_val(qr.get("质检通过率_pct"), qr.get("质检_pass"), qr.get("质检_total"))
            row_data["首次验收通过率"] = _fmt_rate_val(qr.get("首次验收通过率_pct"), qr.get("首次验收_pass"), qr.get("首次验收_total"))
            row_data["累积折损验收通过率"] = _fmt_rate_val(qr.get("累积验收通过率_pct"), qr.get("累积验收_pass"), qr.get("累积验收_total"))
        else:
            for k in ["质检通过率", "首次验收通过率", "累积折损验收通过率"]:
                row_data[k] = "-"
        # 项目B_ANOMALY专属列
        if "质检通过量" in hr_headers:
            qr = qrow.iloc[0] if len(qrow) > 0 else None
            row_data["质检通过量"] = str(int(qr.get("质检_pass", 0) or 0)) if qr is not None else "-"
            row_data["质检驳回量"] = str(int(qr.get("质检_total", 0) or 0) - int(qr.get("质检_pass", 0) or 0)) if qr is not None else "-"
            row_data["质检总量"] = str(int(qr.get("质检_total", 0) or 0)) if qr is not None else "-"
            row_data["验收通过量"] = str(int(qr.get("首次验收_pass", 0) or 0)) if qr is not None else "-"
            row_data["验收驳回量"] = str(int(qr.get("首次验收_total", 0) or 0) - int(qr.get("首次验收_pass", 0) or 0)) if qr is not None else "-"
            row_data["验收总量"] = str(int(qr.get("首次验收_total", 0) or 0)) if qr is not None else "-"
        rows.append(_row_dict_to_list(row_data, hr_headers))
        seq += 1
    return rows


def _build_ac_tracking_rows(ql: pd.DataFrame, serial: int) -> list:
    """构建验收明细 sheet 的今日数据行"""
    if ql.empty:
        return []
    rows = []
    for _, qrow in ql.iterrows():
        bdf = qrow.get("_batch_df")
        if bdf is None or bdf.empty:
            continue
        for _, brow in bdf.iterrows():
            rows.append([
                str(brow.get("批次名称/标注员", brow.get("批次名称", ""))),
                str(serial), "",
                str(int(brow.get("验收_total", 0) or 0)),
                str(int(brow.get("验收_total", 0) or 0)),
                str(int(brow.get("验收_pass", 0) or 0)),
                "", ""
            ])
    return rows


def _row_dict_to_list(row_data: dict, headers: list) -> list:
    """按表头顺序输出行数据，自动匹配列名"""
    result = []
    for h in headers:
        if not h:
            result.append("")
        elif h in row_data:
            result.append(row_data[h])
        elif h == "供应商":
            result.append("")
        elif h == "项目名称":
            result.append("")
        elif h == "上班卡时间" or h == "下班卡时间":
            result.append("")
        else:
            result.append("")
    return result


def _append_and_write_sheet(writer, sheet_name, headers, old_rows, new_rows, header_fmt, normal_fmt):
    """去重合并旧行+新行，用原生 xlsxwriter 控制格式（日期/数字/列宽）"""
    sheet_name = sheet_name[:31]
    seen = set()
    all_rows = []
    for r in new_rows + old_rows:
        # 跳过完全空行
        if not any(str(v).strip() for v in r if v):
            continue
        def _norm_key(v):
            if isinstance(v, float) and v == int(v):
                return str(int(v))
            return str(v)
        key = tuple(_norm_key(v) for v in r[:4]) if len(r) >= 4 else tuple(_norm_key(v) for v in r)
        if key and key not in seen:
            seen.add(key)
            all_rows.append(r)

    wb = writer.book
    date_fmt = wb.add_format({"align": "center", "num_format": "yyyy/mm/dd"})
    num_fmt = wb.add_format({"align": "center", "num_format": "#,##0"})
    pct_fmt = wb.add_format({"align": "center", "num_format": "0.0%"})

    # 列类型判断
    date_cols = set()
    num_cols = set()
    rate_cols = set()
    for ci, h in enumerate(headers):
        if not h:
            continue
        # 排除纯文本列
        if any(k in str(h) for k in ["标注员", "团队", "名称", "任务", "验收人", "供应商", "项目", "上班", "下班"]):
            pass
        elif any(k in h for k in ["日期", "时间"]):
            date_cols.add(ci)
        elif any(k in h for k in ["通过率", "准确率"]):
            rate_cols.add(ci)
        elif any(k in h for k in ["量", "中", "付", "成", "数", "增量", "总量", "驳回", "序号", "包总数",
                                   "已完成", "已标注", "已下发", "已交付", "累计"]):
            num_cols.add(ci)

    ws = wb.add_worksheet(sheet_name)
    # 写表头
    for ci in range(len(headers)):
        ws.write(0, ci, headers[ci] if ci < len(headers) else "", header_fmt)

    # 写数据行
    for ri, row in enumerate(all_rows):
        for ci in range(min(len(headers), len(row))):
            val = row[ci] if ci < len(row) else ""
            # 通过率列：有 "%" 保留明细，缺值 "-"，旧数值智能判定
            if ci in rate_cols:
                s = str(val).strip() if val else ""
                if not s or s in ("-", "nan", "None", ""):
                    ws.write(ri + 1, ci, "-", normal_fmt)
                elif "%" in s:
                    ws.write(ri + 1, ci, s, normal_fmt)
                else:
                    try:
                        v = float(s)
                        if v > 1:
                            ws.write(ri + 1, ci, f"{round(v,1)}%", normal_fmt)
                        else:
                            ws.write(ri + 1, ci, f"{round(v*100,1)}%", normal_fmt)
                    except ValueError:
                        ws.write(ri + 1, ci, "-", normal_fmt)
                continue
            # 日期列
            if ci in date_cols and val and str(val).strip():
                try:
                    d = EXCEL_EPOCH + timedelta(days=int(float(str(val))))
                    ws.write_datetime(ri + 1, ci, d, date_fmt)
                    continue
                except (ValueError, OverflowError):
                    pass
            # 数值列：转数字；非数值列（如名字）原样写入
            s = str(val).strip() if val else ""
            if ci not in num_cols:
                ws.write(ri + 1, ci, s if s and s not in ("nan", "None") else "-", normal_fmt)
            elif not s or s in ("nan", "None"):
                ws.write(ri + 1, ci, "-", normal_fmt)
            elif s == "-":
                ws.write(ri + 1, ci, "-", normal_fmt)
            elif any(c.isalpha() for c in s):
                ws.write_number(ri + 1, ci, 0, num_fmt)
            else:
                try:
                    ws.write_number(ri + 1, ci, float(s), num_fmt)
                except ValueError:
                    ws.write(ri + 1, ci, s, normal_fmt)

    # 自动列宽
    for ci in range(len(headers)):
        max_w = vis_width(str(headers[ci])) + 6
        for ri in range(len(all_rows)):
            v = str(all_rows[ri][ci]) if ci < len(all_rows[ri]) else ""
            max_w = max(max_w, vis_width(v) + 6)
        ws.set_column(ci, ci, min(max_w, 35))


if __name__ == "__main__":
    main()
