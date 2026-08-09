#!/usr/bin/env python3
"""核心业务逻辑（自 统一日报生成器.py 抽离）——波动分析与综合标记引擎。

纯函数、无 I/O，可独立单测。生产脚本 统一日报生成器.py 从这里 import，
本文件与脚本内同名逻辑保持一致（避免双份漂移）。

依赖：无第三方库。
"""

import math

# 波动率 ±10% 触发标记
WARN_THRESHOLD = 10.0


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
