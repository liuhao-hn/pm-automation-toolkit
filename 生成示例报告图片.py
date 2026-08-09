#!/usr/bin/env python3
"""
生成脱敏版示例日报/周报图片（Sample Images for Portfolio）
=========================================================
生成与原始工具相同风格的报告图片，使用脱敏后的占位数据。

用法：
  python3 生成示例报告图片.py
"""

import os
from datetime import datetime
from PIL import Image, ImageDraw
import PIL.ImageFont

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# ========== 字体 ==========
def _get_font(size: int, bold: bool = False):
    try:
        idx = 1 if bold else 0
        return PIL.ImageFont.truetype(
            "/System/Library/Fonts/Hiragino Sans GB.ttc", size, index=idx
        )
    except Exception:
        try:
            return PIL.ImageFont.load_default()
        except Exception:
            return None


_SMALL_FONT = lambda: _get_font(14)
_NORMAL_FONT = lambda: _get_font(16)
_BOLD_FONT = lambda: _get_font(16, True)
_H1_FONT = lambda: _get_font(20, True)
_KPI_VAL_FONT = lambda: _get_font(28, True)

# ========== 颜色 ==========
C_BLUE = "#1B3A5B"
C_ACCENT = "#2D7DD2"
C_GREEN = "#1B8C4A"
C_RED = "#D14343"
C_ORANGE = "#E67E22"
C_BG = "#EEF1F5"
C_CARD = "#FFFFFF"
C_TEXT = "#2C3E50"
C_SUB = "#6B7B8D"


def _fmt_num(n: int) -> str:
    return f"{n:,}"


def _text_width(text, font):
    try:
        return int(font.getlength(text))
    except Exception:
        return len(text) * int(font.size * 0.6)


def _draw_section_header(draw, canvas, x, y, w, text, bg="#2C3E50"):
    idr = ImageDraw.Draw(canvas)
    idr.rectangle([x, y, x + w, y + 36], fill=bg)
    idr.text((x + 10, y + 7), text, fill="#FFFFFF", font=_BOLD_FONT())
    return ImageDraw.Draw(canvas)


def _draw_table(draw, canvas, x, y, col_widths, headers, rows,
                header_bg="#3B5998", header_fg="#FFFFFF", stripe1="#F8F9FC", stripe2="#FFFFFF"):
    idr = ImageDraw.Draw(canvas)
    row_h = 28
    total_w = sum(col_widths)

    # header
    cx = x
    for ci, (cw, hdr) in enumerate(zip(col_widths, headers)):
        idr.rectangle([cx, y, cx + cw, y + row_h], fill=header_bg)
        idr.text((cx + 6, y + 5), hdr, fill=header_fg, font=_SMALL_FONT())
        cx += cw

    # rows
    for ri, row in enumerate(rows):
        ry = y + row_h + ri * row_h
        bg = stripe1 if ri % 2 == 0 else stripe2
        idr.rectangle([x, ry, x + total_w, ry + row_h], fill=bg)
        cx = x
        for ci, cw in enumerate(col_widths):
            cell = str(row[ci]) if ci < len(row) else ""
            idr.text((cx + 6, ry + 5), cell, fill=C_TEXT, font=_SMALL_FONT())
            cx += cw

    return (len(rows) + 1) * row_h


def generate_sample(report_type="日报"):
    """生成一张示例报告图片"""
    is_weekly = report_type == "周报"
    rp_label = report_type
    prefix_t = "本周" if is_weekly else "今日"

    project_name = "项目A_DATA" if report_type == "日报" else "项目B_ANOMALY"
    suppliers = (
        ["供应商Alpha", "供应商Beta", "供应商Gamma"]
        if project_name == "项目A_DATA"
        else ["供应商Delta", "供应商Epsilon", "供应商Zeta", "供应商Eta"]
    )

    W = 1240
    PAD = 30
    CARD_GAP = 16

    # ========== 示例数据 ==========
    total_plan = 56850
    cumul_anno = 48493
    cumul_deliv = 13478
    today_anno = 7900 if not is_weekly else 18200
    today_deliv = 1364 if not is_weekly else 3400
    active_count = 70
    total_count = 134
    att_pct = round(active_count / total_count * 100, 1)

    supplier_active = {"供应商Alpha": 8, "供应商Beta": 25, "供应商Gamma": 37}
    if "供应商Delta" in suppliers:
        supplier_active = {
            "供应商Delta": 12,
            "供应商Epsilon": 8,
            "供应商Zeta": 6,
            "供应商Eta": 10,
        }

    sup_data = {}
    if project_name == "项目A_DATA":
        sup_data = {
            "供应商Alpha": {"day_anno": 4645, "day_deliv": 680, "cum_anno": 22177, "cum_deliv": 6300, "total": 24760},
            "供应商Beta": {"day_anno": 653, "day_deliv": 120, "cum_anno": 9835, "cum_deliv": 3200, "total": 12400},
            "供应商Gamma": {"day_anno": 2602, "day_deliv": 564, "cum_anno": 16481, "cum_deliv": 3978, "total": 19690},
        }
    else:
        sup_data = {
            "供应商Delta": {"day_anno": 2100, "day_deliv": 350, "cum_anno": 15200, "cum_deliv": 4200, "total": 18500},
            "供应商Epsilon": {"day_anno": 980, "day_deliv": 180, "cum_anno": 7200, "cum_deliv": 1900, "total": 9200},
            "供应商Zeta": {"day_anno": 560, "day_deliv": 95, "cum_anno": 4800, "cum_deliv": 1100, "total": 6100},
            "供应商Eta": {"day_anno": 1340, "day_deliv": 220, "cum_anno": 10200, "cum_deliv": 2500, "total": 12800},
        }

    cumul_anno = sum(sd["cum_anno"] for sd in sup_data.values())
    cumul_deliv = sum(sd["cum_deliv"] for sd in sup_data.values())
    today_anno = sum(sd["day_anno"] for sd in sup_data.values())
    today_deliv = sum(sd["day_deliv"] for sd in sup_data.values())
    total_plan = sum(sd["total"] for sd in sup_data.values())
    active_count = sum(supplier_active.values())
    att_pct = round(active_count / total_count * 100, 1)
    cumul_anno_pct = round(cumul_anno / total_plan * 100, 1)

    # ========== 画布 ==========
    img_h = 2400
    img = Image.new("RGB", (W, img_h), C_BG)
    draw = ImageDraw.Draw(img)
    y = 0

    # ═══ 1. 标题栏 ═══
    draw.rectangle([0, y, W, y + 64], fill=C_BLUE)
    title = f"{project_name}  {rp_label}  |  Demo  |  示例输出"
    draw.text((PAD + 8, y + 12), title, fill="#FFFFFF", font=_H1_FONT())
    draw.text((PAD + 8, y + 38), "2026/6/18  (Portfolio Sample)", fill="#B8C5D0", font=_SMALL_FONT())
    y += 64 + CARD_GAP

    # ═══ 2. KPI 卡片 ═══
    kpi_items = [
        (f"{prefix_t}累计标注量", _fmt_num(cumul_anno), "2026/6/18"),
        (f"标注{'周' if is_weekly else '日'}增量", _fmt_num(today_anno), "对比上期"),
        (f"{prefix_t}累计已完成量", _fmt_num(cumul_deliv), "2026/6/18"),
        (f"已完成{'周' if is_weekly else '日'}增量", _fmt_num(today_deliv), "对比上期"),
        ("出勤率", f"{att_pct}%", f"{active_count}/{total_count} 人在岗"),
    ]
    n_kpi = len(kpi_items)
    card_w = (W - PAD * 2 - CARD_GAP * (n_kpi - 1)) // n_kpi
    card_h = 88

    for i, (label, value, sub) in enumerate(kpi_items):
        cx = PAD + i * (card_w + CARD_GAP)
        idr = ImageDraw.Draw(img)
        idr.rounded_rectangle(
            [cx, y, cx + card_w, y + card_h], radius=8, fill=C_CARD, outline="#D0D8E4", width=1
        )
        idr.rounded_rectangle(
            [cx + 1, y + 1, cx + card_w - 1, y + 7], radius=3, fill=C_ACCENT
        )
        draw.text(
            (cx + card_w // 2 - _text_width(label, _SMALL_FONT()) // 2, y + 14),
            label, fill=C_SUB, font=_SMALL_FONT(),
        )
        draw.text(
            (cx + card_w // 2 - _text_width(value, _KPI_VAL_FONT()) // 2, y + 34),
            value, fill=C_TEXT, font=_KPI_VAL_FONT(),
        )
        draw.text(
            (cx + card_w // 2 - _text_width(sub, _SMALL_FONT()) // 2, y + 66),
            sub, fill=C_SUB, font=_SMALL_FONT(),
        )
    y += card_h + CARD_GAP

    # 累计统计
    cumul_text = (
        f"较计划总量 {_fmt_num(total_plan)} 完成 {cumul_anno_pct}%   |   "
        f"已标注 {_fmt_num(cumul_anno)} 条   |   已完成 {_fmt_num(cumul_deliv)} 条"
    )
    win_text = "对比窗口：2026/6/17 → 2026/6/18（1 天）" if not is_weekly else "对比窗口：2026/6/10 → 2026/6/17（7 天）"

    idr = ImageDraw.Draw(img)
    idr.rounded_rectangle([PAD, y, W - PAD, y + 56], radius=6, fill=C_CARD, outline="#D0D8E4", width=1)
    draw.text((PAD + 16, y + 8), cumul_text, fill=C_TEXT, font=_NORMAL_FONT())
    draw.text((PAD + 16, y + 30), win_text, fill=C_SUB, font=_SMALL_FONT())
    y += 56 + CARD_GAP + 6

    # ═══ 3. 人员在岗情况及今日重点 ═══
    draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2,
                                f"   人员在岗情况及{prefix_t}重点工作")
    y += 36 + 10

    att_lines = [
        f"应出勤 {total_count} 人（账号），{prefix_t}有产出 {active_count} 人，出勤率 {att_pct}%。"
    ]
    for sup in suppliers:
        sa = supplier_active.get(sup, 0)
        totals = {"供应商Alpha": 14, "供应商Beta": 56, "供应商Gamma": 64,
                  "供应商Delta": 18, "供应商Epsilon": 10, "供应商Zeta": 8, "供应商Eta": 14}
        st = totals.get(sup, 20)
        att_lines.append(f"  {sup}：{sa}/{st} 人有产出")
    att_h = 12 + len(att_lines) * 22 + 12

    idr = ImageDraw.Draw(img)
    idr.rounded_rectangle([PAD, y, W - PAD, y + att_h], radius=8, fill=C_CARD, outline="#D0D8E4", width=1)
    for li, line in enumerate(att_lines):
        draw.text((PAD + 20, y + 10 + li * 22), line, fill=C_TEXT, font=_NORMAL_FONT())
    y += att_h + CARD_GAP

    # 重点
    foci = [
        "关注异常预警：共 3 项，需复查波动超阈值的标注员/团队。",
        "低质量标注员需重点复核，未达标条目逐一排查。",
        "正向优秀：2 项表现突出，可总结经验推广。",
    ]
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

    col_anno = "本周标注" if is_weekly else "今日标注"
    col_done = "本周已完成" if is_weekly else "今日已完成"
    plan_cols = [160, 120, 90, 130, 130, 110]
    plan_header = ["供应商", col_anno, "人效", col_done, "累计标注", "完成率"]
    plan_rows = []
    for sup in suppliers:
        sd = sup_data[sup]
        cp = f"{round(sd['cum_anno'] / sd['total'] * 100, 1)}%" if sd["total"] > 0 else "-"
        sa = supplier_active.get(sup, 0)
        sup_eff = f"{round(sd['day_anno'] / sa, 2)}" if sa > 0 else "-"
        plan_rows.append([
            sup,
            _fmt_num(sd["day_anno"]),
            sup_eff,
            _fmt_num(sd["day_deliv"]),
            _fmt_num(sd["cum_anno"]),
            cp,
        ])
    th = _draw_table(draw, img, PAD, y, plan_cols, plan_header, plan_rows)
    y += th + CARD_GAP

    desc_lines = [
        f"* 标注{'周' if is_weekly else '日'}增量 {_fmt_num(today_anno)}，"
        f"已完成{'周' if is_weekly else '日'}增量 {_fmt_num(today_deliv)}。",
        f"* 累计标注 {cumul_anno_pct}%，进度过半，保持当前节奏。",
    ]
    for dl in desc_lines:
        draw.text((PAD + 4, y), dl, fill=C_SUB, font=_SMALL_FONT())
        y += 18
    y += CARD_GAP

    # ═══ 5. 整体批次完成情况 ═══
    draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, "   整体批次完成情况")
    y += 36 + 10

    batches = [
        ["batch-001", "供应商Alpha", "100.0%", "75.0%", "验收中"],
        ["batch-002", "供应商Alpha", "100.0%", "33.3%", "验收中"],
        ["batch-003", "供应商Alpha", "100.0%", "85.7%", "验收中"],
        ["batch-004", "供应商Beta", "100.0%", "50.0%", "验收中"],
        ["batch-005", "供应商Beta", "100.0%", "66.7%", "验收中"],
        ["batch-006", "供应商Beta", "0.0%", "100.0%", "验收中"],
        ["batch-007", "供应商Gamma", "100.0%", "71.4%", "验收中"],
        ["batch-008", "供应商Gamma", "60.0%", "0.0%", "质检中"],
        ["batch-009", "供应商Gamma", "100.0%", "75.0%", "验收中"],
        ["batch-010", "供应商Gamma", "100.0%", "95.0%", "已完成"],
    ]
    batch_count = len(batches)
    bt_cols = [140, 200, 90, 90, 90]
    bt_header = ["批次", "供应商", "质检率", "验收率", "状态"]
    show_batches = batches[:10]
    th = _draw_table(draw, img, PAD, y, bt_cols, bt_header, show_batches)
    y += th + 6
    suffix = f"  （共 {batch_count} 个主批次，展示前10个，数据为示例）"
    draw.text((PAD + 4, y), suffix, fill=C_SUB, font=_SMALL_FONT())
    y += 22 + CARD_GAP

    # ═══ 6. 验收明细 ═══
    draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, "   业务方验收明细")
    y += 36 + 10

    acc_rows = [
        ["供应商Alpha", "85.7%(120/140)+2.1pp", "78.2%(115/147)+3.5pp", "92.3%(12/13)", "81.8%(9/11)", "验收中"],
        ["供应商Beta", "92.1%(105/114)-1.3pp", "75.0%(60/80)+5.0pp", "88.9%(8/9)", "71.4%(5/7)", "验收中"],
        ["供应商Gamma", "90.5%(76/84)+0.8pp", "80.0%(40/50)+2.0pp", "100.0%(5/5)", "83.3%(5/6)", "验收中"],
    ]
    ac_cols = [200, 210, 210, 160, 160, 100]
    ac_header = ["团队", "累计质检(波动)", "累计验收(波动)", "日质检", "日验收", "状态"]
    th = _draw_table(draw, img, PAD, y, ac_cols, ac_header, acc_rows)
    y += th + CARD_GAP

    # ═══ 7. 明日/下周工作安排 ═══
    next_sec = "   下周工作安排" if is_weekly else "   明日工作安排"
    draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, next_sec)
    y += 36 + 10

    tasks = [
        "跟进异常标注员复查，确保质量回升至基线。",
        "对低质量标注员完成逐题复盘与专项培训。",
        "跟进待验收批次的业务方反馈。",
        "推进剩余队列标注进度，向收尾冲刺。",
    ]
    tasks_h = 12 + len(tasks) * 24 + 12
    idr = ImageDraw.Draw(img)
    idr.rounded_rectangle([PAD, y, W - PAD, y + tasks_h], radius=8, fill=C_CARD, outline="#D0D8E4", width=1)
    for ti, task in enumerate(tasks):
        draw.text((PAD + 20, y + 10 + ti * 24), f"{ti + 1}. {task}", fill=C_TEXT, font=_NORMAL_FONT())
    y += tasks_h + CARD_GAP + 10

    # ═══ 8. 页脚声明 ═══
    y += 10
    draw = _draw_section_header(draw, img, PAD, y, W - PAD * 2, "   Portfolio Sample  |  数据为示例值  |  原始代码已脱敏")
    y += 36 + 10

    footer_lines = [
        "⚠️ 本图片为脱敏示例（Portfolio Sample），所有数据均为虚构，仅用于展示报告格式与工程能力。",
        "原始代码已在实际生产环境中稳定运行数月，覆盖多个项目、多家供应商的全流程日报/周报自动生成。",
        f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
    ]
    for fl in footer_lines:
        draw.text((PAD + 4, y), fl, fill=C_SUB, font=_SMALL_FONT())
        y += 20

    y += PAD

    # ── 裁剪底部 ──
    final_h = y
    img = img.crop((0, 0, W, final_h))

    out_name = f"{project_name}_{rp_label}_Sample.png"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    img.save(out_path, "PNG", optimize=True)
    print(f"✅ 已生成: {out_path} ({img.width}×{img.height})")
    return out_path


if __name__ == "__main__":
    print("═" * 50)
    print("  生成脱敏版示例报告图片")
    print("═" * 50)
    generate_sample("日报")
    print()
    generate_sample("周报")
    print()
    print("完成！请查看 项目管理-脱敏版/ 目录下的 PNG 文件。")
