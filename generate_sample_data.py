#!/usr/bin/env python3
"""生成脱敏示例数据，让日报生成器可以端到端跑起来。

用法：
    python3 generate_sample_data.py            # 生成项目A_DATA/供应商Alpha/团队/ 示例
    python3 "统一日报生成器.py" --dry-run --date 2026-6-21   # 跑通团队绩效对比

说明：
- 生成的是合成数据（标注员/数字全部虚构），不含任何真实业务数据。
- 覆盖「团队绩效」路径（标注团队绩效明细表）。项目大盘（整体/）与验收进度 xlsx 未含在最小示例中。
- 标注员 ID 刻意避开脚本 EXCLUDED_ANNOTATORS 中的离职名单，否则会被过滤。
"""

from pathlib import Path

BASE = Path(__file__).resolve().parent
TEAM_DIR = BASE / "项目A_DATA" / "供应商Alpha" / "团队"

# (标注员, 已标注, 质检通过率, 验收首次通过率, 验收累积通过率) —— 格式 "90.0% (180/200)"
ROWS_NEW = [
    ("user_i009", 180, "90.0% (180/200)", "85.0% (170/200)", "80.0% (200/250)"),
    ("user_j010", 150, "95.0% (150/158)", "90.0% (140/155)", "85.0% (180/212)"),
    ("未分配", 50, "50.0% (25/50)", "40.0% (20/50)", "35.0% (30/86)"),  # 会被脚本过滤
]
ROWS_OLD = [
    ("user_i009", 160, "88.0% (160/182)", "82.0% (150/183)", "76.0% (170/224)"),
    ("user_j010", 130, "92.0% (130/141)", "88.0% (125/142)", "82.0% (150/183)"),
    ("未分配", 40, "45.0% (18/40)", "38.0% (15/39)", "32.0% (25/78)"),
]

FILES = [
    ("标注团队绩效明细表_历史全量_2026_6_20.xls", ROWS_OLD),
    ("标注团队绩效明细表_历史全量_2026_6_21.xls", ROWS_NEW),
    ("标注团队绩效明细表_当天实时_2026_6_20.xls", ROWS_OLD),
    ("标注团队绩效明细表_当天实时_2026_6_21.xls", ROWS_NEW),
]


def team_html(rows) -> str:
    head = (
        '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>'
        "<table><tr><th>标注员</th><th>已标注</th>"
        "<th>🛡️质检通过率</th><th>🎯验收(首次验收通过)</th><th>🎯验收(累积折损通过)</th></tr>"
    )
    body = "".join(
        f"<tr><td>{n}</td><td>{d}</td><td>{q}</td><td>{a}</td><td>{b}</td></tr>"
        for n, d, q, a, b in rows
    )
    return head + body + "</table></body></html>"


def main() -> None:
    TEAM_DIR.mkdir(parents=True, exist_ok=True)
    for fname, rows in FILES:
        (TEAM_DIR / fname).write_text(team_html(rows), encoding="utf-8")
        print(f"  已生成: 项目A_DATA/供应商Alpha/团队/{fname}")
    print("完成。运行：python3 \"统一日报生成器.py\" --dry-run --date 2026-6-21")


if __name__ == "__main__":
    main()
