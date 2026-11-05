#!/usr/bin/env python3
"""営業共有用：架空ブランドのCPAアラート表（xlsx）を作る。

実案件名・実数値は入れない。緑／黄／赤の見え方だけ見せる。
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parents[1] / "demo" / "sales-demo-cpa-alert.xlsx"

GREEN = "D9EAD3"
YELLOW = "FFF2CC"
RED = "F4CCCC"
HEADER = "F3F3F3"

# 架空データ（緑・黄・赤が揃うように置く）
ROWS = [
    # CP, 媒体, 当月費用, 当月CV, 目標, 許容
    ("ブランドA_検索", "Google", 180000, 90, 2000, 5500),   # CPA 2000 → 緑
    ("ブランドA_Meta", "Meta", 220000, 200, 2000, 5500),    # CPA 1100 → 緑
    ("ブランドB_検索", "Google", 150000, 40, 2500, 7000),   # CPA 3750 → 黄
    ("ブランドB_Meta", "Meta", 96000, 32, 2500, 7000),      # CPA 3000 → 黄
    ("ブランドC_検索", "Google", 210000, 20, 3000, 8000),   # CPA 10500 → 赤
    ("ブランドC_Meta", "Meta", 45000, 5, 3000, 8000),       # CPA 9000 → 赤
    ("ブランドD_認知", "Meta", 80000, 0, None, None),       # CV0 → 色なし
]


def band(cpa, target, allow):
    if cpa is None or target is None or allow is None:
        return ""
    if cpa <= target:
        return "緑"
    if cpa <= allow:
        return "黄"
    return "赤"


def score(label):
    return {"緑": 1, "黄": 2, "赤": 3}.get(label, "")


def main():
    wb = Workbook()
    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    # --- 目標 ---
    goals = wb.active
    goals.title = "目標"
    goals.append(["CP", "目標CPA", "許容CPA", "当月予算"])
    for cp, _media, _cost, _cv, target, allow in ROWS:
        goals.append([cp, target or "", allow or "", 300000])
    for col in range(1, 5):
        goals.cell(1, col).fill = PatternFill("solid", fgColor=HEADER)
        goals.cell(1, col).font = Font(bold=True)

    # --- 月次サマリ ---
    summary = wb.create_sheet("月次サマリ")
    headers = [
        "CP", "媒体", "当月_費用", "当月_CV", "当月_CPA", "目標CPA", "目標比",
        "前月_費用", "前月_CV", "前月_CPA",
        "", "", "", "", "", "", "", "", "", "", "", "", "",
        "許容CPA", "CPA帯", "CPAスコア",
    ]
    summary.append(headers)
    for col, _ in enumerate(headers, 1):
        cell = summary.cell(1, col)
        cell.fill = PatternFill("solid", fgColor=HEADER)
        cell.font = Font(bold=True)

    fills = {"緑": GREEN, "黄": YELLOW, "赤": RED}
    for i, (cp, media, cost, cv, target, allow) in enumerate(ROWS, start=2):
        cpa = round(cost / cv) if cv else None
        gap = round((target - cpa) / target, 4) if target and cpa is not None else ""
        label = band(cpa, target, allow)
        prev_cost = int(cost * 0.9)
        prev_cv = max(cv - 2, 0) if cv else 0
        prev_cpa = round(prev_cost / prev_cv) if prev_cv else None
        row = [""] * 26
        row[0] = cp
        row[1] = media
        row[2] = cost
        row[3] = cv if cv else ""
        row[4] = cpa if cpa is not None else ""
        row[5] = target or ""
        row[6] = gap
        row[7] = prev_cost
        row[8] = prev_cv if prev_cv else ""
        row[9] = prev_cpa if prev_cpa is not None else ""
        row[23] = allow or ""
        row[24] = label
        row[25] = score(label)
        summary.append(row)
        if label in fills:
            summary.cell(i, 5).fill = PatternFill("solid", fgColor=fills[label])
            summary.cell(i, 25).fill = PatternFill("solid", fgColor=fills[label])
        for col in range(1, 27):
            summary.cell(i, col).border = thin
            summary.cell(i, col).alignment = Alignment(horizontal="center")

    # 条件付き書式（再計算しても色が残る）
    last = 1 + len(ROWS)
    green_fill = PatternFill("solid", fgColor=GREEN)
    yellow_fill = PatternFill("solid", fgColor=YELLOW)
    red_fill = PatternFill("solid", fgColor=RED)
    summary.conditional_formatting.add(
        f"E2:E{last}",
        FormulaRule(formula=['AND($E2<>"",$F2<>"",$E2<=$F2)'], fill=green_fill),
    )
    summary.conditional_formatting.add(
        f"E2:E{last}",
        FormulaRule(
            formula=['AND($E2<>"",$F2<>"",$X2<>"",$E2>$F2,$E2<=$X2)'],
            fill=yellow_fill,
        ),
    )
    summary.conditional_formatting.add(
        f"E2:E{last}",
        FormulaRule(formula=['AND($E2<>"",$X2<>"",$E2>$X2)'], fill=red_fill),
    )

    # --- 凡例 ---
    legend = wb.create_sheet("凡例")
    legend.append(["色", "条件", "営業デモ用の説明"])
    legend.append(["緑", "実績CPA ≤ 目標CPA", "目標内。増額や維持の候補"])
    legend.append(["黄", "目標 < 実績 ≤ 許容", "許容内。改善しながら継続"])
    legend.append(["赤", "実績 > 許容", "赤字ライン超え。除外・停止・入札見直し"])
    legend.append(["（なし）", "目標またはCVが空", "認知などCPA対象外"])
    for i, color in enumerate([GREEN, YELLOW, RED], start=2):
        legend.cell(i, 1).fill = PatternFill("solid", fgColor=color)
    legend["A7"] = "※このファイルの数値はすべてダミーです。本番のクライアント名・実数値ではありません。"
    legend["A7"].font = Font(bold=True, color="990000")

    for ws in (goals, summary, legend):
        for col in range(1, 12):
            ws.column_dimensions[get_column_letter(col)].width = 14
        ws.column_dimensions["A"].width = 18

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
