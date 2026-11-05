#!/usr/bin/env python3
"""営業共有用：定例広告報告と同じ型のダミーブックを作る。

全体サマリ＋Google（成果・入札・クエリ・KW・広告文）＋Meta（成果・CR）＋認知＋Next Action。
実案件名・実数値は入れない。
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parents[1] / "demo" / "sales-demo-cpa-alert.xlsx"

GREEN = "D9EAD3"
YELLOW = "FFF2CC"
RED = "F4CCCC"
HEADER = "F3F3F3"
NOTE = "CFE2F3"

# 架空ブランド3本（緑・黄・赤が揃う）
# name, media_label, day_budget, target, allow, cpa7, cv7, cpam, cvm, costm_month, cv_month, band
BRANDS = [
    {
        "key": "A",
        "name": "ブランドA",
        "group": "主力",
        "day_budget": 30000,
        "target": 2000,
        "allow": 5500,
        "g": {
            "cost7": 98000, "cv7": 55, "cpa7": 1782,  # 緑
            "cost_m": 410000, "cv_m": 220, "cpa_m": 1864,
            "cost_cur": 52000, "cv_cur": 28, "cpa_cur": 1857,
            "issue": "7日CPA 1,782円。目標2,000円の内側。",
            "action": "日予算を据え置き、来週も維持する",
        },
        "m": {
            "cost7": 140000, "cv7": 120, "cpa7": 1167,  # 緑
            "cost_m": 520000, "cv_m": 480, "cpa_m": 1083,
            "cost_cur": 68000, "cv_cur": 62, "cpa_cur": 1097,
            "issue": "7日CPA 1,167円。目標内。",
            "action": "予算は据え置き",
        },
    },
    {
        "key": "B",
        "name": "ブランドB",
        "group": "主力",
        "day_budget": 25000,
        "target": 2500,
        "allow": 7000,
        "g": {
            "cost7": 112000, "cv7": 30, "cpa7": 3733,  # 黄
            "cost_m": 380000, "cv_m": 95, "cpa_m": 4000,
            "cost_cur": 48000, "cv_cur": 13, "cpa_cur": 3692,
            "issue": "7日CPA 3,733円。目標超・許容内。",
            "action": "許容超えクエリを除外する",
        },
        "m": {
            "cost7": 95000, "cv7": 32, "cpa7": 2969,  # 黄
            "cost_m": 310000, "cv_m": 110, "cpa_m": 2818,
            "cost_cur": 41000, "cv_cur": 14, "cpa_cur": 2929,
            "issue": "7日CPA 2,969円。目標超・許容内。",
            "action": "CPAが高い動画を止める",
        },
    },
    {
        "key": "C",
        "name": "ブランドC",
        "group": "新規",
        "day_budget": 12000,
        "target": 3000,
        "allow": 8000,
        "g": {
            "cost7": 84000, "cv7": 7, "cpa7": 12000,  # 赤
            "cost_m": 210000, "cv_m": 20, "cpa_m": 10500,
            "cost_cur": 28000, "cv_cur": 2, "cpa_cur": 14000,
            "issue": "7日CPA 12,000円。許容8,000円を超えた。",
            "action": "高CPAクエリを除外し、入札上限を見直す",
        },
        "m": {
            "cost7": 45000, "cv7": 5, "cpa7": 9000,  # 赤
            "cost_m": 120000, "cv_m": 12, "cpa_m": 10000,
            "cost_cur": 18000, "cv_cur": 2, "cpa_cur": 9000,
            "issue": "7日CPA 9,000円。許容超え。",
            "action": "学習中の広告セットを止め、勝ちクリエイティブに寄せる",
        },
    },
]


def yen(n):
    return f"{int(n):,}"


def gap(target, cpa):
    if target is None or cpa is None:
        return "—"
    pct = round((target - cpa) / target * 100)
    if pct < 0:
        return f"△{abs(pct)}%"
    return f"+{pct}%"


def band_fill(cpa, target, allow):
    if cpa is None or target is None or allow is None:
        return None
    if cpa <= target:
        return GREEN
    if cpa <= allow:
        return YELLOW
    return RED


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = PatternFill("solid", fgColor=HEADER)
        cell.font = Font(bold=True)


def autosize(ws, widths=None):
    for i in range(1, 16):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(i, 14)
    ws.column_dimensions["A"].width = 22


def fill_cpa_cell(cell, cpa, target, allow):
    color = band_fill(cpa, target, allow)
    if color:
        cell.fill = PatternFill("solid", fgColor=color)


def add_google_campaign(wb, brand):
    name = brand["name"]
    g = brand["g"]
    target, allow = brand["target"], brand["allow"]
    month_label = "想定月額"

    # 成果
    ws = wb.create_sheet(f"G_{name}＿成果")
    ws["A1"] = f"{name}　成果"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"{brand['group']} / Google。日予算 {yen(brand['day_budget'])}円。{month_label} {yen(brand['day_budget']*30)}円。"
    ws.append([])
    ws["A3"] = "期間"
    for i, h in enumerate(["期間", "費用", "CV", "CPA", "目標CPA", "目標比"], 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 6)
    rows = [
        ("直近7日", g["cost7"], g["cv7"], g["cpa7"]),
        ("前月", g["cost_m"], g["cv_m"], g["cpa_m"]),
        ("当月", g["cost_cur"], g["cv_cur"], g["cpa_cur"]),
    ]
    for r_i, (period, cost, cv, cpa) in enumerate(rows, start=4):
        ws.cell(r_i, 1, period)
        ws.cell(r_i, 2, yen(cost))
        ws.cell(r_i, 3, cv)
        ws.cell(r_i, 4, cpa)
        ws.cell(r_i, 5, yen(target))
        ws.cell(r_i, 6, gap(target, cpa))
        fill_cpa_cell(ws.cell(r_i, 4), cpa, target, allow)
    ws.cell(7, 1, "課題")
    ws.cell(7, 2, g["issue"])
    ws.cell(8, 1, "今日決めること")
    ws.cell(8, 2, g["action"])
    autosize(ws)

    # 入札
    ws = wb.create_sheet(f"G_{name}＿入札")
    ws["A1"] = f"{name}　入札"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "上ほど古い施策、下ほど新しい施策。期間の費用とCPA。新しい施策は下に足す。"
    headers = ["期間", "入札戦略", "入札の目標CPA", "費用", "CV", "期間CPA", "メモ", "判断"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 8)
    bid_rows = [
        ("10/1〜10/15", "コンバージョン数の最大化", yen(int(target * 0.7)), yen(int(g["cost_m"] * 0.4)), int(g["cv_m"] * 0.45), int(g["cpa_m"] * 0.85), f"目標CPA {yen(int(target*0.7))}円", ""),
        ("10/16〜10/25", "コンバージョン数の最大化", "なし", yen(int(g["cost_m"] * 0.35)), int(g["cv_m"] * 0.3), int(g["cpa_m"] * 1.2), "入札の目標CPAを外した", "目標を戻すか"),
        ("10/26〜11/1", "コンバージョン数の最大化", yen(target), yen(g["cost7"]), g["cv7"], g["cpa7"], f"入札目標CPAを{yen(target)}円にした", g["action"]),
    ]
    for r_i, row in enumerate(bid_rows, start=4):
        for c_i, v in enumerate(row, 1):
            ws.cell(r_i, c_i, v)
        fill_cpa_cell(ws.cell(r_i, 6), row[5] if isinstance(row[5], int) else None, target, allow)
    autosize(ws, {3: 18, 7: 28, 8: 22})

    # クエリ
    ws = wb.create_sheet(f"G_{name}＿クエリ")
    ws["A1"] = f"{name}　クエリ"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"直近30日。赤は許容CPA {yen(allow)}円を超えた語句。除外するかを決める。"
    for i, h in enumerate(["検索語句", "表示", "クリック", "費用", "CV", "CPA", "判断"], 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 7)
    # fictional queries - no real client terms
    base = name.replace("ブランド", "ギフト")
    queries = [
        (f"{base}", 4200, 800, 180000, 90, 2000, "残す"),
        (f"{base} 作成", 1800, 420, 95000, 40, 2375, "残す"),
        (f"オリジナル {base}", 900, 210, 52000, 18, 2889, "残す"),
        (f"{base} 安い", 400, 90, 28000, 8, 3500, "残す" if allow >= 3500 else "除外する"),
        (f"{base} デザイン", 350, 70, 24000, 2, 12000, "除外する"),
        (f"オーダーメイド プレゼント", 280, 150, 18000, 1, 18000, "除外する"),
        (f"{base} 最速", 120, 40, 15000, 1, 15000, "除外する"),
    ]
    for r_i, row in enumerate(queries, start=4):
        for c_i, v in enumerate(row, 1):
            ws.cell(r_i, c_i, v if c_i != 4 else yen(v) if c_i == 4 else v)
        # fix cost display
        ws.cell(r_i, 4, yen(row[3]))
        fill_cpa_cell(ws.cell(r_i, 6), row[5], target, allow)
        if row[5] > allow:
            ws.cell(r_i, 6).fill = PatternFill("solid", fgColor=RED)
    ws.cell(12, 1, "課題")
    ws.cell(12, 2, f"許容超えの語句を除外する。今日決めることは1つ。")
    autosize(ws, {1: 24, 7: 12})

    # KW
    ws = wb.create_sheet(f"G_{name}＿KW")
    ws["A1"] = f"{name}　キーワード"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "直近30日、費用の大きい順。止めるキーワードを決める。"
    for i, h in enumerate(["キーワード", "一致", "表示", "クリック", "費用", "CV", "CPA", "判断"], 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 8)
    kws = [
        (f"{base}", "フレーズ", 6500, 1000, 220000, 100, 2200, "残す"),
        (f"{base} 作成", "フレーズ", 2100, 450, 110000, 48, 2292, "残す"),
        (f"{base} オーダー", "フレーズ", 3000, 500, 98000, 35, 2800, "残す"),
        (f"{base} ラベル", "フレーズ", 400, 50, 12000, 0, "", "止める"),
        (f"{base} 推し", "部分一致", 380, 70, 9000, 1, 9000, "残す" if allow >= 9000 else "止める"),
    ]
    for r_i, row in enumerate(kws, start=4):
        for c_i, v in enumerate(row, 1):
            if c_i == 5:
                ws.cell(r_i, c_i, yen(v))
            else:
                ws.cell(r_i, c_i, v)
        if isinstance(row[6], int):
            fill_cpa_cell(ws.cell(r_i, 7), row[6], target, allow)
    autosize(ws)

    # 広告文
    ws = wb.create_sheet(f"G_{name}＿広告文")
    ws["A1"] = f"{name}　広告文"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "直近30日。費用は行ごとに足さない。同時に出た分が各行に入る。"
    ws["A3"] = "審査の「制限付き」はデモ用の表示例。審査落ちではない。"
    for i, h in enumerate(["種類", "文面", "審査", "表示", "クリック", "費用", "CV", "CPA", "判断"], 1):
        ws.cell(4, i, h)
    style_header(ws, 4, 9)
    ads = [
        ("見出し", f"{{Keyword:{base}}}", "適格", 15000, 2200, 480000, 180, 2667, "残す"),
        ("見出し", f"{name} 1本から注文可", "適格", 12000, 1800, 390000, 150, 2600, "残す"),
        ("見出し", "最短翌日発送", "適格", 9000, 1100, 280000, 100, 2800, "残す"),
        ("見出し", "デザイン費無料", "適格", 5000, 700, 150000, 40, 3750, "残す"),
        ("説明文", "写真やロゴでオーダー。全国配送。", "適格", 8000, 900, 200000, 70, 2857, "残す"),
    ]
    for r_i, row in enumerate(ads, start=5):
        for c_i, v in enumerate(row, 1):
            if c_i == 6:
                ws.cell(r_i, c_i, yen(v))
            else:
                ws.cell(r_i, c_i, v)
        fill_cpa_cell(ws.cell(r_i, 8), row[7], target, allow)
    autosize(ws, {2: 32})


def add_meta_campaign(wb, brand):
    name = brand["name"]
    m = brand["m"]
    target, allow = brand["target"], brand["allow"]

    ws = wb.create_sheet(f"M_{name}リード＿成果")
    ws["A1"] = f"{name}リード　成果"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"{brand['group']} / Meta。日予算 {yen(int(brand['day_budget']*1.1))}円。想定月額 {yen(int(brand['day_budget']*1.1*30))}円。"
    for i, h in enumerate(["期間", "費用", "CV", "CPA", "目標CPA", "目標比"], 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 6)
    rows = [
        ("直近7日", m["cost7"], m["cv7"], m["cpa7"]),
        ("前月", m["cost_m"], m["cv_m"], m["cpa_m"]),
        ("当月", m["cost_cur"], m["cv_cur"], m["cpa_cur"]),
    ]
    for r_i, (period, cost, cv, cpa) in enumerate(rows, start=4):
        ws.cell(r_i, 1, period)
        ws.cell(r_i, 2, yen(cost))
        ws.cell(r_i, 3, cv)
        ws.cell(r_i, 4, cpa)
        ws.cell(r_i, 5, yen(target))
        ws.cell(r_i, 6, gap(target, cpa))
        fill_cpa_cell(ws.cell(r_i, 4), cpa, target, allow)
    ws.cell(7, 1, "課題")
    ws.cell(7, 2, m["issue"])
    ws.cell(8, 1, "今日決めること")
    ws.cell(8, 2, m["action"])
    autosize(ws)

    ws = wb.create_sheet(f"M_{name}リード＿CR")
    ws["A1"] = f"{name}リード　クリエイティブ"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "直近7日。止める広告と残す広告を決める。"
    for i, h in enumerate(["広告", "表示", "クリック", "費用", "CV", "CPA", "判断"], 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 7)
    # scale creatives by brand performance
    scale = max(m["cv7"] / 50, 0.3)
    crs = [
        (f"{name}_動画①", int(180000 * scale), int(25000 * scale), int(90000 * scale), max(int(80 * scale), 1), "残す"),
        (f"{name}_動画②", int(90000 * scale), int(12000 * scale), int(35000 * scale), max(int(20 * scale), 1), "残す"),
        (f"{name}_静止画A", int(8000 * scale), int(400 * scale), int(5000 * scale), max(int(3 * scale), 0), "残す"),
        (f"{name}_静止画B", int(2000 * scale), int(80 * scale), int(3000 * scale), 0, "止める"),
    ]
    for r_i, (ad, imp, clk, cost, cv, judge) in enumerate(crs, start=4):
        cpa = round(cost / cv) if cv else ""
        ws.cell(r_i, 1, ad)
        ws.cell(r_i, 2, yen(imp))
        ws.cell(r_i, 3, yen(clk))
        ws.cell(r_i, 4, yen(cost))
        ws.cell(r_i, 5, cv if cv else 0)
        ws.cell(r_i, 6, cpa)
        ws.cell(r_i, 7, judge)
        if isinstance(cpa, int):
            fill_cpa_cell(ws.cell(r_i, 6), cpa, target, allow)
    autosize(ws, {1: 20})


def build_summary(wb):
    ws = wb.active
    ws.title = "全体サマリ"
    ws["A1"] = "全体サマリ"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = (
        "直近7日は10/26〜11/1。前月は10/1〜10/31。CVは友だち追加（デモ）。"
        "目標比は（目標CPA−実績CPA）÷目標CPA。※数値はすべてダミーです。"
    )
    ws["A2"].fill = PatternFill("solid", fgColor=NOTE)
    headers = [
        "区分", "媒体", "キャンペーン", "日予算", "想定月額",
        "7日費用", "7日CV", "7日CPA", "7日目標比",
        "前月費用", "前月CV", "前月CPA", "前月目標比",
        "課題", "今日決めること",
        "当月費用", "当月CV", "当月CPA", "当月目標比",
    ]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    style_header(ws, 3, len(headers))

    r = 4
    for brand in BRANDS:
        for media, block, camp in (
            ("Google", brand["g"], brand["name"]),
            ("Meta", brand["m"], f"{brand['name']}リード"),
        ):
            target, allow = brand["target"], brand["allow"]
            ws.cell(r, 1, brand["group"])
            ws.cell(r, 2, media)
            ws.cell(r, 3, camp)
            day = brand["day_budget"] if media == "Google" else int(brand["day_budget"] * 1.1)
            ws.cell(r, 4, yen(day))
            ws.cell(r, 5, yen(day * 30))
            ws.cell(r, 6, yen(block["cost7"]))
            ws.cell(r, 7, block["cv7"])
            ws.cell(r, 8, block["cpa7"])
            ws.cell(r, 9, gap(target, block["cpa7"]))
            ws.cell(r, 10, yen(block["cost_m"]))
            ws.cell(r, 11, block["cv_m"])
            ws.cell(r, 12, block["cpa_m"])
            ws.cell(r, 13, gap(target, block["cpa_m"]))
            ws.cell(r, 14, block["issue"])
            ws.cell(r, 15, block["action"])
            ws.cell(r, 16, yen(block["cost_cur"]))
            ws.cell(r, 17, block["cv_cur"])
            ws.cell(r, 18, block["cpa_cur"])
            ws.cell(r, 19, gap(target, block["cpa_cur"]))
            for col, cpa in ((8, block["cpa7"]), (12, block["cpa_m"]), (18, block["cpa_cur"])):
                fill_cpa_cell(ws.cell(r, col), cpa, target, allow)
            r += 1

    # 認知行（CPA対象外）
    ws.cell(r, 1, "認知")
    ws.cell(r, 2, "Meta")
    ws.cell(r, 3, "ブランドB認知")
    ws.cell(r, 4, "5,000")
    ws.cell(r, 5, "150,000")
    ws.cell(r, 6, "28,000")
    ws.cell(r, 7, "—")
    ws.cell(r, 8, "")
    ws.cell(r, 9, "—")
    ws.cell(r, 14, "表示は出ている。CPAの対象外。")
    ws.cell(r, 15, "購入数が分かるまで据え置き")

    autosize(ws, {3: 16, 14: 36, 15: 28})


def build_awareness(wb):
    ws = wb.create_sheet("認知＿表示")
    ws["A1"] = "認知　表示"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "ブランドBのみ。日予算5,000円。CPAでは評価しない。直近30日の広告別。"
    for i, h in enumerate(["広告", "表示", "クリック", "費用", "友だち追加", "判断"], 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 6)
    rows = [
        ("認知_デザイン訴求", 76000, 69, 8500, 0, "据え置き"),
        ("認知_イベント訴求", 48000, 48, 5200, 1, "据え置き"),
        ("認知_手軽訴求", 21000, 27, 2300, 0, "据え置き"),
    ]
    for r_i, row in enumerate(rows, start=4):
        for c_i, v in enumerate(row, 1):
            if c_i in (2, 3, 4):
                ws.cell(r_i, c_i, yen(v) if isinstance(v, int) else v)
            else:
                ws.cell(r_i, c_i, v)
    ws.cell(8, 1, "課題")
    ws.cell(8, 2, "表示は出ている。友だち追加は1件。認知を広げるかは購入数を聞いてから決める。")
    autosize(ws, {1: 22, 2: 40})


def build_next_action(wb):
    ws = wb.create_sheet("Next Action")
    ws["A1"] = "Next Action"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "打ち合わせで決めたことだけを残す。済は1、未了は空欄。※デモ用の仮項目。"
    for i, h in enumerate(["施策", "済", "memo"], 1):
        ws.cell(3, i, h)
    style_header(ws, 3, 3)
    items = [
        ("ブランドBの許容超えクエリを除外するか決める", "", "除外する"),
        ("ブランドCの入札上限を見直すか決める", "", "見直す"),
        ("ブランドAの日予算を据え置くか決める", "", "据え置き"),
        ("認知の増額は購入数を聞いてから決める", "", "購入数が分かるまで配分しない"),
        ("来週の定例で7日CPAを再確認する", "", "色分け表で見る"),
    ]
    for r_i, row in enumerate(items, start=4):
        for c_i, v in enumerate(row, 1):
            ws.cell(r_i, c_i, v)
    autosize(ws, {1: 42, 3: 36})


def build_legend(wb):
    ws = wb.create_sheet("凡例")
    ws["A1"] = "凡例・注意"
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(["色", "条件", "説明"])
    style_header(ws, 3, 3)
    rows = [
        ("緑", "実績CPA ≤ 目標CPA", "目標内。増額や維持の候補"),
        ("黄", "目標 < 実績 ≤ 許容", "許容内。改善しながら継続"),
        ("赤", "実績 > 許容", "赤字ライン超え。除外・停止・入札見直し"),
        ("（なし）", "目標またはCVが空", "認知などCPA対象外"),
    ]
    for r_i, (color, cond, note) in enumerate(rows, start=4):
        ws.cell(r_i, 1, color)
        ws.cell(r_i, 2, cond)
        ws.cell(r_i, 3, note)
        if color == "緑":
            ws.cell(r_i, 1).fill = PatternFill("solid", fgColor=GREEN)
        elif color == "黄":
            ws.cell(r_i, 1).fill = PatternFill("solid", fgColor=YELLOW)
        elif color == "赤":
            ws.cell(r_i, 1).fill = PatternFill("solid", fgColor=RED)
    ws["A9"] = "タブ構成は本番の定例報告と同じ型です。Googleは成果・入札・クエリ・KW・広告文、Metaは成果・CR。"
    ws["A10"] = "※このファイルのブランド名・数値はすべてダミーです。本番のクライアント名・実数値ではありません。"
    ws["A10"].font = Font(bold=True, color="990000")
    autosize(ws, {2: 28, 3: 40})


def main():
    wb = Workbook()
    build_summary(wb)
    for brand in BRANDS:
        add_google_campaign(wb, brand)
        add_meta_campaign(wb, brand)
    build_awareness(wb)
    build_next_action(wb)
    build_legend(wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(OUT)
    print(f"sheets={len(wb.sheetnames)}")
    for name in wb.sheetnames:
        print(" -", name)


if __name__ == "__main__":
    main()
