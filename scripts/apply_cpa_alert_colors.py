#!/usr/bin/env python3
"""CPAアラート色をスプシに適用する。

緑: 実績CPA ≤ 目標CPA
黄: 目標CPA < 実績CPA ≤ 許容CPA
赤: 実績CPA > 許容CPA
目標またはCPAが空（認知など）は色なし。

対象:
- Gravity Craft 日次データ（自動更新）の「目標」「月次サマリ」（Looker元）
- 2026.1024定例_広告報告の「全体サマリ」と各成果タブ
"""
from __future__ import annotations

import json
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

DAILY_ID = "1QECFlIPxtbw97XBSk0-eY1dnOnbz0Qt5g_zggS4CADg"
REPORT_ID = "1JiGj1YNLmBVQ6suSqGWlQqN97sfh2JfmCV4R6PH0aqE"

# CPキー → (目標CPA, 許容CPA)
THRESHOLDS = {
    "MEGA_Google": (3000, 8095),
    "MEGA_Meta": (3000, 8095),
    "MEGAオリ缶_Google": (3000, 8095),
    "つくーる_Google": (2400, 6536),
    "つくーる_Meta": (2400, 6536),
    "つくーるオリ缶_Google": (2400, 6536),
    "メンズ_Google": (2400, 6536),
    "メンズ_Meta": (2400, 6536),
    "コンカフェ_Google": (2400, 6536),
    "コンカフェ_Meta": (2400, 6536),
    # P-MAXもブランド目標を流用（空のときは色なしのまま）
    "MEGAオリ缶_Google_PMAX": (3000, 8095),
    "つくーるオリ缶_Google_PMAX": (2400, 6536),
}

REPORT_ROW_THRESHOLDS = {
    # 全体サマリの行番号（1-indexed）→ (目標, 許容)
    4: (3000, 8095),   # MEGAオリシャン
    5: (2400, 6536),   # つくーるオリシャン
    6: (2400, 6536),   # メンズ
    7: (2400, 6536),   # コンカフェ
    8: (3000, 8095),   # MEGAオリ缶
    9: (2400, 6536),   # つくーるオリ缶
    10: (3000, 8095),  # MEGAリード
    11: (2400, 6536),  # つくーるリード
    12: (2400, 6536),  # メンズリード
    13: (2400, 6536),  # コンカフェリード
    14: (3000, 8095),  # MEGAオリ缶リード
    15: (2400, 6536),  # つくーるオリ缶リード
}

GREEN = {"red": 0.851, "green": 0.918, "blue": 0.827}
YELLOW = {"red": 1.0, "green": 0.949, "blue": 0.8}
RED = {"red": 0.957, "green": 0.8, "blue": 0.8}
WHITE = {"red": 1, "green": 1, "blue": 1}


def sheets():
    token = json.loads(Path("/Users/nishi/Cursol/mcp/google_sheets_token.json").read_text())
    creds = Credentials.from_authorized_user_info(token, token["scopes"])
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        Path("/Users/nishi/Cursol/mcp/google_sheets_token.json").write_text(creds.to_json())
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def parse_number(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").replace("円", "").strip()
    if not text or text in ("—", "-", "—"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def band(cpa, target, allow):
    if cpa is None or target is None or allow is None:
        return None
    if cpa <= target:
        return "緑"
    if cpa <= allow:
        return "黄"
    return "赤"


def color_for(label):
    if label == "緑":
        return GREEN
    if label == "黄":
        return YELLOW
    if label == "赤":
        return RED
    return WHITE


def sheet_ids(service, spreadsheet_id):
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties"
    ).execute()
    return {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}


def clear_conditional(service, spreadsheet_id, sheet_id):
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title),conditionalFormats)",
    ).execute()
    requests = []
    for sheet in meta["sheets"]:
        if sheet["properties"]["sheetId"] != sheet_id:
            continue
        for i in range(len(sheet.get("conditionalFormats") or []) - 1, -1, -1):
            requests.append({
                "deleteConditionalFormatRule": {"sheetId": sheet_id, "index": i}
            })
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()


def update_daily(service):
    ids = sheet_ids(service, DAILY_ID)
    # 目標シート: 目標CPA + 許容CPA
    goal_rows = service.spreadsheets().values().get(
        spreadsheetId=DAILY_ID, range="'目標'!A1:C30"
    ).execute().get("values", [])
    out = [["CP", "目標CPA", "許容CPA", "当月予算"]]
    data_updates = []
    for i, row in enumerate(goal_rows[1:], start=2):
        if not row or not row[0]:
            continue
        cp = row[0]
        budget = row[2] if len(row) > 2 else ""
        target, allow = THRESHOLDS.get(cp, (None, None))
        # keep existing target if set and not in map
        existing_target = parse_number(row[1]) if len(row) > 1 else None
        if target is None:
            target = existing_target
        data_updates.append({
            "range": f"'目標'!A{i}:D{i}",
            "values": [[cp, target if target is not None else "", allow if allow is not None else "", budget]],
        })
    data_updates.insert(0, {"range": "'目標'!A1:D1", "values": [["CP", "目標CPA", "許容CPA", "当月予算"]]})

    # 月次サマリ: 目標埋め + 許容CPA + CPA帯 + 条件付き書式
    summary = service.spreadsheets().values().get(
        spreadsheetId=DAILY_ID, range="'月次サマリ'!A1:Z40"
    ).execute().get("values", [])
    if not summary:
        raise SystemExit("月次サマリが空です")
    header = summary[0]
    # Ensure columns
    while len(header) < 25:
        header.append("")
    # Put 許容CPA / CPA帯 at X/Y (index 23/24) to avoid shifting existing cols
    data_updates.append({
        "range": "'月次サマリ'!X1:Z1",
        "values": [["許容CPA", "CPA帯", "CPAスコア"]],
    })

    monthly_id = ids["月次サマリ"]
    for r_i, row in enumerate(summary[1:], start=2):
        if not row or not row[0]:
            continue
        cp = row[0]
        target, _allow = THRESHOLDS.get(cp, (None, None))
        if target is not None:
            data_updates.append({"range": f"'月次サマリ'!F{r_i}", "values": [[target]]})
        # 許容は目標シートから、帯・スコアは当月CPAと目標・許容から自動算出（Looker元）
        data_updates.append({
            "range": f"'月次サマリ'!X{r_i}:Z{r_i}",
            "values": [[
                f"=IFERROR(VLOOKUP(A{r_i},'目標'!A:C,3,FALSE),\"\")",
                f'=IF(OR(E{r_i}="",F{r_i}="",X{r_i}=""),"",IF(E{r_i}<=F{r_i},"緑",IF(E{r_i}<=X{r_i},"黄","赤")))',
                f'=IF(Y{r_i}="緑",1,IF(Y{r_i}="黄",2,IF(Y{r_i}="赤",3,"")))',
            ]],
        })

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=DAILY_ID,
        body={"valueInputOption": "USER_ENTERED", "data": data_updates},
    ).execute()

    # Durable CF on 当月_CPA using 目標/許容 columns
    clear_conditional(service, DAILY_ID, monthly_id)
    last_row = len(summary)
    # E=4, F=5, X=23 — CF custom formulas with X need column X present
    rules = []
    # Green / yellow / red for 当月_CPA (E) using F and X
    for idx, (formula, color) in enumerate([
        ('=AND($E2<>"",$F2<>"",$E2<=$F2)', GREEN),
        ('=AND($E2<>"",$F2<>"",$X2<>"",$E2>$F2,$E2<=$X2)', YELLOW),
        ('=AND($E2<>"",$X2<>"",$E2>$X2)', RED),
    ]):
        rules.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": monthly_id,
                        "startRowIndex": 1,
                        "endRowIndex": last_row,
                        "startColumnIndex": 4,
                        "endColumnIndex": 5,
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": formula}],
                        },
                        "format": {"backgroundColor": color},
                    },
                },
                "index": idx,
            }
        })
    for idx, (formula, color) in enumerate([
        ('=AND($J2<>"",$F2<>"",$J2<=$F2)', GREEN),
        ('=AND($J2<>"",$F2<>"",$X2<>"",$J2>$F2,$J2<=$X2)', YELLOW),
        ('=AND($J2<>"",$X2<>"",$J2>$X2)', RED),
    ], start=3):
        rules.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": monthly_id,
                        "startRowIndex": 1,
                        "endRowIndex": last_row,
                        "startColumnIndex": 9,
                        "endColumnIndex": 10,
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": formula}],
                        },
                        "format": {"backgroundColor": color},
                    },
                },
                "index": idx,
            }
        })
    # CPA帯 text equals
    for idx, (text, color) in enumerate([
        ("緑", GREEN), ("黄", YELLOW), ("赤", RED)
    ], start=6):
        rules.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": monthly_id,
                        "startRowIndex": 1,
                        "endRowIndex": last_row,
                        "startColumnIndex": 24,
                        "endColumnIndex": 25,
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "TEXT_EQ",
                            "values": [{"userEnteredValue": text}],
                        },
                        "format": {"backgroundColor": color},
                    },
                },
                "index": idx,
            }
        })
    service.spreadsheets().batchUpdate(
        spreadsheetId=DAILY_ID, body={"requests": rules}
    ).execute()
    print(f"日次データ: 目標・月次サマリを更新（{last_row - 1}行）")


def update_report(service):
    ids = sheet_ids(service, REPORT_ID)
    summary_id = ids["全体サマリ"]
    values = service.spreadsheets().values().get(
        spreadsheetId=REPORT_ID, range="'全体サマリ'!A1:S20"
    ).execute().get("values", [])
    requests = []
    # CPA columns: H=7日CPA(7), L=10月CPA(11), R=当月CPA if present(17)
    for row_no, (target, allow) in REPORT_ROW_THRESHOLDS.items():
        if row_no > len(values):
            continue
        row = values[row_no - 1]
        for col_idx in (7, 11, 17):  # H, L, R
            if col_idx >= len(row):
                continue
            cpa = parse_number(row[col_idx])
            label = band(cpa, target, allow)
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": summary_id,
                        "startRowIndex": row_no - 1,
                        "endRowIndex": row_no,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1,
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": color_for(label)}},
                    "fields": "userEnteredFormat.backgroundColor",
                }
            })

    # 成果タブ: D列CPA（直近7日・月次・当月の行）
    result_tabs = [
        ("G_MEGAオリシャン＿成果", 3000, 8095),
        ("G_つくーるオリシャン＿成果", 2400, 6536),
        ("G_メンズ＿成果", 2400, 6536),
        ("G_コンカフェ＿成果", 2400, 6536),
        ("G_MEGAオリ缶＿成果", 3000, 8095),
        ("G_つくーるオリ缶＿成果", 2400, 6536),
        ("M_MEGAリード＿成果", 3000, 8095),
        ("M_つくーるリード＿成果", 2400, 6536),
        ("M_メンズリード＿成果", 2400, 6536),
        ("M_コンカフェリード＿成果", 2400, 6536),
    ]
    for title, target, allow in result_tabs:
        if title not in ids:
            continue
        tab = service.spreadsheets().values().get(
            spreadsheetId=REPORT_ID, range=f"'{title}'!A1:F10"
        ).execute().get("values", [])
        for r_i, row in enumerate(tab):
            if r_i < 3:  # header area
                continue
            if not row or row[0] in ("課題", "今日決めること"):
                continue
            cpa = parse_number(row[3] if len(row) > 3 else None)
            label = band(cpa, target, allow)
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": ids[title],
                        "startRowIndex": r_i,
                        "endRowIndex": r_i + 1,
                        "startColumnIndex": 3,
                        "endColumnIndex": 4,
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": color_for(label)}},
                    "fields": "userEnteredFormat.backgroundColor",
                }
            })

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=REPORT_ID, body={"requests": requests}
        ).execute()
    print(f"定例報告: {len(requests)}セルを着色")


def main():
    service = sheets()
    update_daily(service)
    update_report(service)
    print("完了。緑=目標以内 / 黄=目標超〜許容以内 / 赤=許容超")


if __name__ == "__main__":
    main()
