#!/usr/bin/env python3
"""Lookerの当月CVへ、Google広告にまだ出ていないLINE友だち追加を入れる。

日次シートの朝更新は広告のCVを書く。送信が遅れている日は0になるので、
Noroshiに残っている広告クリック由来の友だち追加で上書きする。
当月は1日〜昨日。離脱クリックは入れない。
"""
import json
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SHEET_ID = "1QECFlIPxtbw97XBSk0-eY1dnOnbz0Qt5g_zggS4CADg"
JST = ZoneInfo("Asia/Tokyo")
CAMPAIGN = {
    ("mega", "oricham"): "24125828809",
    ("mega", "orican"): "24370885902",
    ("tsuku", "oricham"): "24124161068",
    ("tsuku", "orican"): "24381978925",
    ("mens", "oricham"): "22267817744",
    ("cafe", "oricham"): "22261553689",
}
CP_OF = {cid: name for name, cid in {
    "MEGA_Google": "24125828809",
    "MEGAオリ缶_Google": "24370885902",
    "つくーる_Google": "24124161068",
    "つくーるオリ缶_Google": "24381978925",
    "メンズ_Google": "22267817744",
    "コンカフェ_Google": "22261553689",
}.items()}


def sheets():
    token = json.loads(Path("/Users/nishi/Cursol/mcp/google_sheets_token.json").read_text())
    creds = Credentials.from_authorized_user_info(token, token["scopes"])
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def noroshi_counts(start, end):
    start_utc = datetime.combine(start, datetime.min.time(), tzinfo=JST).astimezone(ZoneInfo("UTC"))
    end_utc = datetime.combine(end + timedelta(days=1), datetime.min.time(), tzinfo=JST).astimezone(ZoneInfo("UTC"))
    sql = (
        "SELECT CASE "
        "WHEN ifnull(cl.landing_page,'') || ' ' || ifnull(cl.referer,'') LIKE '%mega-oricham.com%' THEN 'mega' "
        "WHEN ifnull(cl.landing_page,'') || ' ' || ifnull(cl.referer,'') LIKE '%orichamtsuku-ru.com%' THEN 'tsuku' "
        "WHEN ifnull(cl.landing_page,'') || ' ' || ifnull(cl.referer,'') LIKE '%mens-oricham.com%' THEN 'mens' "
        "WHEN ifnull(cl.landing_page,'') || ' ' || ifnull(cl.referer,'') LIKE '%kawaii-oricham.com%' THEN 'cafe' "
        "ELSE 'other' END brand, "
        "CASE WHEN ifnull(cl.utm_campaign,'') = 'orican' THEN 'orican' ELSE 'oricham' END camp, "
        "substr(datetime(cv.converted_at, '+9 hours'), 1, 10) d, count(*) n "
        "FROM conversions cv JOIN clicks cl ON cl.id = cv.click_id "
        f"WHERE cv.converted_at >= '{start_utc.strftime('%Y-%m-%d %H:%M:%S')}' "
        f"AND cv.converted_at < '{end_utc.strftime('%Y-%m-%d %H:%M:%S')}' "
        "AND cl.gclid IS NOT NULL AND cl.gclid != '' AND cl.gclid NOT LIKE 'test_%' "
        "GROUP BY 1, 2, 3"
    )
    proc = subprocess.run(
        ["npx", "wrangler", "d1", "execute", "noroshi-db", "--remote", "--json", "--command", sql],
        cwd="/Users/nishi/Cursol/mcp/noroshi",
        capture_output=True,
        text=True,
        timeout=90,
        check=False,
    )
    payload = json.loads(proc.stdout)
    counts = {}
    for row in payload[0]["results"]:
        cid = CAMPAIGN.get((row.get("brand"), row.get("camp")))
        if not cid:
            continue
        counts[(row["d"], cid)] = int(row["n"])
    return counts


def main():
    today = datetime.now(JST).date()
    start = today.replace(day=1)
    end = today - timedelta(days=1)
    if end < start:
        print("no closed day in this month")
        return
    counts = noroshi_counts(start, end)
    svc = sheets()
    daily = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="'Google_日次'!A:K"
    ).execute().get("values", [])
    ads = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="'広告_日次'!A:I"
    ).execute().get("values", [])
    data = []
    month_cv = {cp: 0 for cp in CP_OF.values()}
    for i, row in enumerate(daily[1:], start=2):
        if len(row) < 10 or not str(row[0]).startswith(start.isoformat()[:8]):
            continue
        if row[0] < start.isoformat() or row[0] > end.isoformat():
            continue
        cid = str(row[3])
        cp = CP_OF.get(cid)
        if not cp:
            continue
        cv = counts.get((row[0], cid), 0)
        month_cv[cp] += cv
        if str(row[9]) != str(cv):
            data.append({"range": f"'Google_日次'!J{i}", "values": [[cv]]})
    for i, row in enumerate(ads[1:], start=2):
        if len(row) < 8 or row[1] != "Google":
            continue
        if row[0] < start.isoformat() or row[0] > end.isoformat():
            continue
        cid = next((c for c, name in CP_OF.items() if name == row[2]), None)
        if not cid:
            continue
        cv = counts.get((row[0], cid), 0)
        if str(row[7]) != str(cv):
            data.append({"range": f"'広告_日次'!H{i}", "values": [[cv]]})
    summary = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="'月次サマリ'!A:L"
    ).execute().get("values", [])
    for i, row in enumerate(summary[1:], start=2):
        cp = row[0] if row else ""
        if cp not in month_cv:
            continue
        cv = month_cv[cp]
        cost = float(row[2] or 0)
        target = float(row[5]) if len(row) > 5 and str(row[5]) not in ("",) else None
        prev = float(row[8]) if len(row) > 8 and str(row[8]) not in ("",) else 0
        cpa = round(cost / cv) if cv else ""
        gap = round((target - cpa) / target, 10) if target and cpa != "" else ""
        ratio = round((cv - prev) / prev, 10) if prev else ""
        data.append({"range": f"'月次サマリ'!D{i}:G{i}", "values": [[cv, cpa, row[5] if len(row) > 5 else "", gap]]})
        data.append({"range": f"'月次サマリ'!L{i}", "values": [[ratio]]})
    data.extend(october_cells(svc))
    if data:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"valueInputOption": "RAW", "data": data},
        ).execute()
    print(f"patched {start}..{end} updates={len(data)} cv={month_cv}")


# 10月は締め済み。Googleは日次のNoroshi合計。MetaはリードのLINE友だち追加（Noroshi）。
OCTOBER = {
    "MEGA_Google": (966489, 241.26),
    "つくーる_Google": (892117, 172),
    "メンズ_Google": (578146, 186),
    "コンカフェ_Google": (370376, 75),
    "MEGAオリ缶_Google": (36004, 1.5),
    "つくーるオリ缶_Google": (36111, 5),
    "MEGA_Meta": (963272, 410),
    "つくーる_Meta": (971555, 488),
    "メンズ_Meta": (275823, 106),
    "コンカフェ_Meta": (279786, 82),
}


def october_cells(svc):
    summary = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="'月次サマリ'!A1:Z15"
    ).execute().get("values", [])
    header = summary[0] if summary else []
    if "10月_CV" not in header:
        start_col = len(header) + 1
        header_range = f"'月次サマリ'!{col_letter(start_col)}1:{col_letter(start_col + 2)}1"
        cells = [{"range": header_range, "values": [["10月_費用", "10月_CV", "10月_CPA"]]}]
        cv_col = start_col + 1
        cost_col = start_col
        cpa_col = start_col + 2
    else:
        cost_col = header.index("10月_費用") + 1
        cv_col = header.index("10月_CV") + 1
        cpa_col = header.index("10月_CPA") + 1
        cells = []
    for i, row in enumerate(summary[1:], start=2):
        cp = row[0] if row else ""
        if cp not in OCTOBER:
            continue
        cost, cv = OCTOBER[cp]
        cpa = round(cost / cv) if cv else ""
        cells.append({"range": f"'月次サマリ'!{col_letter(cost_col)}{i}", "values": [[cost]]})
        cells.append({"range": f"'月次サマリ'!{col_letter(cv_col)}{i}", "values": [[cv]]})
        cells.append({"range": f"'月次サマリ'!{col_letter(cpa_col)}{i}", "values": [[cpa]]})
    return cells


def col_letter(n):
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


if __name__ == "__main__":
    main()
