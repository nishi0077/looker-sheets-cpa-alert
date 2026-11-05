---
name: looker-sheets-cpa-alert
description: >-
  Looker元スプシと定例報告に、目標CPA／許容CPAの3色アラート（緑・黄・赤）を反映し、Google CVのNoroshi補正も回す。Use when the user asks to update Looker CPA colors, apply 許容CPA alerts, refresh 月次サマリ band columns, run apply_cpa_alert_colors, patch Looker Google CV, or build a sales demo sheet with dummy CPA colors.
---

# Looker・スプシ CPAアラート自動化

広告の実績CPAを、目標CPAと許容CPAで色分けしてスプシ／Looker元データに載せる。

## 色の定義（固定）

| 色 | 条件 |
|---|---|
| 緑 | 実績CPA ≤ 目標CPA |
| 黄 | 目標CPA < 実績CPA ≤ 許容CPA |
| 赤 | 実績CPA > 許容CPA |
| なし | 目標・許容・CPAのいずれかが空（認知・CV0など） |

## いつ使うか

- 「Looker／スプシのCPAを緑黄赤にして」
- 「目標・許容を月次サマリに反映して」
- 「LookerのGoogle CVが遅れている／Noroshiで補正して」
- 「営業用にダミーのCPA色分けを見せたい」

## 手順（本番・Gravity Craft）

1. 閾値を確認する。正本は日次スプシの「目標」シート（目標CPA・許容CPA）。変更したら `scripts/apply_cpa_alert_colors.py` の `THRESHOLDS` も揃える。
2. CPAアラートを適用する。

```bash
cd "/Users/nishi/Cursol/Gravity Craft"
uv run --no-project --with google-auth --with google-auth-httplib2 --with google-api-python-client \
  python scripts/apply_cpa_alert_colors.py
```

3. GoogleのCVが媒体より遅れているときは、Noroshi補正を回す。

```bash
./scripts/patch_looker_google_cv.sh
```

4. Looker本体の色ルールはAPIで書けない。元スプシの `CPA帯` / `CPAスコア` をデータソースに足し、表の条件付き書式を一度入れる（詳細は [reference.md](reference.md)）。

5. 月曜定例の着色は `monday_ad_report.py` の `cpa_band_color()` が同じ定義でCPAセルを塗る。TARGETS／ALLOWも揃える。

## 書き込む列（Looker元・月次サマリ）

既存列はずらさない。末尾に足す。

| 列 | 内容 |
|---|---|
| X 許容CPA | `VLOOKUP` で「目標」シートから取得 |
| Y CPA帯 | 緑／黄／赤の文言（Lookerの条件付き書式用） |
| Z CPAスコア | 1＝緑、2＝黄、3＝赤（カラースケール用） |

当月_CPA（E）・前月_CPA（J）には条件付き書式を付ける。

## 営業用ダミー

実案件の数字は出さない。架空ブランドの表を作って共有する。

```bash
cd "/Users/nishi/Cursol/looker-sheets-cpa-alert"
uv run --no-project --with openpyxl python scripts/build_sales_demo.py
# 全体サマリ＋Google（成果・入札・クエリ・KW・広告文）＋Meta（成果・CR）＋認知＋Next Action のxlsx
```

共有文の型は「ダミー数値です。本番のクライアント名や実数値ではありません。」を必ず付ける。定例本番シートと同じタブ型で見せる。

## 注意

- トークンは `/Users/nishi/Cursol/mcp/google_sheets_token.json`。秘密を出力しない。
- 入札・予算・配信のオンオフはこのスキルでは変えない。
- 本番シートIDは [reference.md](reference.md) を正本にする。
