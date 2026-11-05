# Looker・スプシ CPAアラート — 参照

## 本番シート（Gravity Craft）

| 用途 | ID |
|---|---|
| Looker元・日次データ | `1QECFlIPxtbw97XBSk0-eY1dnOnbz0Qt5g_zggS4CADg` |
| 定例報告 | `1JiGj1YNLmBVQ6suSqGWlQqN97sfh2JfmCV4R6PH0aqE` |
| Looker閲覧 | https://lookerstudio.google.com/reporting/1beb453d-3bbd-44ce-9acb-edd7734ad9cc/page/TlJ0C |

日次タブの要点: `目標` / `月次サマリ` / `広告_日次` / `Google_日次` / `Meta_日次` / `更新ログ`

## 目標・許容（2026-11-04）

| ブランド | 目標CPA | 許容CPA |
|---|---:|---:|
| MEGA（オリシャン・オリ缶） | 3,000 | 8,095 |
| つくーる・メンズ・コンカフェ | 2,400 | 6,536 |

式（10/22合意）:

- 粗利LTV = 3,100 × 6 = 18,600
- 許容CPA = 粗利LTV × 購入率
- 期待値 = 許容CPA ÷ 3
- 目標CPA = 期待値 × 1.10 を100円単位で切り上げ

## 月次サマリの数式

行 `i` 向け:

```
X: =IFERROR(VLOOKUP(A{i},'目標'!A:C,3,FALSE),"")
Y: =IF(OR(E{i}="",F{i}="",X{i}=""),"",IF(E{i}<=F{i},"緑",IF(E{i}<=X{i},"黄","赤")))
Z: =IF(Y{i}="緑",1,IF(Y{i}="黄",2,IF(Y{i}="赤",3,"")))
```

条件付き書式（当月_CPA = E）:

- `=AND($E2<>"",$F2<>"",$E2<=$F2)` → 緑 `#D9EAD3`
- `=AND($E2<>"",$F2<>"",$X2<>"",$E2>$F2,$E2<=$X2)` → 黄 `#FFF2CC`
- `=AND($E2<>"",$X2<>"",$E2>$X2)` → 赤 `#F4CCCC`

## スクリプト

| ファイル | 役割 |
|---|---|
| `scripts/apply_cpa_alert_colors.py` | 目標・許容・帯列・条件付き書式・定例CPA着色 |
| `scripts/patch_looker_google_cv.py` | Google CVをNoroshi友だち追加で当月補正 |
| `scripts/patch_looker_google_cv.sh` | 上記の実行ラッパ |
| `scripts/build_sales_demo.py` | 営業用ダミーxlsx生成 |
| `monday_ad_report.py`（Gravity Craft側） | 月曜更新時も同じ3色でCPAセルを塗る |

## Looker Studio（手動・初回のみ）

1. データソースを更新し、`CPA帯` または `CPAスコア` を追加する
2. 表の「当月_CPA」または行に条件付き書式
   - CPA帯＝緑 → 背景 `#D9EAD3`
   - CPA帯＝黄 → 背景 `#FFF2CC`
   - CPA帯＝赤 → 背景 `#F4CCCC`
   - または CPAスコアのカラースケール（1→緑、3→赤）

## 再実行の目安

- 目標／許容を変えたあと → `apply_cpa_alert_colors.py`
- 朝の日次更新後にGoogle CVが0のまま → `patch_looker_google_cv.sh`
- 定例シートの色が古い → `apply_cpa_alert_colors.py` または月曜レポート実行

## 公開リポジトリ・営業デモ

- GitHub: https://github.com/nishi0077/looker-sheets-cpa-alert
- ダミーGoogle表: https://docs.google.com/spreadsheets/d/1VJRdjbd1YbFu1qWLaxgrETyVtn-ccgEyvGMb6dcO7cg/edit
