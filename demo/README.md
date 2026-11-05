# 営業用デモ（ダミー数値）

本番の定例報告（全体サマリ＋キャンペーン別の成果・入札・クエリ・KW・広告文／Meta成果・CR）と同じ型のダミーです。

## Googleスプレッドシート（作成済み）

https://docs.google.com/spreadsheets/d/1UXHK4jtI7GvKmCf2G2CKZaWf85P8-OR9qGnECQDcJsk/edit

共有設定（「リンクを知っている全員」）はオーナー側で一度付けてください。

## タブ構成（25）

| 区分 | タブ |
|---|---|
| 全体 | `全体サマリ` |
| Google×3ブランド | `G_ブランド*＿成果` / `＿入札` / `＿クエリ` / `＿KW` / `＿広告文` |
| Meta×3ブランド | `M_ブランド*リード＿成果` / `＿CR` |
| その他 | `認知＿表示` / `Next Action` / `凡例` |

- **ブランドA** … 目標内（緑）
- **ブランドB** … 目標超〜許容内（黄）
- **ブランドC** … 許容超（赤）

ブランド名・費用・CVはすべて架空です。本番のクライアント名・実数値ではありません。

## 再生成

```bash
python scripts/build_sales_demo.py
```

xlsx は `demo/sales-demo-cpa-alert.xlsx` に出力されます。
