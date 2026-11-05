# looker-sheets-cpa-alert

Looker元のGoogleスプレッドシートと定例報告に、**目標CPA／許容CPAの3色アラート（緑・黄・赤）**を載せる自動化です。Cursor用スキルと、営業共有用のダミー表を同梱しています。

## 色の定義

| 色 | 条件 |
|---|---|
| 緑 | 実績CPA ≤ 目標CPA |
| 黄 | 目標CPA < 実績CPA ≤ 許容CPA |
| 赤 | 実績CPA > 許容CPA |

## 中身

| パス | 説明 |
|---|---|
| `SKILL.md` | Cursor Agent Skill（手順の正本） |
| `reference.md` | シートID・数式・Looker手動設定 |
| `scripts/apply_cpa_alert_colors.py` | 本番スプシへ色と帯列を反映 |
| `scripts/patch_looker_google_cv.py` | Google CVをNoroshi友だち追加で補正 |
| `scripts/build_sales_demo.py` | 営業用ダミーxlsxを生成 |
| `demo/sales-demo-cpa-alert.xlsx` | 生成済みダミー表 |

## 営業用ダミーの作り方

```bash
cd looker-sheets-cpa-alert
uv run --no-project --with openpyxl python scripts/build_sales_demo.py
```

`demo/sales-demo-cpa-alert.xlsx` ができます。定例報告と同じ型（全体サマリ・入札・クエリ・KW・広告文・Meta CR など25タブ）のGoogle表は次です。

https://docs.google.com/spreadsheets/d/1UXHK4jtI7GvKmCf2G2CKZaWf85P8-OR9qGnECQDcJsk/edit

**数値はすべて架空です。本番のクライアント名・実数値ではありません。**  
リンク共有はオーナーが「リンクを知っている全員」に一度切り替えてください。

## Cursor Skill

個人スキルとしても置けます。

```bash
cp -R . ~/.cursor/skills/looker-sheets-cpa-alert
```

またはこのリポジトリを `.cursor/skills/looker-sheets-cpa-alert` としてプロジェクトに置く。

## 本番反映（Gravity Craft）

シートIDとトークンは `reference.md` を参照。実行は案件リポジトリ側のラッパを使います。

```bash
python scripts/apply_cpa_alert_colors.py
./scripts/patch_looker_google_cv.sh
```

## License

Private / internal use unless otherwise noted.
