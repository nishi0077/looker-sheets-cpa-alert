# 営業用デモ（ダミー数値）

`sales-demo-cpa-alert.xlsx` は、CPAアラートの見え方を説明するだけの架空データです。

- ブランドA〜D は架空名
- 費用・CV・CPA はデモ用に作った数字
- 緑／黄／赤／色なしが1枚で揃うように配置

## シート

1. **月次サマリ** … 当月CPAに色。`CPA帯` / `CPAスコア` 列あり（Looker接続の見本）
2. **目標** … 目標CPA・許容CPA
3. **凡例** … 色の意味

## 再生成

```bash
python scripts/build_sales_demo.py
```

共有時は必ず「ダミー数値です」と添える。
