# 粗利分析ダッシュボード

毎月の粗利分析Excelファイルをアップロードするだけでダッシュボードを更新できるWebアプリです。

## 🚀 使い方

### 1. GitHub Pages でホスティング（推奨）

```bash
# リポジトリをクローン
git clone https://github.com/YOUR_USERNAME/gross-profit-dashboard.git
cd gross-profit-dashboard

# GitHub Pages を有効化
# Settings → Pages → Source: main branch / root
```

アクセスURL: `https://YOUR_USERNAME.github.io/gross-profit-dashboard/`

### 2. 毎月のデータ更新（ブラウザから）

1. ダッシュボードを開く
2. 右上の「**データ更新**」ボタンをクリック
3. 新しい `粗利分析_YYYYMM_*.xlsx` をアップロード
4. 「更新する」をクリック → ダッシュボードが即時更新

> ⚠️ ブラウザ上での更新は一時的です（リロードで元に戻ります）。  
> 恒久的に更新したい場合は「スクリプトで埋め込みデータ更新」を使用してください。

---

## 📁 ファイル構成

```
/
├── index.html          # ダッシュボード本体（全自己完結）
├── update_data.py      # Excelを読み込んでindex.htmlのデータを更新するスクリプト
└── README.md
```

---

## 🔄 恒久的なデータ更新（推奨ワークフロー）

毎月の更新はPythonスクリプトで行います：

```bash
# 必要なライブラリ
pip install openpyxl

# データ更新
python update_data.py 粗利分析_202504_1.xlsx

# Git にコミット & プッシュ
git add index.html
git commit -m "データ更新: 2025年4月"
git push origin main
```

GitHub Pages が自動的にデプロイされ、ダッシュボードが更新されます。

---

## 📊 ダッシュボードの内容

| セクション | 内容 |
|------------|------|
| KPIカード | 上半期粗利合計、Q1/Q2粗利、前年増減率 |
| 四半期比較 | 部門別 前年同期比棒グラフ |
| カテゴリ構成 | 上半期 ドーナツチャート |
| 粗利率推移 | 主要部門の粗利率比較 |
| 部門別詳細テーブル | 全部門の Q1/Q2 前年比較 |
| コミッション計算 | 大家さん・佐々木さん |

---

## ⚙️ Excelファイルの要件

以下のシート構成のxlsxファイルに対応しています：

- **粗利分析サマリー**: 部門別 Q1/Q2 粗利・前年比
- **部門別明細**: 売上・原価・粗利の詳細データ

---

## 🌐 外部依存

- [Chart.js 4.4.1](https://www.chartjs.org/) — グラフ描画
- [SheetJS (xlsx) 0.18.5](https://sheetjs.com/) — Excelファイル読み込み
- [Google Fonts](https://fonts.google.com/) — Noto Sans JP, DM Mono
