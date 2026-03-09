# ビジネス書類データベース

見積書・請求書・納品書・発注書・領収書をAIで自動抽出・管理するWebアプリです。

## 機能

- 📄 **5種類の書類対応**: 見積書・請求書・納品書・発注書・領収書
- 🤖 **AI自動抽出**: PDF/Excel/Word/画像から品目・金額を自動読み取り
- 📈 **相場価格検索**: 品目名からAIが日本市場の相場を推定
- 📊 **ダッシュボード**: 月別推移・取引先分析・グラフ表示
- 📥 **CSVエクスポート**: 書類一覧・品目一覧をCSV出力
- ✏️ **見積書自動作成**: 品目を入力するとAIが見積書を生成

## セットアップ

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."
uvicorn main:app --reload --port 8000
```

## Railway デプロイ

1. このリポジトリをGitHubにプッシュ
2. railway.app でGitHub連携してデプロイ
3. 環境変数 `ANTHROPIC_API_KEY` を設定

## ファイル構成

```
├── main.py          # FastAPI サーバ・APIエンドポイント
├── database.py      # SQLiteデータベース定義
├── extractor.py     # AI抽出エンジン・相場検索・見積生成
├── requirements.txt # 依存ライブラリ
├── Procfile         # Railway起動設定
└── static/
    └── index.html   # フロントエンド（全機能）
```
