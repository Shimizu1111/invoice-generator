# CLAUDE.md

## プロジェクト概要

請求書・見積書を Google Sheets テンプレートから自動生成するツール。
Web UI（`docs/`）と CLI（`node index.js`）の2つのインターフェースがある。

## 技術スタック

- フロントエンド: バニラ HTML/JS（`docs/index.html`, `docs/app.js`）
- バックエンド(CLI): Node.js（`googleapis` パッケージ）
- 外部API: Google Sheets API, Google Drive API, OpenAI API, GitHub API
- ホスティング: GitHub Pages（`docs/` フォルダから配信）

## 開発環境

- Node.js 必須（CLIのみ）
- Web UI はローカルサーバー経由で動作確認（`make open` → `http://localhost:8000`）
- `file://` で開くと Google API のスクリプトが初期化できずログインボタンが表示されない

## Make コマンド

```
make help       # コマンド一覧
make open       # ローカルサーバーで Web UI を起動（localhost:8000）
make deploy     # docs/ を commit & push して GitHub Pages にデプロイ
make install    # npm install
```

## デプロイ

- GitHub Pages: https://shimizu1111.github.io/invoice-generator/
- `docs/` フォルダが GitHub Pages のソース
- `make deploy` で `docs/` の変更を commit → push → 自動デプロイ
- リポジトリ: git@github.com:Shimizu1111/invoice-generator.git

## ファイル構成

- `docs/` - Web UI（GitHub Pages で配信）
  - `index.html` - メインHTML
  - `app.js` - 全ロジック（API連携、フォーム、AI入力）
- `src/` - CLI用モジュール
  - `auth.js` - Google OAuth 認証（CLI用、credentials.json / token.json）
  - `config.js` - テンプレートID、セルマッピング、マスターシート設定
  - `sheets.js` - Sheets/Drive API ラッパー
  - `estimate.js` - 見積書作成
  - `invoice.js` - 請求書作成
  - `gitUtils.js` - Git 履歴からの作業項目抽出
- `index.js` - CLI エントリポイント

## 重要な設定値

- テンプレート（見積書）: `1k4TzjsW6N1w5kuflUJJLZbpAtztehzUXzVqXBIhGcvs`
- テンプレート（請求書）: `1S2DGu4FCLDZwGZX0uS-vl4hEPtU7Y4SHnBPxiEAANTc`
- マスター管理シート: `1ufhqWGyI0to5fR5y50qHILaCl46ljuuabtFbQ6tWZz0`
- 会社情報シート: `1jcWUiz2JxHi_fzn-1h75LOGx50DAiF5uWJPvIWEVlus`
- Drive ルートフォルダ: `1nX6rrInrQ2mBQK_Y-J6AO5FkWfc2-fh-`
- OAuth Client ID: `707875244824-kuhb9drhcanafjnqrs7fk9n7l3kjkssc.apps.googleusercontent.com`

## 注意点

- 行挿入時は `inheritFromBefore: false` にすること（ヘッダー行のスタイルを継承しないため）
- 備考欄は `wrapStrategy: 'WRAP'` + `autoResizeDimensions` で行高さを自動調整すること
- Web UI と CLI で同じテンプレート・セルマッピングを使用しているが、定義が二重管理になっている（`src/config.js` と `docs/app.js`）
