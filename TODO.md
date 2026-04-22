# 請求書・見積書自動作成ツール

## 概要
Google Sheets のテンプレートを元に、請求書・見積書を自動作成するツール。
Claude Code から自然言語で指示して作成する。

## Google Sheets テンプレート
- 見積書: https://docs.google.com/spreadsheets/d/1CJGUEnZk2nuwxXDKXO__EGlDZ5bAuY9ErFEJqxf9Q9E/edit?usp=drivesdk
- 請求書: https://docs.google.com/spreadsheets/d/1Ln2FNWsLLnTrazkgKJobwXA3Cvwpz22YmI91zj0-fxo/edit?usp=drivesdk

## やること

### 1. Google Drive / Sheets API 認証セットアップ
- OAuth 認証の設定（既存の仕組みに合わせる）
- credentials.json の配置場所を確認

### 2. テンプレート解析
- 見積書テンプレートの構造・記載例を確認
- 請求書テンプレートの構造・記載例を確認
- どのセルにどの情報を入れるかマッピング

### 3. 見積書作成機能
- 入力ソース1: Git コミット履歴から作業内容を自動抽出して見積書を作成
- 入力ソース2: ユーザーからの明示的な指示内容で見積書を作成

### 4. 請求書作成機能
- 見積書ベース: 見積書のURLを紐付けて、見積書の内容をもとに請求書を作成
- 単独作成: 見積書なしで請求書のみ作成するパターンにも対応

### 5. プロジェクト構成
- Node.js プロジェクト（他プロジェクトと同様の構成）
- googleapis パッケージで Google Sheets API を操作
- GitHub にプッシュ

## 認証方式
- OAuth（ユーザーが既に使用中）

## インターフェース
- Claude Code から直接指示する方式（自然言語）
