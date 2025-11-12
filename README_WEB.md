# PPTX Script Slides - Web API (Render.com)

PowerPoint の「ノート欄」に書かれた文字を、読み上げ用の大きな黒背景スライドに作り直すWeb APIです。Render.comで簡単にデプロイできます。

---

## 🚀 Render.comでのデプロイ方法

### 1. **GitHubリポジトリにプッシュ**
```bash
git add .
git commit -m "Add web API for Render.com deployment"
git push origin main
```

### 2. **Render.comで新規Webサービスを作成**
1. Render.comにログイン
2. 「New +」→「Web Service」を選択
3. GitHubリポジトリを接続
4. 以下の設定を行う：
   - **Environment**: Docker
   - **Branch**: main
   - **Root Directory**: （空のまま）
   - **Dockerfile Path**: ./Dockerfile
   - **Plan**: Free（または必要に応じてStarter）

### 3. **環境変数の設定**
- `PORT`: 8000（Render.comが自動設定）

---

## 📡 APIエンドポイント

### ヘルスチェック
```
GET /
```
レスポンス：
```json
{"status": "healthy", "service": "PPTX Script Slides API"}
```

### PPTX変換
```
POST /convert
Content-Type: multipart/form-data
```
パラメータ：
- `file`: PowerPointファイル (.pptx)

レスポンス：
```json
{
  "task_id": "uuid-string",
  "status": "processing",
  "message": "Conversion started. Check status endpoint.",
  "download_url": null
}
```

### ステータス確認
```
GET /status/{task_id}
```
レスポンス：
```json
{
  "task_id": "uuid-string",
  "status": "completed",
  "message": "Conversion completed successfully!",
  "download_url": "/download/uuid-string"
}
```

### ファイルダウンロード
```
GET /download/{task_id}
```
変換されたPowerPointファイルをダウンロード

### クリーンアップ
```
DELETE /cleanup/{task_id}
```
タスク関連ファイルを削除

---

## 🔧 ローカル開発

### 1. **依存関係のインストール**
```bash
pip install -r requirements_web.txt
```

> **Poppler について**: `pdf2image` を利用するため、Poppler バイナリが必要です。Render.com の Dockerfile では `poppler-utils` をインストール済みです。ローカルでは `brew install poppler` (macOS) などで追加してください。

### 2. **サーバーの起動**
```bash
python web_app.py
```

### 3. **APIテスト**
```bash
# 変換リクエスト
curl -X POST "http://localhost:8000/convert" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@test.pptx"

# ステータス確認
curl "http://localhost:8000/status/{task_id}"
```

---

## 📝 仕様

### サポートされるファイル
- **形式**: .pptx (PowerPoint 2007+)
- **最大サイズ**: Render.comのFreeプラン制限による（通常50MB）

### 変換機能
- ノート欄の文章を自動で約200文字ごとに分割
- 話者ごとに文字色を自動変更（話者1: 黄色 / 話者2: シアン / 話者3: 緑）
- 黒い背景・大きな文字（メイリオ 40pt ボールド）で表示
- 右下に元スライドのサムネイルを自動挿入
- ページ番号の自動付与（分割された場合）

### 制限事項
- FreeプランではメモリとCPUに制限あり
- 大きなファイルの処理には時間がかかる場合あり
- 同時処理数はサーバースペックに依存

---

## 🛠️ カスタマイズ

### 環境変数
- `PORT`: サーバーポート（デフォルト: 8000）

### Dockerカスタマイズ
`Dockerfile`を編集して以下が可能：
- Pythonバージョンの変更
- 追加のシステム依存関係
- メモリ制限の調整

---

## 📊 モニタリング

Render.comのダッシュボードで以下を監視：
- サーバーの状態
- リクエストログ
- リソース使用量
- エラーレート

---

## 🔒 セキュリティ

- ファイルアップロードの種類制限（.pptxのみ）
- 一時ファイルの自動クリーンアップ
- CORS設定（本番環境では適切に設定）

---

## 📞 サポート

問題が発生した場合：
1. Render.comのログを確認
2. GitHub Issuesで報告
3. APIレスポンスの詳細を確認
