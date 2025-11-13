# GitHub Actions 変換ツール

ローカルのメモリが不足する場合、GitHub Actionsを利用して大容量ファイルを処理します。

## 使い方

### 1. PPTXファイルをアップロード
PPTXファイルをファイル共有サービス（Google Drive、Dropbox等）にアップロードし、公開ダウンロードURLを取得します。
Googleスライドの場合は、共有リンク（`https://docs.google.com/presentation/d/.../edit` など）をそのまま指定すれば、自動的に PPTX ダウンロードリンクへ変換されます。

### 2. ワークフローを実行
1. リポジトリの **Actions** タブに移動
2. **PowerPoint Converter** ワークフローを選択
3. **Run workflow** をクリック
4. 以下を入力：
   - **File URL**: PPTXファイルのダウンロードリンク
   - **DPI**: サムネイル品質（50-150、デフォルト: 100）
5. **Run workflow** をクリック

### 3. 結果をダウンロード
- ワークフロー完了を待ちます（通常5-15分）
- ワークフローのアーティファクトから変換ファイルをダウンロード（保存期間：1日）

## メリット

- **大容量メモリ**: 7GB RAM（ローカルの512MB制限を回避）
- **高速処理**: 2コアCPUで変換に最適化
- **セットアップ不要**: ローカルインストール不要
- **自動クリーンアップ**: 1日後にファイル自動削除

## 利用可能なファイルサービス

ファイルホスティングにおすすめのサービス：
- **Google Drive**: 共有 > リンクを取得 > リンクを知っている全員
- **Dropbox**: 公開アクセスで共有リンク作成
- **GitHub Releases**: リリースアセットとしてアップロード
- **一時ファイルサービス**: WeTransferなど

## 制限事項

- **ファイルサイズ**: ワークフローあたり最大2GB
- **処理時間**: 最大30分
- **保存期間**: 1日後にファイル削除
- **公開URL**: ファイルが公開アクセス可能である必要

## URLの例

```
Google Drive: https://drive.google.com/uc?export=download&id=FILE_ID
Dropbox: https://dl.dropboxusercontent.com/s/FILE_ID/filename.pptx
直接リンク: https://example.com/path/to/file.pptx
```
