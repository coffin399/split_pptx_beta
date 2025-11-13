# PowerPoint スクリプトスライド変換ツール（ローカル版）

簡単に使えるローカルWebアプリ版です。

## セットアップ

### macOSの場合
1. ターミナルを開く
2. このフォルダで実行：
   ```bash
   chmod +x install_macos.sh
   ./install_macos.sh
   ```

### Windowsの場合
1. コマンドプロンプトを管理者権限で開く
2. このフォルダで実行：
   ```cmd
   install_windows.bat
   ```

## 起動方法

### 共通
1. ターミナル/コマンドプロンプトでこのフォルダに移動
2. 実行：
   ```bash
   python web_app.py
   ```
3. ブラウザで開く：http://localhost:8000

## macOS非署名対応について
このツールはシステム環境を変更せず、LibreOfficeのquarantine属性のみを削除します：
- LibreOfficeの quarantine属性を削除
- システムのセキュリティ設定は変更しない

**注意**: 初回起動時にセキュリティダイアログが表示される場合があります。「開く」をクリックしてください。

## トラブルシューティング

### LibreOfficeが見つからない場合
- macOS: `/Applications/LibreOffice.app` を確認
- Windows: `C:\Program Files\LibreOffice\` を確認

### セキュリティ警告（macOS）
「システム環境設定」→「セキュリティとプライバシー」→「一般」で「許可」をクリック

## ファイル構成
```
├── web_app.py              # Webサーバー
├── app.py                  # 変換処理
├── install_macos.sh        # macOS自動セットアップ
├── install_windows.bat     # Windows自動セットアップ
└── static/
    └── index.html          # Webインターフェース
```

## 機能
- LibreOffice直接PNG変換（高速・低メモリ）
- PDF経由変換（フォールバック）
- リアルタイムログ表示
- メモリ最適化（512MB以下）
- キャッシュ機能
