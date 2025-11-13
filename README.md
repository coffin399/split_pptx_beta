# PPTX Script Slide Generator

PowerPoint の「ノート欄」に書かれた文字を、読み上げ用の大きな黒背景スライドに作り直すツールです。Web API とデスクトップアプリの両方を提供しており、ボタンを押すだけで自動変換できるので、パソコン操作に不慣れな方でも安心して使えます。

---

## 📋 プロジェクトサマリー

### 最新状態 (2025-11-13)
- **✅ Render.com デプロイ対応完了**
- **🔄 Flask バックエンドに統合**
- **🌐 Web API + Vue.js フロントエンド構成**
- **🚀 クラウド即時利用可能**

### アーキテクチャ変更履歴
1. **初期**: デスクトップアプリ (PySide6) + ローカルFlask
2. **中間**: FastAPI Web API + デプロイ対応
3. **現在**: Flask Web API + Render.com最適化

---

## 🌐 Web版（Render.com）

### 特徴
- **クラウドベース**: インストール不要、ブラウザから即利用
- **非同期処理**: バックグラウンドでPPTX変換を実行
- **メモリ管理**: 自動クリーンアップ + 400MB制限
- **キュー処理**: 複数リクエストを順次処理

### 利用方法
1. Render.com にデプロイされたURLにアクセス
2. PPTXファイルをドラッグ＆ドロップ
3. 変換完了後にダウンロード

### 技術スタック
- **バックエンド**: Flask + python-pptx + psutil
- **フロントエンド**: Vue.js 3 + axios
- **デプロイ**: Docker + Render.com
- **システム**: Python 3.10 + 日本語フォント

---

## 💻 デスクトップ版（従来）

### できること
- ノート欄の文章を自動で読みやすい長さ（約150文字）ごとに分割
- 話者ごとに文字色を自動変更
  - 仲條: 水色 (#00FDFF)
  - 三村: 白色 (#FFFFFF)  
  - 星野: 黄色 (#FFFF00)
  - その他: 自動割り当て（ピンク、オレンジなど）
- 黒い背景・大きな文字（メイリオ 40pt）で表示
- ノートが長いときは 1/2, 2/2 のようにページ番号を右下に表示
- 右下に元スライドのサムネイルを自動で挿入
- 出力ファイル名は「スクリプトスライド_自動生成.pptx」で固定

---

## 🚀 デプロイ構成

### ファイル構成
```
split_pptx_beta/
├── .dockerignore          # Dockerビルド除外設定
├── .git/                  # Gitリポジトリ  
├── Dockerfile             # Render.comデプロイ用
├── README.md              # 本ファイル
├── app.py                 # Flaskバックエンド（統合版）
├── render.yaml            # Render.comサービス設定
├── requirements.txt       # Python依存関係
└── static/
    └── index.html         # Vue.jsフロントエンド
```

### APIエンドポイント
| エンドポイント | メソッド | 機能 |
|---|---|---|
| `/` | GET/HEAD | フロントエンド配信 |
| `/health` | GET | ヘルスチェック + メモリ状況 |
| `/convert` | POST | PPTXアップロード・変換 |
| `/status/{task_id}` | GET | 変換ステータス取得 |
| `/download/{task_id}` | GET | 変換結果ダウンロード |
| `/cleanup/{task_id}` | DELETE | タスククリーンアップ |

---

## 🛠️ 開発者向け

### ローカル開発環境
```bash
# Python 3.10 を使用
python --version

# 仮想環境作成
python -m venv .venv
.venv\Scripts\activate         # Windows
source .venv/bin/activate       # macOS

# 依存関係インストール
pip install -r requirements.txt

# Flaskアプリ起動
python app.py
```

### 依存関係
```txt
flask>=2.3.0
werkzeug>=2.3.0
python-pptx>=0.6.21
psutil>=5.9.0
PySide6>=6.6.0,<6.8  # GUI版用
Pillow>=10.0.0
pdf2image>=1.17.0
```

### Render.com デプロイ
```bash
git add .
git commit -m "Update Flask backend for Render.com"
git push origin main
# 自動デプロイ開始
```

---

## 📁 出力されるスライド仕様

### レイアウト設定
- **スライドサイズ**: 33.867 × 19.05 cm (16:9)
- **テキスト領域**: 左 0.79cm, 上 0.80cm, 幅 25.2cm, 高さ 15.6cm
- **フォント**: メイリオ 40pt, 太字
- **背景**: 黒 (#000000)
- **ページ番号**: 右下 21.94cm, 16.93cm (青色 #009DFF)
- **サムネイル枠**: 右下 25.87cm, 14.55cm, 8.0×4.5cm

### テキスト処理
- **最大文字数**: 150文字/スライド
- **話者検出**: `《話者名》` パターンで自動判定
- **文字分割**: 句読点優先、150文字単位で強制分割
- **連結処理**: 同じ話者の連続セグメントを自動結合

---

## 🔄 更新履歴

- **2025-11-13**: Flaskバックエンドに統合、Render.comデプロイ対応
- **2025-11-13**: FastAPIからFlaskへ移行、依存関係整理
- **2025-11-13**: ローカルアプリファイル削除、Docker構成最適化
- **以前**: デスクトップアプリ + PySide6 GUI
