# PPTX Script Slide Generator

PowerPoint のノート欄からシナリオ用スライドを自動生成する Python GUI アプリです。

## 機能概要

- ノート欄のテキストを句読点優先で分割して 1 枚あたり 200 文字以内で表示
- 話者ごとに指定色でテキストを着色（話者1: #FFFF00, 話者2: #00FFFF, 話者3: #00F900）
- 黒背景・メイリオ 40pt の固定フォーマット
- 右下に元スライドのサムネイル画像（LibreOffice が無い場合はプレースホルダー）
- 分割されたスライドに 1/2, 2/2 のようなページ表示（右下, 40pt, #00B0F0）
- 出力ファイル名は固定で `スクリプトスライド_自動生成.pptx`

## 必要条件

- Python 3.9 – 3.11
- 必要ライブラリ：python-pptx, PySimpleGUI, Pillow
- （任意）LibreOffice / soffice がインストールされているとサムネイル生成が可能

## セットアップ

```bash
python -m venv .venv
.venv\Scripts\activate         # Windows
source .venv/bin/activate       # macOS / Linux
pip install -r requirements.txt
```

## 使い方

```bash
python app.py
```

1. GUI 上で入力 PPTX を選択
2. 出力フォルダを指定（未指定なら入力ファイルと同じ場所）
3. 「変換」を押すと処理が始まり、ログ欄に状況が表示されます
4. 完了すると出力フォルダに `スクリプトスライド_自動生成.pptx` が作成されます

## サムネイルについて

- LibreOffice (soffice) が利用できる場合は自動でスライドごとの PNG を生成
- 利用できない場合はプレースホルダー画像で代替

## 注意事項

- ノート欄が空のスライドはスキップされます
- 話者ラベルは `話者1：テキスト` のように全角コロンまたは半角コロンで判定
- 句読点が存在しない長文は 200 文字ごとに自動分割されます

## CI/CD とリリース

- `master` ブランチに push すると GitHub Actions が自動で macOS / Windows / Linux 向けに Nuitka ビルドを実行します
- ビルドが全 OS で成功すると、既存の `latest` タグおよびリリースを削除し、新しい成果物で再作成します
- ワークフローは `.github/workflows/build-release.yml` に定義されています（手動トリガーも可能）

## ライセンス

MIT License
