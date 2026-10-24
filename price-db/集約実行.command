#!/bin/bash
# Macでダブルクリックすると価格マスタを作成します
cd "$(dirname "$0")" || exit 1

if [ ! -f config.ini ]; then
  cp config.example.ini config.ini
  echo "config.ini を作成しました。input_dir を共有フォルダのパスに書き換えてから、もう一度実行してください。"
  open -e config.ini
  read -r -p "Enterキーで閉じます"
  exit 1
fi

if [ ! -d .venv ]; then
  echo "初回セットアップ中..."
  python3 -m venv .venv || { echo "python3 が見つかりません。README の手順を確認してください。"; read -r -p "Enterキーで閉じます"; exit 1; }
  .venv/bin/pip install -q -r requirements.txt
fi

.venv/bin/python collect.py && open output/ 2>/dev/null
read -r -p "完了しました。Enterキーで閉じます"
