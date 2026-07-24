#!/bin/bash
set -e

REPO="mksmkss/Display"
INSTALL_DIR="$HOME/Applications"
ZIP_PATH="$(mktemp -t Display-macOS).zip"

echo "最新版の情報を取得しています..."
DOWNLOAD_URL=$(curl -fsSL "https://api.github.com/repos/$REPO/releases/latest" \
  | grep browser_download_url \
  | grep Display-macOS.zip \
  | cut -d '"' -f 4)

if [ -z "$DOWNLOAD_URL" ]; then
  echo "Display-macOS.zip が見つかりませんでした。まだリリースが公開されていない可能性があります。" >&2
  exit 1
fi

echo "ダウンロード中..."
curl -fsSL "$DOWNLOAD_URL" -o "$ZIP_PATH"

echo "インストール中..."
mkdir -p "$INSTALL_DIR"
rm -rf "$INSTALL_DIR/Display"
unzip -oq "$ZIP_PATH" -d "$INSTALL_DIR"
rm -f "$ZIP_PATH"

APP_PATH="$INSTALL_DIR/Display/Display.app"

# 署名されていないアプリのため,Gatekeeperの「開発元が未確認」ブロックを解除する
xattr -cr "$APP_PATH"

echo ""
echo "インストール完了！ $APP_PATH をダブルクリックすれば起動できます。"
open "$INSTALL_DIR/Display"
