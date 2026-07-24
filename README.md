# Licosha写真展用

このリポジトリはLicosha写真展用のリポジトリです.

## フォーム

1. フォームの作成 \
   まず,[こちら](https://forms.app/myforms)からフォームを複製してください.
   フォームの名前は,`2022 早稲田祭写真展`のようにしてください.
2. フォームの編集　\
    フォームの編集は,丸が三つ並んだボタンから,`Edit`を押すとできます.\
    以下の項目を編集してください.\
    **なお,変更があったたびに`Save`を押すか,`Ctrl + S`を押してください.**
    - form>edit\
      タイトルを変更してください.
    - form>settings>integrations>google sheets \
      ここで,`Connect to Google Sheets`を押してください.
      すると,`Select a spreadsheet`のところに,`Create a new spreadsheet`と出てくるので,それを選択してください.\
      *スプレッドシートは,はじめの回答が来たタイミングで作成されます.*
    - form>settings>share　\
        `Privacy Settings`のところで`Unlisted`にしてください.\
        最後に`copy link`を押して,リンクをコピーしてください.

なお,**プログラミングのコード内では,列番号で各情報を指定しているため,フォーム自体の編集（組み替え）をおこなってしまうと狂います.**

## アプリケーション

1. スプレッドシートからエクセルファイルを作成 \
   まず,[こちら](https://docs.google.com/spreadsheets/u/0/)からスプレッドシートを開きます.\
   ファイル>ダウンロード>Microsoft Excelを選択してください.
2. 保存するファイルの作成 \
    生成されるファイル類を保存するファイルを作成します.
    保存するファイルは,`2022 早稲田祭写真展`のようにしてください.
3. プログラムの実行 \
アプリケーション内で,エクセルファイルと保存するフォルダーを選択して,`Generate`を押してください.

## アプリのインストール（コードを書けない人向け）

Windows / Mac 用の実行ファイルは，GitHub Actions が自動でビルドして
[Releases](../../releases) ページに公開します．開発者がMacで手動でWindows用に
ビルドする必要はありません．

- **使う人（インストールする人・Windows）**\
  Windowsキーを押して`PowerShell`と入力し,PowerShellを開いてください．\
  以下の1行をコピーして貼り付け,Enterキーを押すだけでインストールできます．
  ```powershell
  irm https://raw.githubusercontent.com/mksmkss/Display/main/install.ps1 | iex
  ```
  自動的に最新版がダウンロードされ,デスクトップに「Display」というアイコンが作成されます．\
  次回以降アップデートしたいときも,同じ1行をもう一度実行すれば最新版に上書きされます．

  ※ zipを手動でダウンロードしたい場合は,[Releases](../../releases) ページから
  `Display-Windows.zip`（Windowsの場合）または `Display-macOS.zip`（Macの場合）を
  ダウンロードして展開し,中の `Display.exe`（Mac版は`Display.app`）を実行しても構いません．\
  いずれの方法でも,インストール作業やPythonのセットアップは不要です．

- **開発する人（新しいバージョンを配布する人）**\
  1. `gui_2.py` 内の `version_label` の表示を更新するなど，通常通り変更を加えて `main` にマージしてください．
  2. バージョンタグを作成してpushします．
     ```sh
     git tag v3.1.0
     git push origin v3.1.0
     ```
  3. タグをpushすると `.github/workflows/build-release.yml` が自動的に実行され,
     Windows / macOS 両方の実行ファイルがビルドされ,Releaseとして公開されます．
     数分待てば[Releases](../../releases)ページに新しいzipが並びます．

  タグをpushせずに動作確認だけしたい場合は,GitHubの「Actions」タブから
  `Build and Release` を選択して「Run workflow」を押すと,Releaseを作らずに
  ビルド結果だけをArtifactsとしてダウンロードできます．