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