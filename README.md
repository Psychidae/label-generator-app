PythonによるExcelデータシート自動生成ツール（最終版）
このツールは、最小限の情報（緯度経度、日付など）が書かれたシンプルなCSVファイルを読み込み、Google Maps APIを利用して詳細な住所と高度を取得します。

最終的に、あなたが普段お使いのExcelファイルの「入力用シート」と全く同じレイアウト・内容の新しいExcelファイルを自動で生成します。

1. 必要なもの
Python 3: 公式サイトからインストールしてください。

Google Maps APIキー: 以下の両方のAPIを有効にしたキーが必要です。

Geocoding API （住所取得用）

Elevation API （高度取得用）

2. セットアップ
スクリプトの実行に必要なPythonライブラリをインストールします。

macのターミナルを開きます。

以下のコマンドを1行ずつ実行します。

pip3 install pandas requests tqdm openpyxl pykakasi

3. 実行前の準備
generate_data_sheet.py と、input_data.csv（サンプル）を同じフォルダに保存します。

input_data.csv を編集し、ご自身のデータに書き換えてください。列名はサンプルに合わせておくのが最も簡単です。

4. スクリプトの実行方法
ターミナルで、ファイルを保存したフォルダに移動します (cd コマンド)。

以下のコマンドの**「AIzaSyCBQgSdWWeDvu2wXX98wIResMBDSBC01pU」**の部分を、ご自身のキーに置き換え、入力ファイル名と出力ファイル名を指定して実行します。

python3 generate_data_sheet.py "AIzaSyCBQgSdWWeDvu2wXX98wIResMBDSBC01pU" input_data.csv labels_data_output.xlsx

5. 出力結果
labels_data_output.xlsx というExcelファイルが自動で生成されます。

このファイルには**「入力用シート」**という名前のシートが作成されており、その中身は、あなたが普段使っているExcelファイルの形式と完全に一致しています。

この生成されたExcelシートの必要な部分を、あなたのラベル用ファイルにコピー＆ペーストするだけで、すべての作業が完了します。
