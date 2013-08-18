機能：
指定したフォルダ配下にあるcsvファイルを１つに纏めxlsにする。

Usage: mergecsv [options] <pattern file name> <search directory>
  options : -expected ... expected directory. need same hierarchy of search directory
          : -xls      ... expected xls file. need sheet name is "expected"
          : -o        ... output file
          : -dbg      ... debug mode

仕様：
検索対象のファイルは入力パラメータで指定したファイルのみ（正規表現未対応）
ファイルのフォーマットは発見した一つ目をベースとし同じ形式のファイルのみを集める
CSVのフォーマットは、key,value,... の連続のみ
出力ファイル名は、"検索ファイル名"+.xlsとなる（.は_に変換される）
xls ファイル出力時は、ms excel は必要ない。
perl がインストール済みであること。
ppm で spreadsheet 関連モジュールがインストール済みであること。

動作確認：
windows 7(32,64)環境のみ

同梱内容：
mergecsv.pl
ツール本体

run.bat
同梱されているdummy配下にある"リスト結果.csv"ファイルを検索し"リスト結果_csv.xls"を作成するサンプル

run_drop.bat
test.bat 同様の結果となるが、検索対象フォルダを本バッチにドロップすることで起動するサンプル

