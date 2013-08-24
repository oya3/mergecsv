機能：
指定したフォルダ配下にあるcsvファイルを１つに纏めxlsにする。
また、同階層にあるcsvや１つに纏めたxlsとの突合せができる。
オプションで-noguiを設定しないとgui版が起動する。

使い方：
Usage: mergecsv [options] <pattern file name> <search directory>
  options : -expected_path     ... expected directory. need same hierarchy of search directory.
          : -expected_xls_path ... expected xls file. need sheet name is "expected".
          : -o                 ... output file.
          : -nogui             ... no gui mode.
          : -dbg               ... debug mode.

仕様：
検索対象のファイルは入力パラメータで指定したファイルのみ（正規表現未対応）
ファイルのフォーマットは発見した一つ目をベースとし同じ形式のファイルのみを集める
CSVのフォーマットは、key,value,... の連続のみ
xls ファイル出力時は、ms excel は必要ない。
perl がインストール済みであること。（exe 版利用の場合は不要)
ppm で spreadsheet 関連モジュールがインストール済みであること。

動作確認：
windows 7(32,64)環境のみ

同梱内容：
mergecsv.pl
mergecsv.exe

