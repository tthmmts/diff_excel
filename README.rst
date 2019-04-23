Diff Excel
========================

2つのエクセルデータを比較して差分を表示する．

機能
------------------------

- 2つのエクセルファイルを比較して差分を表示する．
- （オプション）差分をエクセルシートに書き出す．

限界
------------------------

- 遅い！
- データベース的な Excel ファイルにのみ対応
    - ファイルの先頭1行をフィールド（カラム|列）名として強制的に利用
    - しかも，最初のシートしか比べない
- シート内の値はテキストとして処理
    - 数値とテキストの相互変換は検知できない
- サポートなし


要件
------------------------

- Python 3.5+
- pipenv or pip
- modeules
    - numpy, openpyxl, xlrd, pytz, pandas, docopt

導入
------------------------
::

  cd <your path>
  git clone https://github.com/tthmmts/diff_excel.git
  cd  diff_excel
  pip install -e .

or::

  pipenv install

使い方
------------------------

usage::

    diff_excel.py [-h]
    diff_excel.py [-e | --excel] [--key=<key>] FROMFILE TOFILE
    diff_excel.py [-e | --excel] [-k <key>] FROMFILE TOFILE

options:
    -h, --help    show this help message and exit
    -e, --excel   結果をExcelファイルで出力する（指定しなければ，テキスト出力のみ）
    --key=<key>   IDとなる列カラムの名前．指定しなければ行番号となる．


データを比較して，コマンドラインに表示． ::

  diff_excel --key="id" from_book.xlsx to_book.xlsx

keyを省略すると行番号を利用する（つまりシートの2行目がkey "1" となる）． ::

  diff_excel from_book.xlsx to_book.xlsx

結果を Excel ファイルに書き出す．
ファイルはカレントディレクトリに diff\_results\_％Y%m%dT%H%M%S.xlsx" の名前で出力される． ::

  diff_excel --excel --key="id" from_book.xlsx to_book.xlsx


実行例
------------------------
./tests/test_data/ ディレクトリにテスト用のファイルがある．


いくつかのセルで数値を変更 ::

  diff_excel ./tests/test_data/from_book.xlsx ./tests/test_data/to_book.xlsx

>>>
行と列の削除や追加はありませんでした．
======================================
1行:
  b列: 2 -> 3
2行:
  b列: 5 -> 3
3行:
  a列: 7 -> 12
  c列: 9 -> 2
======================================


行と列の削除と挿入 ::

  diff_excel --key="id" --excel ./tests/test_data/from_book.xlsx ./tests/test_data/insert_delete_book.xlsx


>>>
1行が削除されました．
  行番号: 2
1列が削除されました．
  列番号: b
1行が追加されました．
  行番号: 2+1
1列が追加されました．
  列番号: a+1
======================================
./diff_results_00000000T000000.xlsx に結果を保存しました．

MacOS ::

  open ./diff_results_00000000T000000.xlsx

Linux ::

  xdg-open ./diff_results_00000000T000000.xlsx

Windows ::

  start ./diff_results_00000000T000000.xlsx
