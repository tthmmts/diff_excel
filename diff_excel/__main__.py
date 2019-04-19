import sys
import platform
from docopt import docopt
from core import diff_excel_data, show_diff_excel_data_as_text, show_diff_excel_data_as_xlsx
from pathlib import Path

__doc__ = """エクセルの新旧データ比較

usage: 
    diff_excel.py [-h]
    diff_excel.py [-e | --excel] [--key=<key>] FROMFILE TOFILE
    diff_excel.py [-e | --excel] [-k <key>] FROMFILE TOFILE

options:
    -h, --help    show this help message and exit
    -e, --excel   結果をExcelファイルで出力する（指定しなければ，テキスト出力のみ）
    --key=<key>   IDとなる列カラムの名前．指定しなければ行番号となる．
"""

if __name__ == '__main__':
    args = docopt(__doc__)
    # print(args)
    if args["FROMFILE"] and args["TOFILE"]:
        diff_excel_data_results = diff_excel_data(Path(args.get("FROMFILE","")), Path(args.get("TOFILE","")), key=args.get("--key"))
        text = show_diff_excel_data_as_text(diff_excel_data_results)
        print(text)
    else:
        print("ファイルを指定してください．")    
        print(__doc__)   

    if args["--excel"]:
        show_diff_excel_data_as_xlsx(diff_excel_data_results)
        pf = platform.system()
        if pf == 'Windows':
            pass
        elif pf == 'Darwin':
            pass
        elif pf == 'Linux':
            pass

    sys.exit(1)
