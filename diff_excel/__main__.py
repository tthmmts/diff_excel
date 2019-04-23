import sys
import platform
from docopt import docopt
from . import core
from pathlib import Path

__version__ = "0.2.5"

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


def main():
    args = docopt(__doc__)
    print(args)
    if args.get("-h") or args.get("--help"):
        print(__doc__)
    elif args.get("FROMFILE", None) and args.get("TOFILE", None):
        diff_excel_data_results = core.diff_excel_data(Path(args.get("FROMFILE", "")),
                                                       Path(args.get("TOFILE", "")),
                                                       key=args.get("--key"))
        text = core.show_diff_excel_data_as_text(diff_excel_data_results)
        print(text)

        if args["--excel"]:
            core.show_diff_excel_data_as_xlsx(diff_excel_data_results)
            pf = platform.system()
            # if pf == 'Windows':
            #     pass
            # elif pf == 'Darwin':
            #     pass
            # elif pf == 'Linux':
            #     pass
    else:
        print("ファイルを指定してください．")
        print(__doc__)
        sys.exit(1)


if __name__ == '__main__':
    main()
    sys.exit(0)
