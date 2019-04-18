import difflib
import pandas as pd
from pathlib import Path


def get_dict_data_from_excel_file(file_path: Path, key: str = "") -> tuple:
    """

    Args:
        file_path: エクセルファイルのフルパス
        key: IDとなる列の文字

    Returns:

    """

    if not file_path.is_file():
        raise Exception("{0} というファイルは見つからない".format(file_path))

    df = pd.read_excel(file_path, dtype=str)
    df.columns = [str(i) for i in df.columns]

    if key == "":
        df.index = [str(i + 1) for i in df.index]
    elif key in df.columns:
        if len(set(df[key].values)) == len(df.index):
            df.index = df[key].values
            del df[key]
    else:
        raise Exception("key:\"{0}\"が\"{1}\"で見つかりません．".format(key, file_path))
    d_index = df.to_dict(orient='index')

    return d_index, list(d_index.keys()), list(df.columns)


def list_difference(first: list, second: list) -> list:
    """
    Source: https://stackoverflow.com/questions/6486450/python-compute-list-difference
    :param first:
    :param second:
    :return: list
    """
    second = set(second)

    return [item for item in first if item not in second]


def list_intersection(first: list, second: list) -> list:
    """
    積集合

    :param first:
    :param second:
    :return: list
    """
    second = set(second)

    return [item for item in first if item in second]


def list_ndiff(first: list, second: list) -> list:
    """

    Args:
        first:
        second:

    Returns:

    """
    results = []
    # print(difflib.ndiff(first, second))
    for l in difflib.ndiff(first, second):
        results.append(l)
    return results
