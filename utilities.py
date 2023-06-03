from pandas import ExcelFile
from datetime import timedelta


def worksheets_dimensions(path: str) -> dict:
    file = ExcelFile(path)
    info = dict()
    for sheet_name in file.sheet_names:
        data = file.parse(sheet_name)
        info.update({sheet_name: data.shape})
    return info


def colorprint(color: str, function_name: str, index_current: int, index_ending: int, time_current: int, time_beginning: int, path_name: str, message: str) -> None:
    colors = {"r": "\x1b[31m", "g": "\x1b[32m", "y": "\x1b[33m", "b": "\x1b[34m", "m": "\x1b[35m", "c": "\x1b[36m", "w": "\x1b[37m", "n": "\x1b[39m"}
    print(f"{colors.get('m')}{function_name} - [{index_current}/{index_ending}] - [{timedelta(seconds=time_current)}/{timedelta(seconds=time_beginning)}] -> {colors.get('b')}{path_name} -> {colors.get(color)}{message}{colors.get('n')}")
