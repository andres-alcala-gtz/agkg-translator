from time import perf_counter
from numpy import array_split
from pandas import ExcelFile
from datetime import datetime, timedelta
from dataclasses import dataclass


@dataclass
class Watch:
    index_current: int = None
    index_ending: int = None
    time_current: float = None
    time_beginning: float = None


def list_split(values: list, sections: int = 8) -> list:
    return [list(array) for array in array_split(values, sections)]


def worksheets_dimensions(path: str) -> dict:
    file = ExcelFile(path)
    info = dict()
    for sheet_name in file.sheet_names:
        data = file.parse(sheet_name)
        info.update({sheet_name: data.shape})
    return info


def colorprint(color: str, path: str, message: str, watch: Watch) -> None:
    colors = {"r": "\x1b[31m", "g": "\x1b[32m", "y": "\x1b[33m", "b": "\x1b[34m", "m": "\x1b[35m", "c": "\x1b[36m", "w": "\x1b[37m", "n": "\x1b[39m"}
    print(f"{colors.get('m')}{datetime.now():%X} - [{watch.index_current}/{watch.index_ending}] - [{timedelta(seconds=int(perf_counter() - watch.time_current))}/{timedelta(seconds=int(perf_counter() - watch.time_beginning))}] -> {colors.get('b')}{path} -> {colors.get(color)}{message}{colors.get('n')}")
