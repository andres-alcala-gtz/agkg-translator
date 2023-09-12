import time
import numpy
import pandas
import datetime
import dataclasses


@dataclasses.dataclass
class Watch:
    index_current: int = None
    index_ending: int = None
    time_current: float = None
    time_beginning: float = None


def list_split(values: list, sections: int = 8) -> list:
    return [list(array) for array in numpy.array_split(values, sections)]


def worksheets_dimensions(path: str) -> dict:
    file = pandas.ExcelFile(path)
    info = dict()
    for sheet_name in file.sheet_names:
        data = file.parse(sheet_name)
        info.update({sheet_name: data.shape})
    return info


def colorprint(color: str, path: str, message: str, watch: Watch) -> None:
    colors = {"r": "\x1b[31m", "g": "\x1b[32m", "y": "\x1b[33m", "b": "\x1b[34m", "m": "\x1b[35m", "c": "\x1b[36m", "w": "\x1b[37m", "n": "\x1b[39m"}
    print(f"{colors.get('m')}{datetime.datetime.now():%X} - [{watch.index_current}/{watch.index_ending}] - [{datetime.timedelta(seconds=int(time.perf_counter() - watch.time_current))}/{datetime.timedelta(seconds=int(time.perf_counter() - watch.time_beginning))}] -> {colors.get('b')}{path} -> {colors.get(color)}{message}{colors.get('n')}")
