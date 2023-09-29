import time
import numpy
import pandas
import pathlib
import datetime


class Watch:

    def __init__(self) -> None:
        self.time_current = time.perf_counter()
        self.time_beginning = time.perf_counter()
        self.index_current = 0
        self.index_ending = 0
        self.path_current = pathlib.Path()

    def current(self, index: int, path: pathlib.Path) -> None:
        self.time_current = time.perf_counter()
        self.index_current = index
        self.path_current = path

    def beginning(self, index: int) -> None:
        self.index_ending = index

    def print(self, message: str) -> None:
        print(f"{datetime.datetime.now():%X} - [{self.index_current}/{self.index_ending}] - [{datetime.timedelta(seconds=int(time.perf_counter() - self.time_current))}/{datetime.timedelta(seconds=int(time.perf_counter() - self.time_beginning))}] -> {self.path_current.name} -> {message}")


def list_split(values: list, sections: int) -> list[list]:
    return [list(array) for array in numpy.array_split(values, sections)]


def worksheets_dimensions(path: str) -> dict[int | str, tuple[int, int]]:
    file = pandas.ExcelFile(path)
    info = dict()
    for sheet_name in file.sheet_names:
        data = file.parse(sheet_name)
        info.update({sheet_name: data.shape})
    return info
