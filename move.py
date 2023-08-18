from copy import deepcopy
from time import perf_counter
from pathlib import Path
from openpyxl import load_workbook
from pebble.concurrent import process
from win32com.client.gencache import EnsureDispatch

from utilities import Watch, colorprint, worksheets_dimensions


@process(timeout=300)
def move_xls(path_source: str, path_destination: str) -> None:
    app = EnsureDispatch("Excel.Application")
    app.DisplayAlerts = False
    file = app.Workbooks.Open(path_source)
    file.SaveAs(path_destination, 51)
    file.Close()
    app.Quit()


@process(timeout=300)
def move_doc(path_source: str, path_destination: str) -> None:
    app = EnsureDispatch("Word.Application")
    app.DisplayAlerts = False
    file = app.Documents.Open(path_source)
    file.SaveAs(path_destination, 16)
    file.Close()
    app.Quit()


@process(timeout=300)
def move_ppt(path_source: str, path_destination: str) -> None:
    app = EnsureDispatch("Powerpoint.Application")
    app.DisplayAlerts = False
    file = app.Presentations.Open(path_source, WithWindow=False)
    file.SaveAs(path_destination, 24)
    file.Close()
    app.Quit()


def move(directory_src: Path = Path("input"), directory_dst: Path = Path("output"), watch: Watch = Watch()) -> None:

    suffix_to_function_suffix = {".xlsx": (move_xls, ".xlsx"), ".xls": (move_xls, ".xlsx"), ".docx": (move_doc, ".docx"), ".doc": (move_doc, ".docx"), ".pptx": (move_ppt, ".pptx"), ".ppt": (move_ppt, ".pptx")}

    path_idx = Path([path.name for path in directory_src.glob("*") if path.suffix == ".xlsx" and path.stat().st_file_attributes != 34][0])
    paths = dict()

    workbook = load_workbook(str(directory_src / path_idx))
    for sheetname, (rows, cols) in worksheets_dimensions(str(directory_src / path_idx)).items():
        for row in range(1, rows + 2):
            for col in range(1, cols + 1):
                cell = workbook[sheetname].cell(row, col)
                if cell.hyperlink is not None:
                    path = Path(cell.hyperlink.target)
                    if path.suffix in (".xlsx", ".xls", ".docx", ".doc", ".pptx", ".ppt") and ".." not in path.parts:
                        if path not in paths:
                            paths[path] = [cell]
                        else:
                            paths[path].append(cell)

    watch.index_ending = len(paths)
    for index, (path, cells) in enumerate(paths.items(), start=1):
        watch.index_current = index
        watch.time_current = perf_counter()
        colorprint("c", path.name, "moving", deepcopy(watch))
        try:
            function, suffix = suffix_to_function_suffix[path.suffix]
            path_old = path
            path_new = path.with_suffix(suffix)
            path_src = directory_src / path_old
            path_dst = directory_dst / path_new
            path_dst.parent.mkdir(parents=True, exist_ok=True)
            function(str(path_src.resolve()), str(path_dst.resolve())).result()
        except Exception as exception:
            for cell in cells: cell.hyperlink.target = str(path_old)
            colorprint("r", path.name, exception, deepcopy(watch))
        else:
            for cell in cells: cell.hyperlink.target = str(path_new)
            colorprint("g", path.name, "moved", deepcopy(watch))

    workbook.save(str(directory_dst / path_idx))
