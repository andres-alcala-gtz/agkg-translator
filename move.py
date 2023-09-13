import time
import copy
import pathlib
import openpyxl
import win32com.client
import pebble.concurrent

import utilities


@pebble.concurrent.process(timeout=300)
def move_xls(path_source: str, path_destination: str) -> None:
    app = win32com.client.gencache.EnsureDispatch("Excel.Application")
    app.DisplayAlerts = False
    file = app.Workbooks.Open(path_source)
    file.SaveAs(path_destination, 51)
    file.Close()
    app.Quit()


@pebble.concurrent.process(timeout=300)
def move_doc(path_source: str, path_destination: str) -> None:
    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.DisplayAlerts = False
    file = app.Documents.Open(path_source)
    file.SaveAs(path_destination, 16)
    file.Close()
    app.Quit()


@pebble.concurrent.process(timeout=300)
def move_ppt(path_source: str, path_destination: str) -> None:
    app = win32com.client.gencache.EnsureDispatch("Powerpoint.Application")
    app.DisplayAlerts = False
    file = app.Presentations.Open(path_source, WithWindow=False)
    file.SaveAs(path_destination, 24)
    file.Close()
    app.Quit()


def move(directory_src: pathlib.Path, directory_dst: pathlib.Path, watch: utilities.Watch) -> None:

    suffix_to_function_suffix = {".xlsx": (move_xls, ".xlsx"), ".xls": (move_xls, ".xlsx"), ".docx": (move_doc, ".docx"), ".doc": (move_doc, ".docx"), ".pptx": (move_ppt, ".pptx"), ".ppt": (move_ppt, ".pptx")}

    path_idx = pathlib.Path([path.name for path in directory_src.glob("*") if path.suffix == ".xlsx" and path.stat().st_file_attributes != 34][0])
    paths = dict()

    workbook = openpyxl.load_workbook(str(directory_src / path_idx))
    for sheetname, (rows, cols) in utilities.worksheets_dimensions(str(directory_src / path_idx)).items():
        for row in range(1, rows + 2):
            for col in range(1, cols + 1):
                cell = workbook[sheetname].cell(row, col)
                if cell.hyperlink is not None:
                    path = pathlib.Path(cell.hyperlink.target)
                    if path.suffix in (".xlsx", ".xls", ".docx", ".doc", ".pptx", ".ppt") and ".." not in path.parts:
                        if path not in paths:
                            paths[path] = [cell]
                        else:
                            paths[path].append(cell)

    watch.index_ending = len(paths)
    for index, (path, cells) in enumerate(paths.items(), start=1):
        watch.index_current = index
        watch.time_current = time.perf_counter()
        utilities.colorprint("c", path.name, "moving", copy.deepcopy(watch))
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
            utilities.colorprint("r", path.name, exception, copy.deepcopy(watch))
        else:
            for cell in cells: cell.hyperlink.target = str(path_new)
            utilities.colorprint("g", path.name, "moved", copy.deepcopy(watch))

    workbook.save(str(directory_dst / path_idx))
