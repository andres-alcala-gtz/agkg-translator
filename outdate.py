from time import perf_counter
from pathlib import Path
from win32com.client.gencache import EnsureDispatch

from utilities import colorprint


def outdate_xls(paths: set[Path]) -> None:
    app = EnsureDispatch("Excel.Application")
    app.DisplayAlerts = False
    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "outdate_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "outdating")
        try:
            file = app.Workbooks.Open(str(path))
            file.SaveAs(str(path.with_suffix(".xls")), 56)
            file.Close()
        except Exception as exception:
            colorprint("r", "outdate_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, exception)
        else:
            colorprint("g", "outdate_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "outdated")
    app.Quit()


def outdate_doc(paths: set[Path]) -> None:
    app = EnsureDispatch("Word.Application")
    app.DisplayAlerts = False
    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "outdate_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "outdating")
        try:
            file = app.Documents.Open(str(path))
            file.SaveAs(str(path.with_suffix(".doc")), 0)
            file.Close()
        except Exception as exception:
            colorprint("r", "outdate_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, exception)
        else:
            colorprint("g", "outdate_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "outdated")
    app.Quit()


def outdate_ppt(paths: set[Path]) -> None:
    app = EnsureDispatch("Powerpoint.Application")
    app.DisplayAlerts = False
    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "outdate_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "outdating")
        try:
            file = app.Presentations.Open(str(path), WithWindow=False)
            file.SaveAs(str(path.with_suffix(".ppt")), 1)
            file.Close()
        except Exception as exception:
            colorprint("r", "outdate_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, exception)
        else:
            colorprint("g", "outdate_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "outdated")
    app.Quit()


def outdate(directory: Path = Path("output")) -> None:

    paths_xls = {path.absolute() for path in directory.rglob("*") if path.suffix == ".xlsx" and path.stat().st_file_attributes != 34}
    paths_doc = {path.absolute() for path in directory.rglob("*") if path.suffix == ".docx" and path.stat().st_file_attributes != 34}
    paths_ppt = {path.absolute() for path in directory.rglob("*") if path.suffix == ".pptx" and path.stat().st_file_attributes != 34}

    outdate_xls(paths_xls)
    outdate_doc(paths_doc)
    outdate_ppt(paths_ppt)
