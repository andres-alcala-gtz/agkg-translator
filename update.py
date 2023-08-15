from time import perf_counter
from pathlib import Path
from win32com.client.gencache import EnsureDispatch

from utilities import colorprint


def update_xls(paths: set[Path]) -> None:
    app = EnsureDispatch("Excel.Application")
    app.DisplayAlerts = False
    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "update_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "updating")
        try:
            file = app.Workbooks.Open(str(path))
            file.SaveAs(str(path.with_suffix(".xlsx")), 51)
            file.Close()
        except Exception as exception:
            colorprint("r", "update_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, exception)
            path.unlink()
        else:
            colorprint("g", "update_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "updated")
            if path.suffix == ".xls":
                path.unlink()
    app.Quit()


def update_doc(paths: set[Path]) -> None:
    app = EnsureDispatch("Word.Application")
    app.DisplayAlerts = False
    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "update_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "updating")
        try:
            file = app.Documents.Open(str(path))
            file.SaveAs(str(path.with_suffix(".docx")), 16)
            file.Close()
        except Exception as exception:
            colorprint("r", "update_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, exception)
            path.unlink()
        else:
            colorprint("g", "update_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "updated")
            if path.suffix == ".doc":
                path.unlink()
    app.Quit()


def update_ppt(paths: set[Path]) -> None:
    app = EnsureDispatch("Powerpoint.Application")
    app.DisplayAlerts = False
    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "update_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "updating")
        try:
            file = app.Presentations.Open(str(path), WithWindow=False)
            file.SaveAs(str(path.with_suffix(".pptx")), 24)
            file.Close()
        except Exception as exception:
            colorprint("r", "update_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, exception)
            path.unlink()
        else:
            colorprint("g", "update_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "updated")
            if path.suffix == ".ppt":
                path.unlink()
    app.Quit()


def update_mso(directory: Path = Path("output")) -> None:

    paths_xls = {path.absolute() for path in directory.rglob("*") if path.suffix in (".xlsx", ".xls") and path.stat().st_file_attributes != 34}
    paths_doc = {path.absolute() for path in directory.rglob("*") if path.suffix in (".docx", ".doc") and path.stat().st_file_attributes != 34}
    paths_ppt = {path.absolute() for path in directory.rglob("*") if path.suffix in (".pptx", ".ppt") and path.stat().st_file_attributes != 34}

    update_xls(paths_xls)
    update_doc(paths_doc)
    update_ppt(paths_ppt)
