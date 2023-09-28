import shutil
import contextlib
import win32com.client


def clear_temp() -> None:
    shutil.rmtree(win32com.__gen_path__, ignore_errors=True)


@contextlib.contextmanager
def open_xls(path: str):
    app = win32com.client.gencache.EnsureDispatch("Excel.Application")
    app.DisplayAlerts = False
    file = app.Workbooks.Open(path)
    try:
        yield file
    finally:
        file.Close()
        app.Quit()


@contextlib.contextmanager
def open_doc(path: str):
    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.DisplayAlerts = False
    file = app.Documents.Open(path)
    try:
        yield file
    finally:
        file.Close()
        app.Quit()


@contextlib.contextmanager
def open_ppt(path: str):
    app = win32com.client.gencache.EnsureDispatch("Powerpoint.Application")
    app.DisplayAlerts = False
    file = app.Presentations.Open(path, WithWindow=False)
    try:
        yield file
    finally:
        file.Close()
        app.Quit()
