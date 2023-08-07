from time import perf_counter
from shutil import copyfile
from pathlib import Path
from openpyxl import load_workbook

from utilities import colorprint, worksheets_dimensions


def move_mso(directory_src: Path = Path("input"), directory_dst: Path = Path("output")) -> None:

    path_idx = Path([path.name for path in directory_src.glob("*") if path.suffix == ".xlsx" and path.stat().st_file_attributes != 34][0])
    paths = {path_idx}

    workbook = load_workbook(str(directory_src / path_idx))
    for sheetname, (rows, cols) in worksheets_dimensions(str(directory_src / path_idx)).items():
        for row in range(1, rows + 2):
            for col in range(1, cols + 1):
                cell = workbook[sheetname].cell(row, col)
                if cell.hyperlink is not None:
                    path = Path(cell.hyperlink.target)
                    if path.suffix in (".xlsx", ".xls", ".docx", ".doc", ".pptx", ".ppt") and ".." not in path.parts:
                        paths.add(path)

    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "move_mso", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "moving")
        try:
            path_src = directory_src / path
            path_dst = directory_dst / path
            path_dst.parent.mkdir(parents=True, exist_ok=True)
            copyfile(src=str(path_src), dst=str(path_dst))
        except Exception as exception:
            colorprint("r", "move_mso", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, exception)
        else:
            colorprint("g", "move_mso", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "moved")
