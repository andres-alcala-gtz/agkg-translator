from re import search
from time import perf_counter
from pathlib import Path
from functools import cache
from multiprocessing import Process
from deep_translator import GoogleTranslator

from docx import Document
from pptx import Presentation
from openpyxl import load_workbook as Workbook

from utilities import colorprint, worksheets_dimensions


@cache
def translate_str(text: str) -> str:
    return GoogleTranslator(source="auto", target="en").translate(text)


def translate_xls(paths: set[Path]) -> None:

    def _translate(container):
        if type(container.value) == str and search("[\u4E00-\u9FFF]", container.value) and container.data_type != "f":
            try:
                translation = translate_str(container.value)
                if type(translation) == str:
                    container.value = translation
            except Exception as exception:
                colorprint("y", "translate_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, f"{container.value} - {exception}")
            else:
                colorprint("w", "translate_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, f"{container.value}")

    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "translate_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "translating")
        workbook = Workbook(str(path))
        for sheetname, (rows, cols) in worksheets_dimensions(str(path)).items():
            for row in range(1, rows + 2):
                for col in range(1, cols + 1):
                    cell = workbook[sheetname].cell(row, col)
                    _translate(cell)
        workbook.save(str(path))
        colorprint("g", "translate_xls", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "translated")


def translate_doc(paths: set[Path]) -> None:

    def _translate(container):
        if type(container.text) == str and search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(container.text)
                if type(translation) == str:
                    container.text = translation
            except Exception as exception:
                colorprint("y", "translate_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, f"{container.text} - {exception}")
            else:
                colorprint("w", "translate_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, f"{container.text}")

    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "translate_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "translating")
        document = Document(str(path))
        paragraphs = document.paragraphs
        tables = document.tables
        for section in document.sections:
            paragraphs += section.header.paragraphs + section.footer.paragraphs
            tables += section.header.tables + section.footer.tables
        for paragraph in paragraphs:
            for run in paragraph.runs:
                _translate(run)
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            _translate(run)
                    for subtable in cell.tables:
                        for subrow in subtable.rows:
                            for subcell in subrow.cells:
                                for subparagraph in subcell.paragraphs:
                                    for subrun in subparagraph.runs:
                                        _translate(subrun)
        document.save(str(path))
        colorprint("g", "translate_doc", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "translated")


def translate_ppt(paths: set[Path]) -> None:

    def _translate(container):
        if type(container.text) == str and search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(container.text)
                if type(translation) == str:
                    container.text = translation
            except Exception as exception:
                colorprint("y", "translate_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, f"{container.text} - {exception}")
            else:
                colorprint("w", "translate_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, f"{container.text}")

    time_beg = perf_counter()
    for index_path, path in enumerate(paths, start=1):
        time_cur = perf_counter()
        colorprint("c", "translate_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "translating")
        presentation = Presentation(str(path))
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            _translate(run)
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    _translate(run)
        presentation.save(str(path))
        colorprint("g", "translate_ppt", index_path, len(paths), int(perf_counter() - time_cur), int(perf_counter() - time_beg), path.name, "translated")


def translate(directory: Path = Path("output")) -> None:

    paths_xls = {path for path in directory.rglob("*") if path.suffix == ".xlsx" and path.stat().st_file_attributes != 34}
    paths_doc = {path for path in directory.rglob("*") if path.suffix == ".docx" and path.stat().st_file_attributes != 34}
    paths_ppt = {path for path in directory.rglob("*") if path.suffix == ".pptx" and path.stat().st_file_attributes != 34}

    process_xls = Process(target=translate_xls, args=[paths_xls])
    process_doc = Process(target=translate_doc, args=[paths_doc])
    process_ppt = Process(target=translate_ppt, args=[paths_ppt])

    process_xls.start()
    process_doc.start()
    process_ppt.start()

    process_xls.join()
    process_doc.join()
    process_ppt.join()
