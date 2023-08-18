from re import search
from copy import deepcopy
from time import perf_counter
from pathlib import Path
from functools import cache
from multiprocessing import Process
from deep_translator import GoogleTranslator
from pebble.concurrent import process

from docx import Document
from pptx import Presentation
from openpyxl import load_workbook as Workbook

from utilities import Watch, colorprint, worksheets_dimensions, list_split


@cache
@process(timeout=30)
def translate_str(text: str) -> str:
    return GoogleTranslator(source="auto", target="en").translate(text)


def translate_xls(path: Path, watch: Watch) -> None:

    def _translate(container):
        if type(container.value) == str and search("[\u4E00-\u9FFF]", container.value) and container.data_type != "f":
            try:
                translation = translate_str(container.value).result()
                if type(translation) == str:
                    container.value = translation
            except Exception as exception:
                colorprint("y", path.name, f"{container.value} - {exception}", deepcopy(watch))
            else:
                colorprint("w", path.name, f"{container.value}", deepcopy(watch))

    workbook = Workbook(str(path))
    for sheetname, (rows, cols) in worksheets_dimensions(str(path)).items():
        for row in range(1, rows + 2):
            for col in range(1, cols + 1):
                cell = workbook[sheetname].cell(row, col)
                _translate(cell)
    workbook.save(str(path))


def translate_doc(path: Path, watch: Watch) -> None:

    def _translate(container):
        if type(container.text) == str and search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(container.text).result()
                if type(translation) == str:
                    container.text = translation
            except Exception as exception:
                colorprint("y", path.name, f"{container.text} - {exception}", deepcopy(watch))
            else:
                colorprint("w", path.name, f"{container.text}", deepcopy(watch))

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


def translate_ppt(path: Path, watch: Watch) -> None:

    def _translate(container):
        if type(container.text) == str and search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(container.text).result()
                if type(translation) == str:
                    container.text = translation
            except Exception as exception:
                colorprint("y", path.name, f"{container.text} - {exception}", deepcopy(watch))
            else:
                colorprint("w", path.name, f"{container.text}", deepcopy(watch))

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


def translate_mso(paths: list[Path], watch: Watch) -> None:

    suffix_to_function = {".xlsx": translate_xls, ".docx": translate_doc, ".pptx": translate_ppt}

    watch.index_ending = len(paths)
    for index, path in enumerate(paths, start=1):
        watch.index_current = index
        watch.time_current = perf_counter()
        colorprint("c", path.name, "translating", deepcopy(watch))
        function = suffix_to_function[path.suffix]
        function(path, deepcopy(watch))
        colorprint("g", path.name, "translated", deepcopy(watch))


def translate(directory: Path = Path("output"), watch: Watch = Watch()) -> None:

    paths = [path for path in directory.rglob("*") if path.suffix in (".xlsx", ".docx", ".pptx") and path.stat().st_file_attributes != 34]

    processes = [Process(target=translate_mso, args=(values, deepcopy(watch))) for values in list_split(paths)]
    for process in processes: process.start()
    for process in processes: process.join()
