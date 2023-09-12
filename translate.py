import re
import time
import copy
import docx
import pptx
import pathlib
import openpyxl
import functools
import multiprocessing
import deep_translator
import pebble.concurrent

import utilities


@functools.cache
@pebble.concurrent.process(timeout=30)
def translate_str(text: str) -> str:
    return deep_translator.GoogleTranslator(source="auto", target="en").translate(text)


def translate_xls(path: pathlib.Path, watch: utilities.Watch) -> None:

    def _translate(container):
        if type(container.value) == str and re.search("[\u4E00-\u9FFF]", container.value) and container.data_type != "f":
            try:
                translation = translate_str(container.value).result()
                if type(translation) == str:
                    container.value = translation
            except Exception as exception:
                utilities.colorprint("y", path.name, f"{container.value} - {exception}", copy.deepcopy(watch))
            else:
                utilities.colorprint("w", path.name, f"{container.value}", copy.deepcopy(watch))

    workbook = openpyxl.load_workbook(str(path))
    for sheetname, (rows, cols) in utilities.worksheets_dimensions(str(path)).items():
        for row in range(1, rows + 2):
            for col in range(1, cols + 1):
                cell = workbook[sheetname].cell(row, col)
                _translate(cell)
    workbook.save(str(path))


def translate_doc(path: pathlib.Path, watch: utilities.Watch) -> None:

    def _translate(container):
        if type(container.text) == str and re.search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(container.text).result()
                if type(translation) == str:
                    container.text = translation
            except Exception as exception:
                utilities.colorprint("y", path.name, f"{container.text} - {exception}", copy.deepcopy(watch))
            else:
                utilities.colorprint("w", path.name, f"{container.text}", copy.deepcopy(watch))

    document = docx.Document(str(path))
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


def translate_ppt(path: pathlib.Path, watch: utilities.Watch) -> None:

    def _translate(container):
        if type(container.text) == str and re.search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(container.text).result()
                if type(translation) == str:
                    container.text = translation
            except Exception as exception:
                utilities.colorprint("y", path.name, f"{container.text} - {exception}", copy.deepcopy(watch))
            else:
                utilities.colorprint("w", path.name, f"{container.text}", copy.deepcopy(watch))

    presentation = pptx.Presentation(str(path))
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


def translate_mso(paths: list[pathlib.Path], watch: utilities.Watch) -> None:

    suffix_to_function = {".xlsx": translate_xls, ".docx": translate_doc, ".pptx": translate_ppt}

    watch.index_ending = len(paths)
    for index, path in enumerate(paths, start=1):
        watch.index_current = index
        watch.time_current = time.perf_counter()
        utilities.colorprint("c", path.name, "translating", copy.deepcopy(watch))
        function = suffix_to_function[path.suffix]
        function(path, copy.deepcopy(watch))
        utilities.colorprint("g", path.name, "translated", copy.deepcopy(watch))


def translate(directory: pathlib.Path = pathlib.Path("output"), watch: utilities.Watch = utilities.Watch()) -> None:

    paths = [path for path in directory.rglob("*") if path.suffix in (".xlsx", ".docx", ".pptx") and path.stat().st_file_attributes != 34]

    processes = [multiprocessing.Process(target=translate_mso, args=(values, copy.deepcopy(watch))) for values in utilities.list_split(paths)]
    for process in processes: process.start()
    for process in processes: process.join()
