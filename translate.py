import re
import copy
import docx
import pptx
import pathlib
import openpyxl
import functools
import subprocess
import multiprocessing
import argostranslate.package

import utilities


LANGUAGE_SRC = "zh"
LANGUAGE_DST = "en"


def install_translator() -> None:
    argostranslate.package.update_package_index()
    available_packages = argostranslate.package.get_available_packages()
    package_to_install = next(filter(lambda package: package.from_code == LANGUAGE_SRC and package.to_code == LANGUAGE_DST, available_packages))
    argostranslate.package.install_from_path(package_to_install.download())


@functools.cache
def translate_str(safe: bool, text: str) -> str:
    translation = subprocess.run(args=["argos-translate", "--from-lang", LANGUAGE_SRC, "--to-lang", LANGUAGE_DST, text], capture_output=True, text=True, check=True, timeout=30).stdout.strip()
    return translation


def translate_xls(safe: bool, path: pathlib.Path, watch: utilities.Watch) -> None:

    def _translate(container) -> None:
        if isinstance(container.value, str) and re.search("[\u4E00-\u9FFF]", container.value) and container.data_type != "f":
            try:
                translation = translate_str(safe, container.value)
                if isinstance(translation, str): container.value = translation
            except Exception as exception:
                watch.print(f"{exception}")

    workbook = openpyxl.load_workbook(str(path))
    for sheetname, (rows, cols) in utilities.worksheets_dimensions(str(path)).items():
        for row in range(1, rows + 2):
            for col in range(1, cols + 1):
                cell = workbook[sheetname].cell(row, col)
                _translate(cell)
    workbook.save(str(path))


def translate_doc(safe: bool, path: pathlib.Path, watch: utilities.Watch) -> None:

    def _translate(container) -> None:
        if isinstance(container.text, str) and re.search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(safe, container.text)
                if isinstance(translation, str): container.text = translation
            except Exception as exception:
                watch.print(f"{exception}")

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


def translate_ppt(safe: bool, path: pathlib.Path, watch: utilities.Watch) -> None:

    def _translate(container) -> None:
        if isinstance(container.text, str) and re.search("[\u4E00-\u9FFF]", container.text):
            try:
                translation = translate_str(safe, container.text)
                if isinstance(translation, str): container.text = translation
            except Exception as exception:
                watch.print(f"{exception}")

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


def translate_mso(safe: bool, paths: list[tuple[int, pathlib.Path]], watch: utilities.Watch) -> None:

    suffix_to_function = {".xlsx": translate_xls, ".docx": translate_doc, ".pptx": translate_ppt}

    while paths:
        index, path = paths.pop(0)
        watch.current(index, path)
        watch.print("translating")
        function = suffix_to_function[path.suffix]
        function(safe, path, copy.deepcopy(watch))
        watch.print("translated")


def translate(safe: bool, directory_dst: pathlib.Path, processes: int, watch: utilities.Watch) -> None:

    install_translator()

    with multiprocessing.Manager() as manager:

        paths_raw = [path for path in directory_dst.rglob("*") if path.suffix in (".xlsx", ".docx", ".pptx")]
        paths_raw.sort(key=lambda path: path.stat().st_size, reverse=True)

        paths = manager.list(enumerate(paths_raw, start=1))
        watch.beginning(len(paths_raw))

        pool = [multiprocessing.Process(target=translate_mso, args=(safe, paths, copy.deepcopy(watch))) for _ in range(processes)]
        for process in pool: process.start()
        for process in pool: process.join()
