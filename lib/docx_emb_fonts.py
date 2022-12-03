import logging
import shutil
import os
import re
from zipfile import ZipFile
from xml.etree import ElementTree
from pathlib import Path

logger = logging.getLogger(__name__)

DOCX_NS = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        }


class DocxReaderFont:
    def __init__(self, font_name, font_type, font_id, key):
        self.font_name = font_name
        self.font_type = font_type
        self.font_id = font_id
        self.key = key
        self.file_name = None


class DocxReaderException(Exception):
    pass


class DocxReader:
    def __init__(self, file_name):
        self._file_name = file_name
        self._zip_file: ZipFile | None = None
        self._font_list: list[DocxReaderFont] = []

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exception_type, exception_value, exception_traceback):
        self.close()

    def open(self):
        self._zip_file = ZipFile(self._file_name, "r")
        logger.info("docx file opened")

    def close(self):
        self._zip_file.close()
        logger.info("docx file closed")

    @staticmethod
    def _ns_tag(ns, tag_name):
        if ns in DOCX_NS:
            return "{" + DOCX_NS[ns] + "}" + tag_name
        else:
            raise DocxReaderException("Unknown namespace " + ns)

    def _find_embedded_font(self, node, font_name, tag_name, font_type):
        item = node.find(tag_name, DOCX_NS)
        if item is not None:
            font_id = item.get(self._ns_tag("r", "id")).replace("rId", "font")
            font_key = item.get(self._ns_tag("w", "fontKey"))
            self._font_list.append(DocxReaderFont(font_name, font_type, font_id, font_key))
            logger.info("    id = %s, type = %s, key = %s", font_id, font_type, font_key)

    def get_font_list(self):
        self._font_list.clear()
        logger.info("Reading font list")
        if self._zip_file:
            root = ElementTree.fromstring(self._zip_file.read("word/fontTable.xml").decode("UTF8"))
            for font_item in root.findall('w:font', DOCX_NS):
                font_name = font_item.get(self._ns_tag("w", "name"))
                logger.info("  Font found: %s", font_name)
                self._find_embedded_font(font_item, font_name, "w:embedRegular", "Regular")
                self._find_embedded_font(font_item, font_name, "w:embedBold", "Bold")
                self._find_embedded_font(font_item, font_name, "w:embedItalic", "Italic")
                self._find_embedded_font(font_item, font_name, "w:embedBoldItalic", "Bold Italic")
        else:
            raise DocxReaderException("docx file must be opened first")

        return self._font_list

    def _extract_font(self, font_item, target_dir):
        font_file = font_item.font_id + ".odttf"
        target_file = os.path.join(target_dir, font_item.font_name + " " + font_item.font_type + ".odttf")
        font_item.file_name = target_file
        with self._zip_file.open('word/fonts/' + font_file) as zf, open(target_file, 'wb') as tf:
            shutil.copyfileobj(zf, tf)
            logger.info("  Extracted file: %s", target_file)

    def extract_all(self, target_dir=""):
        logger.info("Extracting all fonts")
        if self._zip_file:
            if not self._font_list:
                self.get_font_list()

            if target_dir:
                Path(target_dir).mkdir(parents=True, exist_ok=True)
            else:
                target_dir = ""

            for embedded_font in self._font_list:
                self._extract_font(embedded_font, target_dir)

            return self._font_list
        else:
            raise DocxReaderException("docx file must be opened first")

    def extract(self, target_dir="", *, name_type_list=None, id_list=None):
        logger.info("Extracting list of fonts")
        if self._zip_file:
            if not self._font_list:
                self.get_font_list()

            if target_dir:
                Path(target_dir).mkdir(parents=True, exist_ok=True)
            else:
                target_dir = ""

            font_list = []
            if name_type_list:
                for index in name_type_list:
                    for embedded_font in self._font_list:
                        # if index == embedded_font.font_name + " " + embedded_font.font_type:
                        if re.match(index, embedded_font.font_name + " " + embedded_font.font_type):
                            font_list.append(embedded_font)
            elif id_list:
                for index in id_list:
                    for embedded_font in self._font_list:
                        if index == embedded_font.font_id:
                            font_list.append(embedded_font)
                            break
            else:
                raise DocxReaderException("No list provided")

            for embedded_font in font_list:
                self._extract_font(embedded_font, target_dir)

            return font_list
        else:
            raise DocxReaderException("docx file must be opened first")
