import os
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from os.path import abspath, exists, join, splitext
import pathlib
import argparse
from docx import Document
from . import __version__

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_MAP = {"w": WORD_NS}
ET_WORD_NS = "{" + WORD_NS + "}"
ET_TXT = ET_WORD_NS + "t"
ET_PPR = ET_WORD_NS + "pPr"
ET_DEL = ET_WORD_NS + "del"
ET_INS = ET_WORD_NS + "ins"
NWORDS_START = 4
INS_BEGIN, INS_END, DEL_BEGIN, DEL_END = "<ins>", "</ins>", "<del>", "</del>"


def str_from_deltext_elems(child) -> str:
    deltext_s = ""
    for elem in child.findall(".//w:delText", NS_MAP):
        deltext_s = deltext_s + elem.text
    return deltext_s


def str_from_t_elems(child) -> str:
    text_s = ""
    for elem in child.findall(".//w:t", NS_MAP):
        text_s = text_s + elem.text
    return text_s


def str_surround_ins(txt) -> str:
    return INS_BEGIN + txt.strip() + INS_END


def str_surround_del(txt) -> str:
    return DEL_BEGIN + txt.strip() + DEL_END


class DocxReviews:
    def __init__(self, file_docx) -> None:
        assert exists(file_docx)
        self.reviews = []
        self.file_docx = abspath(file_docx)
        # use tmp file
        self.target_file = join(tempfile.gettempdir(), "docx_reviews_to_txt.docx")
        if exists(self.target_file):
            os.remove(self.target_file)
            assert not exists(self.target_file)
        try:
            shutil.copyfile(file_docx, self.target_file)
        except Exception as exc:
            # at windows, shutil.copy fail if docx opened and only can be copied from powershell
            if os.name == "nt":
                cmd = f"Copy-Item {file_docx} {self.target_file}"
                subprocess.run(["powershell", "-Command", cmd], capture_output=True, check=True)
            else:
                raise exc
        assert exists(self.target_file)
        self.paragraphs = Document(self.target_file).paragraphs

    def _parse(self) -> None:
        self.reviews.append("# Typos suggestions (using HTML tags <ins> and <del>)")
        for p in self.paragraphs:
            texts = []
            root = ET.fromstring(p._p.xml)

            # skip paragraph if no w:delText, w:ins, elems
            if not len(root.findall(".//w:del", NS_MAP)) and not len(
                root.findall(".//w:ins", NS_MAP)
            ):
                continue

            # short first elem to NWORDS if text
            index = 1  # skip index 0 (ppr elem)
            elem = root[index]
            if not elem.tag == ET_DEL and not elem.tag == ET_INS:
                text_s = str_from_t_elems(elem)
                text_s_as_ar = text_s.split()
                if len(text_s_as_ar) > NWORDS_START:
                    sufix = " " if text_s[-1] == " " else ""  # check ending with space
                    texts.append(" ".join(text_s_as_ar[-NWORDS_START:]) + sufix)
                else:
                    texts.append(text_s)
                # print(f"1:'{texts[0]}'")
                index = 2

            # find w:delText, w:ins, or text elems
            for index in range(index, len(root) - 1):  # skip index 0 (ppr elem)
                elem = root[index]
                if elem.tag == ET_DEL:  # it is del elem
                    result = str_surround_del(str_from_deltext_elems(elem))
                elif elem.tag == ET_INS:  # it is ins elem
                    result = str_surround_ins(str_from_t_elems(elem))
                else:  # considerer only text
                    result = str_from_t_elems(elem)
                # print(str(index) + ":'" + result + "'")
                texts.append(result)

            # add review line
            self.reviews.append("- " + "".join(texts))

    def save_reviews(self) -> None:
        if not self.reviews:
            self._parse()
        filename = splitext(self.file_docx)[0] + "_review.txt"
        with open(filename, "w") as file:
            for change in self.reviews:
                file.write(f"{change}\n")
        assert filename
        print(f"txt reviews at {pathlib.Path(filename).as_uri()}")

    def save_xml_p_elems(self) -> None:
        filename = splitext(self.file_docx)[0] + ".xml"
        with open(filename, "w") as file:
            for p in self.paragraphs:
                xml = p._p.xml
                file.write(f"{xml}\n")
        assert filename
        print(f"xml paragraphs at {pathlib.Path(filename).as_uri()}")


def docxreviews_cli(argv=None) -> None:
    parser = argparse.ArgumentParser(
        prog="docxreviews2txt",
        description="Command line tool to extract review changes from a docx file as plain text",
    )
    parser.add_argument("docx", help="input docx", type=pathlib.Path)
    parser.add_argument(
        "--version", help="show version", action="version", version="%(prog)s " + __version__
    )
    args = parser.parse_args(argv)
    docx_reviews = DocxReviews(file_docx=args.docx)
    docx_reviews.save_reviews()
