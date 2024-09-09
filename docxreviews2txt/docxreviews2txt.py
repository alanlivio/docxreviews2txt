import argparse
import os
import pathlib
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from os.path import abspath, exists, join, splitext

from .version import __version__
from docx import Document


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_MAP = {"w": WORD_NS}
ET_WORD_NS = "{" + WORD_NS + "}"
ET_TXT = ET_WORD_NS + "t"
ET_PPR = ET_WORD_NS + "pPr"
ET_DEL = ET_WORD_NS + "del"
ET_INS = ET_WORD_NS + "ins"
NWORDS_AROUND = 4
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
    return INS_BEGIN + txt + INS_END


def str_surround_del(txt) -> str:
    return DEL_BEGIN + txt + DEL_END


def _get_left_txt(root, index) -> str:
    res = ""
    while index > 0:
        elem = root[index]
        if elem.tag == ET_DEL or elem.tag == ET_INS:
            break
        text_s = str_from_t_elems(elem)
        res = text_s + res
        index = -1
    return res


def _get_right_txt(root, index, limit) -> tuple[str, bool]:
    res = ""
    change_ahead = False
    while index < limit:
        elem = root[index]
        if elem.tag == ET_DEL or elem.tag == ET_INS:
            change_ahead = True
            break
        text_s = str_from_t_elems(elem)
        res = res + text_s
        index += 1
    return res, change_ahead


def _get_tagged_change(root, index) -> str:
    elem = root[index]
    if elem.tag == ET_DEL:
        string = str_from_deltext_elems(elem)
        return str_surround_del(string) if string else ""
    elif elem.tag == ET_INS:
        string = str_from_t_elems(elem)
        return str_surround_ins(string) if string else ""
    else:
        return "<error>"


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
                subprocess.run(
                    ["powershell", "-noprofile", "-Command", cmd], capture_output=True, check=True
                )
            else:
                raise exc
        assert exists(self.target_file)
        self.paragraphs = Document(self.target_file).paragraphs

    def _parse(self) -> None:
        self.reviews.append("Typos suggestions using HTML tags <ins> and <del>:")
        for p in self.paragraphs:
            root = ET.fromstring(p._p.xml)
            if not len(root.findall(".//w:del", NS_MAP)) and not len(
                root.findall(".//w:ins", NS_MAP)
            ):
                continue
            limit = len(root)
            index = 0
            left, right, result = "", "", ""
            change_ahead = False
            while index < limit:
                if root[index].tag == ET_DEL or root[index].tag == ET_INS:
                    # if not consecutive change, save left txt and change on result
                    if not change_ahead:
                        left = _get_left_txt(root, index - 1)
                        if len(left.split()) > NWORDS_AROUND:
                            left = " ".join(left.split()[-NWORDS_AROUND:]) + (
                                " " if left[-1] == " " else ""
                            )
                        result = left + result
                    result = result + _get_tagged_change(root, index)
                    # look ahead
                    right, change_ahead = _get_right_txt(root, index + 1, limit)
                    right_len = len(right.split())
                    # if change_ahead near, concatenate result
                    if change_ahead and right_len < NWORDS_AROUND:
                        index += 1
                        continue
                    # if change_ahead far, finish result
                    elif change_ahead and right_len > NWORDS_AROUND:
                        right = (" " if right[0] == " " else "") + " ".join(
                            right.split()[:NWORDS_AROUND]
                        )
                        result = result + right
                        self.reviews.append("- " + result)
                        left, right, result = "", "", ""
                        change_ahead = False
                    # if no change_ahead, finish result
                    if not change_ahead and result:
                        if right_len > NWORDS_AROUND:
                            right = (" " if right[0] == " " else "") + " ".join(
                                right.split()[:NWORDS_AROUND]
                            )
                        result = result + right
                        self.reviews.append("- " + result)
                        left, right, result = "", "", ""
                        change_ahead = False
                index += 1

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
        description="Command line tool to extract review changes from a docx file as plain text using HTML tags <ins> and <del>.",
    )
    parser.add_argument("docx", help="input docx", type=pathlib.Path)
    parser.add_argument(
        "--version", help="show version", action="version", version="%(prog)s " + __version__
    )
    args = parser.parse_args(argv)
    docx_reviews = DocxReviews(file_docx=args.docx)
    docx_reviews.save_reviews()
