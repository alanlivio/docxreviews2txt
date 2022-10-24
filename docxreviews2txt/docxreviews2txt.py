from docx import Document
import xml.etree.ElementTree as ET
import os
import subprocess
from os.path import exists
import zipfile
import tempfile
import shutil

__version__ = '0.3'

WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_MAP = {"w": WORD_NS}
ET_WORD_NS = '{' + WORD_NS + '}'
ET_TEXT = ET_WORD_NS + 't'
ET_DEL = ET_WORD_NS + 'del'
ET_INS = ET_WORD_NS + 'ins'
DEFAULT_WORDS_AROUND_CHANGE = 4


class DocxReviews:
    def __init__(self, file_docx, verbose, words_around_change=DEFAULT_WORDS_AROUND_CHANGE) -> None:
        self.verbose = verbose
        self.reviews = []
        self.file_docx = file_docx
        self.words_around_change = words_around_change
        # extract docxZip and paragraphs
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, 'docx_reviews_to_txt.docx')
        if exists(temp_path):
            os.remove(temp_path)
        try:
            shutil.copy(file_docx, temp_path)
        except:
            # at windows, shutil.copy fail if docx opened and only can be copied from powershell
            if os.name == 'nt':
                cmd = f"Copy-Item {file_docx} {temp_path}"
                subprocess.run(["powershell", "-Command", cmd], capture_output=True)
        self.docxZip = zipfile.ZipFile(temp_path, mode="r")
        self.paragraphs = Document(temp_path).paragraphs

    def str_deltext_elms(self, child) ->str:
        x = [text.text for text in child.findall(
            './/w:delText', NS_MAP)]
        return "".join(x)

    def str_t_elms(self, child)->str:
        x = [text.text for text in child.findall(
            './/w:t', NS_MAP)]
        return "".join(x)

    def str_left_t_elms(self, root_p, index)->str:
        left_ar = []
        for i in range(index, 0, -1):
            if (root_p[i].tag == ET_INS):
                continue
            left_ar = self.str_t_elms(root_p[i]).split(" ") + left_ar
            if(len(left_ar) >= self.words_around_change):
                left_ar = left_ar[-self.words_around_change:]
            break
        return " ".join(left_ar)

    def str_right_t_elms(self, root_p, index)->str:
        right_ar = []
        for i in range(index, len(root_p)):
            if (root_p[i].tag == ET_INS):
                continue
            right_ar = self.str_t_elms(root_p[i]).split(" ") + right_ar
            if(len(right_ar) >= self.words_around_change):
                right_ar = right_ar[:self.words_around_change]
            break
        return " ".join(right_ar)

    def reviews_append(self, text) -> None:
        if not len(text):
            return
        self.reviews.append(text)
        if self.verbose:
            print(self.reviews[-1])

    def parse(self) -> None:
        # comments
        if 'word/comments.xml' in [member.filename for member in self.docxZip.infolist()]:
            commentsXML = self.docxZip.read('word/comments.xml')
            # print(commentsXML)
            root = ET.fromstring(commentsXML)
            comments = root.findall('.//w:comment', NS_MAP)
            if len(comments):
                self.reviews_append("# Comments")
                for comment in comments:
                    lines = comment.findall('.//w:r', NS_MAP)
                    for line in lines:
                        self.reviews_append(self.str_t_elms(line))

        # changes
        self.reviews_append("# Typos and rewriting suggestions")
        for p in self.paragraphs:
            xml = p._p.xml
            root = ET.fromstring(xml)
            if len(root.findall('.//w:del', NS_MAP)) == 0 and len(root.findall('.//w:ins', NS_MAP)) == 0:
                continue
            for index in range(len(root) - 1):
                prev = root[index - 1]
                cur = root[index]
                next = root[index + 1]
                # DEL followed by INS
                if (cur.tag == ET_DEL and next.tag == ET_INS):
                    del_text = self.str_deltext_elms(cur)
                    ins_text = self.str_t_elms(next)
                    left_text = self.str_left_t_elms(root, index - 1)
                    right_text = self.str_right_t_elms(root, index + 2)
                    self.reviews_append("- " + left_text + del_text + right_text +
                                        " -> " + left_text + ins_text + right_text)
                # INS alone
                elif (cur.tag == ET_INS and prev.tag != ET_DEL):
                    ins_text = self.str_t_elms(cur)
                    left_text = self.str_left_t_elms(root, index - 1)
                    right_text = self.str_right_t_elms(root, index)
                    self.reviews_append("- " + left_text + right_text +
                                        " -> " + left_text + ins_text + right_text)
                # DEL alone
                elif (cur.tag == ET_DEL and next.tag != ET_INS):
                    del_text = self.str_deltext_elms(cur)
                    left_text = self.str_left_t_elms(root, index - 1)
                    right_text = self.str_right_t_elms(root, index + 1)
                    self.reviews_append("- " + left_text + del_text + right_text +
                                        " -> " + left_text + right_text)

    def save_reviews_to_file(self):
        file_txt_name = str(os.path.splitext(self.file_docx)[0]) + '_review.txt'
        with open(file_txt_name, "w") as file:
            for change in self.reviews:
                file.write(f"{change}\n")

    def save_xml_p_elems(self):
        file_txt_name = str(os.path.splitext(self.file_docx)[0]) + '.xml'
        with open(file_txt_name, "w") as file:
            for p in self.paragraphs:
                xml = p._p.xml
                file.write(f"{xml}\n")
