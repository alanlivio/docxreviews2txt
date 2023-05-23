import os
import pathlib
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from os.path import abspath, exists, join, splitext

from docx import Document

__version__ = '0.4.1'

WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_MAP = {"w": WORD_NS}
ET_WORD_NS = '{' + WORD_NS + '}'
ET_TXT = ET_WORD_NS + 't'
ET_DEL = ET_WORD_NS + 'del'
ET_INS = ET_WORD_NS + 'ins'
DEFAULT_NWORDS = 4


class DocxReviews:

  def __init__(self, file_docx, nwords=DEFAULT_NWORDS) -> None:
    assert exists(file_docx)
    self.reviews = []
    self.file_docx = abspath(file_docx)
    self.nwords = nwords
    # use tmp file
    self.target_file = join(tempfile.gettempdir(), 'docx_reviews_to_txt.docx')
    if exists(self.target_file):
      os.remove(self.target_file)
      assert not exists(self.target_file)
    try:
      shutil.copyfile(file_docx, self.target_file)
    except Exception as exc:
      # at windows, shutil.copy fail if docx opened and only can be copied from powershell
      if os.name == 'nt':
        cmd = f"Copy-Item {file_docx} {self.target_file}"
        subprocess.run(["powershell", "-Command", cmd], capture_output=True, check=True)
      else:
        raise exc
    assert exists(self.target_file)
    self.paragraphs = Document(self.target_file).paragraphs

  def _str_deltext_elms(self, child) -> str:
    x = [text.text for text in child.findall('.//w:delText', NS_MAP)]
    return "".join(x)

  def _str_t_elms(self, child) -> str:
    x = [text.text for text in child.findall('.//w:t', NS_MAP)]
    return "".join(x)

  def _str_left_t_elms(self, root_p, index) -> str:
    left_ar = []
    for i in range(index, 0, -1):
      if root_p[i].tag == ET_INS:
        continue
      left_ar = self._str_t_elms(root_p[i]).split(" ") + left_ar
      if len(left_ar) >= self.nwords:
        left_ar = left_ar[-self.nwords:]
      break
    return " ".join(left_ar)

  def _str_right_t_elms(self, root_p, index) -> str:
    right_ar = []
    for i in range(index, len(root_p)):
      if root_p[i].tag == ET_INS:
        continue
      right_ar = self._str_t_elms(root_p[i]).split(" ") + right_ar
      if len(right_ar) >= self.nwords:
        right_ar = right_ar[:self.nwords]
      break
    return " ".join(right_ar)

  def _append(self, text) -> None:
    if len(text) == 0:
      return
    self.reviews.append(text)

  def _parse(self) -> None:
    # comments
    with zipfile.ZipFile(self.target_file, mode="r") as docxZip:
      if 'word/comments.xml' in [member.filename for member in docxZip.infolist()]:
        commentsXML = docxZip.read('word/comments.xml')
        root = ET.fromstring(commentsXML)
        comments = root.findall('.//w:comment', NS_MAP)
        if len(comments):
          self._append("# Comments")
          for comment in comments:
            lines = comment.findall('.//w:r', NS_MAP)
            for line in lines:
              self._append(self._str_t_elms(line))

    # changes
    self._append("# Typos and rewriting suggestions")
    for p in self.paragraphs:
      xml = p._p.xml
      root = ET.fromstring(xml)
      if len(root.findall('.//w:del', NS_MAP)) == 0 \
        and len(root.findall('.//w:ins', NS_MAP)) == 0:
        continue
      for index in range(len(root) - 1):
        prev_w = root[index - 1]
        cur_w = root[index]
        next_w = root[index + 1]
        # DEL followed by INS
        if cur_w.tag == ET_DEL and next_w.tag == ET_INS:
          del_txt = self._str_deltext_elms(cur_w)
          ins_txt = self._str_t_elms(next_w)
          left_txt = self._str_left_t_elms(root, index - 1)
          right_txt = self._str_right_t_elms(root, index + 2)
          self._append("- " + left_txt + del_txt + right_txt + " -> " + left_txt + ins_txt +
                       right_txt)
        # INS alone
        elif cur_w.tag == ET_INS and prev_w.tag != ET_DEL:
          ins_txt = self._str_t_elms(cur_w)
          left_txt = self._str_left_t_elms(root, index - 1)
          right_txt = self._str_right_t_elms(root, index)
          self._append("- " + left_txt + right_txt + " -> " + left_txt + ins_txt + right_txt)
        # DEL alone
        elif cur_w.tag == ET_DEL and next_w.tag != ET_INS:
          del_txt = self._str_deltext_elms(cur_w)
          left_txt = self._str_left_t_elms(root, index - 1)
          right_txt = self._str_right_t_elms(root, index + 1)
          self._append("- " + left_txt + del_txt + right_txt + " -> " + left_txt + right_txt)

  def save_reviews_to_file(self) -> None:
    if not self.reviews:
      self._parse()
    filename = splitext(self.file_docx)[0] + '_review.txt'
    with open(filename, "w") as file:
      for change in self.reviews:
        file.write(f"{change}\n")
    assert filename
    print(f'txt reviews at {pathlib.Path(filename).as_uri()}')

  def save_xml_p_elems(self) -> None:
    filename = splitext(self.file_docx)[0] + '.xml'
    with open(filename, "w") as file:
      for p in self.paragraphs:
        xml = p._p.xml
        file.write(f"{xml}\n")
    assert filename
    print(f'xml paragraphs at {pathlib.Path(filename).as_uri()}')
