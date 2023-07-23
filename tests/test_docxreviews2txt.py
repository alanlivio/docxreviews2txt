import contextlib
from io import StringIO, open
import pathlib
import unittest
from os.path import abspath, exists, join

import docxreviews2txt

DOCX = join("tests", "lorem_ipsum.docx")
TXT_OUT = "tests/lorem_ipsum_review.txt"
TXT_EXPECTED = "tests/lorem_ipsum_expected.txt"
XML_OUT = "tests/lorem_ipsum.xml"
XML_EXPECTED = "tests/lorem_ipsum_expected.xml"


class TestCase(unittest.TestCase):

  def test_lorem_ipsum_txt(self) -> None:
    assert exists(TXT_EXPECTED)
    docx_reviews = docxreviews2txt.DocxReviews(DOCX)
    output = StringIO()
    with contextlib.redirect_stdout(output):
      docx_reviews.save_reviews_to_file()
      cli_l = output.getvalue().split('\n')[:-1]
      cli_expected_l = [f'txt reviews at {pathlib.Path(abspath(TXT_OUT)).as_uri()}']
      self.assertEqual(cli_l, cli_expected_l)
    assert exists(TXT_OUT)
    with open(TXT_OUT) as f:
      ouput_l = f.read().splitlines()
    with open(TXT_EXPECTED) as f:
      expected_l = f.read().splitlines()
    self.assertEqual(ouput_l, expected_l)

  def test_lorem_ipsum_xml_p_elems(self) -> None:
    assert exists(XML_EXPECTED)
    docx_reviews = docxreviews2txt.DocxReviews(DOCX)
    output = StringIO()
    with contextlib.redirect_stdout(output):
      docx_reviews.save_xml_p_elems()
      cli_l = output.getvalue().split('\n')[:-1]
      cli_expected_l = [f'xml paragraphs at {pathlib.Path(abspath(XML_OUT)).as_uri()}']
      self.assertEqual(cli_l, cli_expected_l)
    assert exists(XML_OUT)
    with open(XML_OUT) as f:
      ouput_l = f.read().splitlines()
    with open(XML_EXPECTED) as f:
      expected_l = f.read().splitlines()
    self.assertEqual(ouput_l, expected_l)

  def test_lorem_ipsum_docxreviews_cli(self) -> None:
    output = StringIO()
    with contextlib.redirect_stdout(output):
      argv = [DOCX]
      docxreviews2txt.docxreviews_cli(argv)
      cli_l = output.getvalue().split('\n')[:-1]
      cli_expected_l = [f'txt reviews at {pathlib.Path(abspath(TXT_OUT)).as_uri()}']
      self.assertEqual(cli_l, cli_expected_l)
    assert exists(TXT_OUT)
    with open(TXT_OUT) as f:
      ouput_l = f.read().splitlines()
    with open(TXT_EXPECTED) as f:
      expected_l = f.read().splitlines()
    self.assertEqual(ouput_l, expected_l)
