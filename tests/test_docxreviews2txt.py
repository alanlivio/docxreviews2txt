import contextlib
from io import StringIO, open
import pathlib
import unittest
from os import listdir
from os.path import abspath, exists, join

from docxreviews2txt.docxreviews2txt import DocxReviews

TEST_FOLDER = "tests"

class TestCase(unittest.TestCase):
    def test_input_docx_files(self) -> None:
        files = [join(TEST_FOLDER, file) for file in listdir(TEST_FOLDER) if file.endswith('.docx')]
        for file in files:
            txt_out = file.replace('.docx', '_review.txt')
            txt_expected = file.replace('.docx', '_review_expected.txt')
            assert exists(txt_expected)
            docx_reviews = DocxReviews(file)
            output = StringIO()
            with contextlib.redirect_stdout(output):
                docx_reviews.save_reviews()
                cli_l = output.getvalue().split("\n")[:-1]
                cli_expected_l = [f"txt reviews at {pathlib.Path(abspath(txt_out)).as_uri()}"]
                self.assertEqual(cli_l, cli_expected_l)
            assert exists(txt_out)
            with open(txt_out) as f:
                ouput_l = f.read().splitlines()
            with open(txt_expected) as f:
                expected_l = f.read().splitlines()
            self.assertEqual(ouput_l, expected_l)