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
        files = [
            join(TEST_FOLDER, file)
            for file in listdir(TEST_FOLDER)
            if file.startswith("input_") and file.endswith(".docx")
        ]
        for file in files:
            for fmt in ["tags", "diff"]:
                suffix = "_tags_expected.txt" if fmt == "tags" else "_diff_expected.txt"
                txt_expected = file.replace(".docx", f"_review{suffix}")
                txt_out = file.replace(".docx", f"_review_{fmt}.txt")

                assert exists(txt_expected)
                docx_reviews = DocxReviews(file, output_format=fmt)

                output = StringIO()
                with contextlib.redirect_stdout(output):
                    docx_reviews.save_reviews()

                real_out = file.replace(".docx", "_review.txt")
                assert exists(real_out)

                with open(real_out) as f:
                    output_l = f.read().splitlines()
                with open(txt_expected) as f:
                    expected_l = f.read().splitlines()

                self.assertEqual(output_l, expected_l, f"Failed for {file} in format {fmt}")
