from os import listdir
from os.path import exists, join
import pytest

from docxreviews2txt.docxreviews2txt import DocxReviews

TEST_FOLDER = "tests"

INPUT_FILES = [
    join(TEST_FOLDER, d, "input.docx")
    for d in sorted(listdir(TEST_FOLDER))
    if d.startswith("sample_") and exists(join(TEST_FOLDER, d, "input_review_diff_expected.txt"))
]


@pytest.mark.parametrize("file", INPUT_FILES)
@pytest.mark.parametrize("fmt", ["tags", "diff"])
def test_input_docx_files(file: str, fmt: str, capsys: pytest.CaptureFixture[str]) -> None:
    suffix = "_tags_expected.txt" if fmt == "tags" else "_diff_expected.txt"
    txt_expected = file.replace(".docx", f"_review{suffix}")

    assert exists(txt_expected)
    docx_reviews = DocxReviews(file, output_format=fmt)
    docx_reviews.save_reviews()
    capsys.readouterr()

    real_out = file.replace(".docx", "_review.txt")
    assert exists(real_out)

    with open(real_out) as f:
        output_l = f.read().splitlines()
    with open(txt_expected) as f:
        expected_l = f.read().splitlines()

    assert output_l == expected_l, f"Failed for {file} in format {fmt}"

