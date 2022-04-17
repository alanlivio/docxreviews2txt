from docxreviews2txt.docxreviews2txt import main
import unittest
import io
import contextlib


class TestCase(unittest.TestCase):
    def test_lorem_ipsum_docx(self):
        lorem_ipsum = [
            '# Typos and rewriting suggestions',
            '- sit amet, consectetur  -> sit amet, consectetur Lorem ipsum',
            '- sit amet, consectetur adipiscing elit, sed do -> sit amet, consectetur elit, sed do',
            '- sit amet, consectetur adipiscing elit, sed -> sit amet, consectetur adipiscings elit, sed',
            '- enim ad minim veniam, quis nostrud -> enim ad minim do veniam, quis nostrud',
            '- enim ad minim veniam -> enim ad minim Lorem veniam',
            '- veniam, quis nostrud -> veniam ipsum, quis nostrud',
            '- sit amet, consectetur adipiscing elit, sed do -> sit amet, consectetur elit, sed do'
        ]
        # redirect stdout (https://stackoverflow.com/questions/54824018/get-output-of-a-function-as-string)
        f = io.StringIO()
        with contextlib.redirect_stdout(f):
            # pass args to main (https://jugmac00.github.io/blog/testing-argparse-applications-the-better-way/)
            main(["tests/lorem_ipsum.docx"])
            self.assertEqual(lorem_ipsum, f.getvalue().split('\n')[:-1])


if __name__ == '__main__':
    unittest.main()
