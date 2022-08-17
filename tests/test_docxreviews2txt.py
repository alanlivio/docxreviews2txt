import contextlib
import io
import unittest
import docxreviews2txt


class TestCase(unittest.TestCase):
    def test_lorem_ipsum_docx(self):
        lorem_ipsum = [
            '# Typos and rewriting suggestions',
            '- sit amet, consectetur  -> sit amet, consectetur Lorem ipsum',
            '- sit amet, consectetur adipiscing elit, sed do eiusmod -> sit amet, consectetur elit, sed do eiusmod',
            '- sit amet, consectetur adipiscing elit, sed do -> sit amet, consectetur adipiscings elit, sed do',
            '- enim ad minim veniam, quis nostrud exercitation -> enim ad minim do veniam, quis nostrud exercitation',
            '- enim ad minim veniam -> enim ad minim Lorem veniam',
            '- veniam, quis nostrud exercitation -> veniam ipsum, quis nostrud exercitation',
            '- sit amet, consectetur adipiscing elit, sed do eiusmod -> sit amet, consectetur elit, sed do eiusmod',
        ]
        # redirect stdout (https://stackoverflow.com/questions/54824018/get-output-of-a-function-as-string)
        f = io.StringIO()
        with contextlib.redirect_stdout(f):
            # pass args to main (https://jugmac00.github.io/blog/testing-argparse-applications-the-better-way/)
            docx_reviews = docxreviews2txt.DocxReviews("tests/lorem_ipsum.docx", True)
            docx_reviews.parse()
            output = f.getvalue().split('\n')[:-1]
            self.assertEqual(len(lorem_ipsum), len(output))
            self.assertEqual(lorem_ipsum, output)


if __name__ == '__main__':
    unittest.main()
