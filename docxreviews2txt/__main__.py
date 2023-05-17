import argparse
import pathlib
import sys

from docxreviews2txt import DocxReviews, DEFAULT_NWORDS, __version__


def main() -> None:
  parser = argparse.ArgumentParser(
      prog='docxreviews2txt',
      description="Extract review changes and comments from a docx file as plain text.")
  parser.add_argument("docx", help="input docx", type=pathlib.Path)
  parser.add_argument('-nwords',
                      nargs='?',
                      type=int,
                      default=DEFAULT_NWORDS,
                      help=f'words around each change (default: {DEFAULT_NWORDS})')
  parser.add_argument('--version',
                      help='show version',
                      action='version',
                      version='%(prog)s ' + __version__)
  argv = sys.argv[1:]
  args = parser.parse_args(argv)
  docx_reviews = DocxReviews(file_docx=args.docx, nwords=args.nwords)
  docx_reviews.save_reviews_to_file()


if __name__ == "__main__":
  main()
