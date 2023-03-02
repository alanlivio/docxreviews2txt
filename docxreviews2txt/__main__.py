#!/usr/bin/env python

import argparse
import pathlib
import sys

import docxreviews2txt


def main()  -> None:
    parser = argparse.ArgumentParser(prog='docxreviews2txt',
        description="Extract review changes and comments from a docx file as plain text.")
    parser.add_argument("docx", help="input docx", type=pathlib.Path)
    parser.add_argument('--save_p_xml', help='also save extracted Docx paragraphs as xml for debugging', action="store_true")
    parser.add_argument('--version', help='show version', action='version',
                        version='%(prog)s ' + docxreviews2txt.__version__)
    argv = sys.argv[1:]
    args = parser.parse_args(argv)
    docx_reviews = docxreviews2txt.DocxReviews(args.docx)
    docx_reviews.save_reviews_to_file()
    if args.save_p_xml:
        docx_reviews.save_xml_p_elems()

if __name__ == "__main__":
    main()
