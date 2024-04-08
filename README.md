# docxreviews2txt

Command line tool to extract review changes from a docx file as plain text. It is useful when reviewing a PDF file as docx, and you need to share the changes as plain text.

## How to install?

```bash
pip install docxreviews2txt
```

## How to use it?

```txt
usage: docxreviews2txt [-h] [--version] docx

Command line tool to extract review changes from a docx file as plain text using HTML tags <ins> and <del>.

positional arguments:
  docx        input docx

options:
  -h, --help  show this help message and exit
  --version   show version
```

Example:

```txt
$ docxreviews2txt tests/lorem_ipsum.docx
txt reviews at file:///home/alan/src/docxreviews2txt/tests/lorem_ipsum_review.txt
```

```txt
$ cat /home/alan/src/docxreviews2txt/tests/lorem_ipsum_review.txt
Typos suggestions using HTML tags <ins> and <del>:
- dolor sit amet, consectetur <ins>Lorem ipsum</ins><del>adipiscing</del>
- sit amet, consectetur adipiscing<ins>s</ins> elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim <ins>do</ins>
- Ut enim ad minim <ins>Lorem</ins>veniam<ins>ipsum</ins>
- dolor sit amet, consectetur <del>adipiscing</del>
```

## Known issues

The tool fails to capture changes in Docx files with text organized in tables (e.g., pdf2docx converts columns to tables).

## References

This project takes inspiration from:

- <https://github.com/ankushshah89/python-docx2txt>
- <https://stackoverflow.com/questions/47390928/extract-docx-comments>
- <https://stackoverflow.com/questions/38247251/how-to-extract-text-inserted-with-track-changes-in-python-docx>
