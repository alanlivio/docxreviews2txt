# docxreviews2txt

Command line tool to extract review changes from a docx file as plain text. It is useful when reviewing a PDF file as docx, and you need to share the changes as plain text.

## How to install?

```bash
pip install docxreviews2txt
```

## How to use it?

```txt
usage: docxreviews2txt [-h] [--format {diff,tags}] [--version] docx

Command line tool to extract review changes from a docx file as plain text.

positional arguments:
  docx                  input docx

options:
  -h, --help            show this help message and exit
  --format {diff,tags}  output format: 'diff' (PREVIOUS -> AFTER) or 'tags' (<ins>/<del>). Default is 'diff'.
  --version             show version
```

Example (Default format: `diff`):

```txt
$ docxreviews2txt tests/lorem_ipsum.docx
txt reviews at file:///home/alan/src/docxreviews2txt/tests/lorem_ipsum_review.txt
```

```txt
$ cat tests/lorem_ipsum_review.txt
- dolor sit amet, consectetur adipiscing elit, sed do eiusmod -> dolor sit amet, consectetur Lorem ipsum elit, sed do eiusmod
- sit amet, consectetur adipiscing elit, sed do eiusmod -> sit amet, consectetur adipiscings elit, sed do eiusmod
```

Example (Tags format):

```txt
$ docxreviews2txt --format tags tests/lorem_ipsum.docx
txt reviews at file:///home/alan/src/docxreviews2txt/tests/lorem_ipsum_review.txt
```

```txt
$ cat tests/lorem_ipsum_review.txt
- dolor sit amet, consectetur <ins>Lorem ipsum</ins><del>adipiscing</del> elit, sed do eiusmod
- sit amet, consectetur adipiscing<ins>s</ins> elit, sed do eiusmod
```

## Known issues

The tool fails to capture changes in Docx files with text organized in tables (e.g., pdf2docx converts columns to tables).

## References

This project takes inspiration from:

- <https://github.com/ankushshah89/python-docx2txt>
- <https://stackoverflow.com/questions/47390928/extract-docx-comments>
- <https://stackoverflow.com/questions/38247251/how-to-extract-text-inserted-with-track-changes-in-python-docx>
