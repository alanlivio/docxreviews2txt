# docxreviews2txt

Command line tool to extract review changes and comments from a docx file as plain text.
It is particullary usefull after do review changes in pdf files at docx editor (e.g., MS Word, gdocs).

## How to install?

```bash
pip install docxreviews2txt
```

## How to use it?

```txt
usage: docxreviews2txt [-h] [--save_p_xml] [--version] docx

Extract review changes and comments from a docx file as plain text.

positional arguments:
  docx          input docx

optional arguments:
  -h, --help    show this help message and exit
  --save_p_xml  also save extracted Docx paragraphs as xml for debugging
  --version     show version
```

Example:

```txt
$ docxreviews2txt tests/lorem_ipsum.docx
txt reviews at file:///C:/Users/alan/src/docxreviews2txt/tests/lorem_ipsum_review.txt
```

```txt
$ cat c:/Users/alan/src/docxreviews2txt/tests/lorem_ipsum_review.txt
# comments
- This is a comment from docx
# Typos and rewriting suggestions
- sit amet, consectetur  -> sit amet, consectetur Lorem ipsum
- sit amet, consectetur adipiscing elit, sed do -> sit amet, consectetur elit, sed do
- sit amet, consectetur adipiscing elit, sed -> sit amet, consectetur adipiscings elit, sed
- enim ad minim veniam, quis nostrud -> enim ad minim do veniam, quis nostrud
- enim ad minim veniam -> enim ad minim Lorem veniam
- veniam, quis nostrud -> veniam ipsum, quis nostrud
- sit amet, consectetur adipiscing elit, sed do -> sit amet, consectetur elit, sed do
```

## TODO

- [ ] improve N words extractions for reviews changes and enable pass it as a param
- [ ] organized extracted reviews by the input Docx headings
- [ ] save txt as Docx to enable editing
- [ ] support drag-and-drop GUI

## Known issues

The tool fails to parse Docx files with text organized in tables (e.g., pdf2docx converts columns to tables).

## Thanks

This tool takes inspiration from:

- https://github.com/ankushshah89/python-docx2txt
- https://stackoverflow.com/questions/47390928/extract-docx-comments
- https://stackoverflow.com/questions/38247251/how-to-extract-text-inserted-with-track-changes-in-python-docx