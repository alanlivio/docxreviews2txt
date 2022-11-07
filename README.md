# docxreviews2txt

Extract review changes and comments from a docx file as plan text.
It is useful when reviewing a pdf files imported as Word. 

## How to install? 

```bash
pip install docxreviews2txt
```

## How to use?

```txt
$ docxreviews2txt -h
usage: docxreviews2txt [-h] [--save_txt | --save_p_xml] docx

positional arguments:
  docx          input docx

options:
  -h, --help    show this help message and exit
  --save_txt    save review as txt
  --save_p_xml  save extracted paragraphs xml for debugging
```
  
Example:

```txt
$ docxreviews2txt tests/lorem_ipsum.docx
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

## Known issues

The tool fails in parse docx files with text organized in tables (e.g., pdf2docx convert collums to tables). 

## Thanks

This tool take inspiration from:

- https://github.com/ankushshah89/python-docx2txt
- https://stackoverflow.com/questions/47390928/extract-docx-comments
- https://stackoverflow.com/questions/38247251/how-to-extract-text-inserted-with-track-changes-in-python-docx