# docx_reviews_to_txt


## usage 

```txt
$ python docx_reviews_to_txt.py -h
usage: docx_reviews_to_txt.py [-h] [--save] [--dump_paragraphs_xml] docx

positional arguments:
  docx                  input docx

options:
  -h, --help            show this help message and exit
  --save                save review as txt
  --dump_paragraphs_xml
                        save extracted paragraphs xml
```
  
## example

```txt
$ python docx_reviews_to_txt.py lorem_ipsum.docx
# Typos and rewriting suggestions 
* sit amet, consectetur  -> sit amet, consectetur Lorem ipsum
* sit amet, consectetur adipiscing elit, sed do -> sit amet, consectetur elit, sed do
* sit amet, consectetur adipiscing elit, sed -> sit amet, consectetur adipiscings elit, sed
* enim ad minim veniam, quis nostrud -> enim ad minim do veniam, quis nostrud
* enim ad minim veniam -> enim ad minim Lorem veniam
* veniam, quis nostrud -> veniam ipsum, quis nostrud
* sit amet, consectetur adipiscing elit, sed do -> sit amet, consectetur elit, sed do
```
  