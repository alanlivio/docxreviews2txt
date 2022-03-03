from docx import Document
import argparse
import xml.etree.ElementTree as ET
import os
import pathlib
import zipfile

ET_WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ET_TEXT = ET_WORD_NS + 't'
ET_DEL = ET_WORD_NS + 'del'
ET_INS = ET_WORD_NS + 'ins'
MAX_LEN_LEFT = 4
MAX_LEN_RIGHT = 3
NS_MAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def str_deltext_elms(child):
    x = [text.text for text in child.findall(
        './/w:delText', NS_MAP)]
    return "".join(x)


def str_t_elms(child):
    x = [text.text for text in child.findall(
        './/w:t', NS_MAP)]
    return "".join(x)


def str_left_t_elms(root_p, index):
    left_ar = []
    for i in range(index, 0, -1):
        if (root_p[i].tag == ET_INS):
            continue
        left_ar = str_t_elms(root_p[i]).split(" ") + left_ar
        if(len(left_ar) >= MAX_LEN_LEFT):
            left_ar = left_ar[-MAX_LEN_LEFT:]
        break
    return " ".join(left_ar)


def str_right_t_elms(root_p, index):
    right_ar = []
    for i in range(index, len(root_p)):
        if (root_p[i].tag == ET_INS):
            continue
        right_ar = str_t_elms(root_p[i]).split(" ") + right_ar
        if(len(right_ar) >= MAX_LEN_RIGHT):
            right_ar = right_ar[:MAX_LEN_RIGHT]
        break
    return " ".join(right_ar)


def docx_p_elems_save_xml(file_docx):
    file_txt_name = os.path.splitext(file_docx)[0]+'.xml'
    doc = Document(file_docx)
    with open(file_txt_name, "w") as file:
        for p in doc.paragraphs:
            xml = p._p.xml
            file.write(f"{xml}\n")


def reviews_append(reviews, text, verbose):
    if not len(text):
        return
    reviews.append(text)
    if verbose:
        print(reviews[-1])


def docx_reviews_to_txt(file_docx, verbose, save_to_file):

    reviews = []

    # comments
    docxZip = zipfile.ZipFile(file_docx, "r")
    if "'word/comments.xml'" in [member.filename for member in docxZip.infolist()]:
        commentsXML = docxZip.read('word/comments.xml')
        root = ET.fromstring(commentsXML)
        comments = root.findall('.//w:comment', NS_MAP)
        if len(comments):
            reviews_append(reviews, "# Comments ", verbose)
            for c in comments:
                reviews_append(reviews, str_t_elms(c), verbose)

    # changes
    reviews_append(reviews, "\n" if len(reviews) else "", verbose)
    reviews_append(reviews, "# Typos and rewriting suggestions ", verbose)
    doc = Document(file_docx)
    for p in doc.paragraphs:
        xml = p._p.xml
        root = ET.fromstring(xml)
        if len(root.findall('.//w:del', NS_MAP)) == 0 and len(root.findall('.//w:ins', NS_MAP)) == 0:
            continue
        for index in range(len(root)-1):
            prev = root[index-1]
            cur = root[index]
            next = root[index+1]
            # DEL followed by INS
            if (cur.tag == ET_DEL and next.tag == ET_INS):
                del_text = str_deltext_elms(cur)
                ins_text = str_t_elms(next)
                left_text = str_left_t_elms(root, index-1)
                right_text = str_right_t_elms(root, index+2)
                reviews_append(reviews,
                               "* " + left_text + del_text + right_text +
                               " -> " + left_text + ins_text + right_text, verbose)
            # INS alone
            elif (cur.tag == ET_INS and prev.tag != ET_DEL):
                ins_text = str_t_elms(cur)
                left_text = str_left_t_elms(root, index-1)
                right_text = str_right_t_elms(root, index)
                reviews_append(reviews,
                               "* " + left_text + right_text +
                               " -> " + left_text + ins_text + right_text, verbose)
            # DEL alone
            elif (cur.tag == ET_DEL and next.tag != ET_INS):
                del_text = str_deltext_elms(cur)
                left_text = str_left_t_elms(root, index-1)
                right_text = str_right_t_elms(root, index+1)
                reviews_append(reviews,
                               "* " + left_text + del_text + right_text +
                               " -> " + left_text + right_text, verbose)

    # save_to_file
    if save_to_file:
        file_txt_name = os.path.splitext(file_docx)[0]+'_review.txt'
        with open(file_txt_name, "w") as file:
            for change in reviews:
                file.write(f"{change}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "docx", help="input docx", type=pathlib.Path)
    parser.add_argument(
        '--save', help='save review as txt', action="store_true")
    parser.add_argument(
        '--extract_xml', help='save extracted docx as xml', action="store_true")
    args = parser.parse_args()
    if args.extract_xml:
        docx_p_elems_save_xml(args.docx)
    else:
        docx_reviews_to_txt(args.docx, not args.save, args.save)
