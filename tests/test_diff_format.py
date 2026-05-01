import pytest
from lxml import etree
from docxreviews2txt.docxreviews2txt import ChangeDetector, NS_MAP
from docx.oxml.ns import qn

def test_single_insertion_diff():
    detector = ChangeDetector(n_words_around=2) # default format is 'diff'
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Hello </w:t></w:r>
        <w:ins><w:r><w:t>beautiful </w:t></w:r></w:ins>
        <w:r><w:t>world</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["Hello world -> Hello beautiful world"]

def test_single_deletion_diff():
    detector = ChangeDetector(n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Hello </w:t></w:r>
        <w:del><w:r><w:delText>old </w:delText></w:r></w:del>
        <w:r><w:t>world</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["Hello old world -> Hello world"]

def test_mixed_change_diff():
    detector = ChangeDetector(n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>The </w:t></w:r>
        <w:del><w:r><w:delText>quick </w:delText></w:r></w:del>
        <w:ins><w:r><w:t>fast </w:t></w:r></w:ins>
        <w:r><w:t>fox</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["The quick fox -> The fast fox"]

def test_multiple_changes_diff():
    detector = ChangeDetector(n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>A </w:t></w:r>
        <w:del><w:r><w:delText>small </w:delText></w:r></w:del>
        <w:ins><w:r><w:t>big </w:t></w:r></w:ins>
        <w:r><w:t>red </w:t></w:r>
        <w:del><w:r><w:delText>cat</w:delText></w:r></w:del>
        <w:ins><w:r><w:t>dog</w:t></w:r></w:ins>
        <w:r><w:t> jumps.</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    # red is 1 word, so it should group.
    assert changes == ["A small red cat jumps. -> A big red dog jumps."]
