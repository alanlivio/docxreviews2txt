import pytest
from lxml import etree
from docxreviews2txt.docxreviews2txt import ChangeDetector, NS_MAP
from docx.oxml.ns import qn

def test_single_insertion():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Hello </w:t></w:r>
        <w:ins><w:r><w:t>beautiful </w:t></w:r></w:ins>
        <w:r><w:t>world</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["Hello <ins>beautiful </ins>world"]

def test_single_deletion():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Hello </w:t></w:r>
        <w:del><w:r><w:delText>old </w:delText></w:r></w:del>
        <w:r><w:t>world</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["Hello <del>old </del>world"]

def test_consecutive_changes():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Start </w:t></w:r>
        <w:ins><w:r><w:t>ins1 </w:t></w:r></w:ins>
        <w:del><w:r><w:delText>del1 </w:delText></w:r></w:del>
        <w:r><w:t>End</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["Start <ins>ins1 </ins><del>del1 </del>End"]

def test_near_changes_grouping():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Very long sentence before. </w:t></w:r>
        <w:r><w:t>Context </w:t></w:r>
        <w:ins><w:r><w:t>ins1 </w:t></w:r></w:ins>
        <w:r><w:t>mid </w:t></w:r>
        <w:del><w:r><w:delText>del1 </w:delText></w:r></w:del>
        <w:r><w:t>Context </w:t></w:r>
        <w:r><w:t>Very long sentence after.</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    # n_words_around=2, 'mid ' is 1 word, so it should group.
    assert changes == ["before. Context <ins>ins1 </ins>mid <del>del1 </del>Context Very"]

def test_far_changes_no_grouping():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Start </w:t></w:r>
        <w:ins><w:r><w:t>ins1 </w:t></w:r></w:ins>
        <w:r><w:t>word1 word2 word3 word4 </w:t></w:r>
        <w:del><w:r><w:delText>del1 </w:delText></w:r></w:del>
        <w:r><w:t>End</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert len(changes) == 2
    assert changes[0] == "Start <ins>ins1 </ins>word1 word2"
    assert changes[1] == "word3 word4 <del>del1 </del>End"

def test_formatting_change_run():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Normal </w:t></w:r>
        <w:r>
            <w:rPr><w:rPrChange w:id="0" w:author="User" w:date="2024-01-01T00:00:00Z"/></w:rPr>
            <w:t>Formatted</w:t>
        </w:r>
        <w:r><w:t> text</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["Normal [fmt]Formatted text"]

def test_formatting_change_para():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:pPr>
            <w:pPrChange w:id="0" w:author="User" w:date="2024-01-01T00:00:00Z"/>
        </w:pPr>
        <w:r><w:t>This paragraph changed formatting.</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["[para fmt]This paragraph"]

def test_change_at_start():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:ins><w:r><w:t>Added </w:t></w:r></w:ins>
        <w:r><w:t>at the beginning.</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["<ins>Added </ins>at the"]

def test_change_at_end():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Something at the </w:t></w:r>
        <w:del><w:r><w:delText>end.</w:delText></w:r></w:del>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["at the <del>end.</del>"]

def test_very_short_para():
    detector = ChangeDetector(output_format='tags', n_words_around=5)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Hi </w:t></w:r>
        <w:ins><w:r><w:t>there </w:t></w:r></w:ins>
        <w:r><w:t>!</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    assert changes == ["Hi <ins>there </ins>!"]

def test_empty_ins():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Before </w:t></w:r>
        <w:ins><w:r><w:t></w:t></w:r></w:ins>
        <w:r><w:t>After</w:t></w:r>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    # Should not produce a change if insertion is empty
    assert changes == []

def test_multiple_changes_in_para():
    detector = ChangeDetector(output_format='tags', n_words_around=2)
    xml = f"""
    <w:p xmlns:w="{NS_MAP['w']}">
        <w:r><w:t>Word1 </w:t></w:r>
        <w:ins><w:r><w:t>ins1 </w:t></w:r></w:ins>
        <w:r><w:t>Word2 </w:t></w:r>
        <w:del><w:r><w:delText>del1 </w:delText></w:r></w:del>
        <w:r><w:t>Word3 </w:t></w:r>
        <w:ins><w:r><w:t>ins2</w:t></w:r></w:ins>
    </w:p>
    """
    p_element = etree.fromstring(xml)
    changes = detector.process_paragraph(p_element)
    # All should be grouped because Word2 is 1 word, Word3 is 1 word.
    assert changes == ["Word1 <ins>ins1 </ins>Word2 <del>del1 </del>Word3 <ins>ins2</ins>"]
