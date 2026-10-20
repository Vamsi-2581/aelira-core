"""Regression tests for reading picture alternative text from PresentationML."""

from types import SimpleNamespace
from xml.sax.saxutils import quoteattr

from pptx.oxml import parse_xml

from src.education.pptx_processor import PowerPointProcessor


_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _picture_shape(descr=None):
    descr_attribute = f" descr={quoteattr(descr)}" if descr is not None else ""
    element = parse_xml(
        f'<p:pic xmlns:p="{_P_NS}" xmlns:a="{_A_NS}">'
        "<p:nvPicPr>"
        f'<p:cNvPr id="2" name="Picture 1"{descr_attribute}/>'
        "</p:nvPicPr>"
        "</p:pic>"
    )
    return SimpleNamespace(_element=element, shape_id=2, name="Picture 1")


def _processor():
    processor = PowerPointProcessor.__new__(PowerPointProcessor)
    processor.generate_alt_text = False
    processor.image_generator = None
    processor.validate_alt_text = False
    return processor


def test_check_alt_text_reads_description_from_picture_nonvisual_properties():
    processor = _processor()
    shape = _picture_shape("A labelled bar chart")

    assert processor._check_alt_text(1, shape) is None


def test_check_alt_text_reports_picture_without_description():
    processor = _processor()
    shape = _picture_shape()

    issue = processor._check_alt_text(1, shape)

    assert issue is not None
    assert issue.has_alt_text is False
    assert issue.existing_alt_text is None


def test_saved_file_rescan_recognizes_remediated_picture_description():
    processor = _processor()
    shape = _picture_shape()

    assert processor._check_alt_text(1, shape) is not None

    c_nv_pr = shape._element.xpath("./p:nvPicPr/p:cNvPr")[0]
    c_nv_pr.set("descr", "A labelled bar chart")

    assert processor._check_alt_text(1, shape) is None
