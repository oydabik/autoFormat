import re

from docx.enum.text import WD_ALIGN_PARAGRAPH

from models import BlockKind, ParagraphStyle


LIST_ITEM_PATTERN = re.compile(r"^\d+\.\s")


class ReportConfig:
    _TEXT = ParagraphStyle(
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        font_name="Times New Roman",
        font_size=14,
        line_spacing=1.5,
    )
    STYLES = {
        BlockKind.TITLE: ParagraphStyle(
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            font_name="Times New Roman",
            font_size=14,
            line_spacing=1.5,
        ),
        BlockKind.SECTION: _TEXT,
        BlockKind.SUBHEADER: _TEXT,
        BlockKind.BODY: _TEXT,
        BlockKind.CODE: ParagraphStyle(
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            font_name="Courier New",
            font_size=12,
            line_spacing=1.0,
        ),
        BlockKind.LIST_ITEM: ParagraphStyle(
            alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            font_name="Times New Roman",
            font_size=14,
            line_spacing=1.5,
            left_indent=1.25,
            first_line_indent=-1.25,
        ),
    }