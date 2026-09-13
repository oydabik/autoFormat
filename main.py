import docx
from dataclasses import dataclass
from enum import Enum, auto
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import zipfile, io


@dataclass
class ParagraphStyle:
    """Параметры форматирования одного абзаца."""
    alignment: WD_ALIGN_PARAGRAPH
    font_name: str
    font_size: int
    line_spacing: float
    space_before: float = 0
    space_after: float = 0
    left_indent: float = 0
    first_line_indent: float = 0


class BlockKind(Enum):
    TITLE = auto()
    SECTION = auto()
    SUBHEADER = auto()
    BODY = auto()
    CODE = auto()
    LIST_ITEM = auto()
    IMAGE = auto()


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


class ReportGenerator:
    def __init__(self):
        self.doc = docx.Document()

    def add_styled_paragraph(self, text, kind):
        """Добавить абзац с заданным стилем."""
        style = ReportConfig.STYLES[kind]
        p = self.doc.add_paragraph()
        p.alignment = style.alignment
        p.paragraph_format.space_before = Pt(style.space_before)
        p.paragraph_format.space_after = Pt(style.space_after)
        p.paragraph_format.line_spacing = style.line_spacing
        p.paragraph_format.left_indent = Cm(style.left_indent)
        p.paragraph_format.first_line_indent = Cm(style.first_line_indent)

        run = p.add_run(text)
        run.font.name = style.font_name
        run.font.size = Pt(style.font_size)

    def parse(self, input_doc, docx_path):
        is_header = True
        current_mode = "main"

        extracted_images = []

        try:
            with zipfile.ZipFile(docx_path, 'r') as archive:
                media_files = [f for f in archive.namelist() if f.startswith('word/media/')]
                media_files.sort()

                for file_name in media_files:
                    img_data = archive.read(file_name)
                    extracted_images.append(io.BytesIO(img_data))
        except Exception:
            print("cant extract media")

        for paragraph in input_doc.paragraphs:
            text = paragraph.text

            if 'w:drawing' in paragraph._p.xml and extracted_images:
                img_stream = extracted_images.pop(0)
                p_img = self.doc.add_paragraph()
                # p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_stream, width=Cm(16.5))
                continue
                
            if not text.strip():
                self.doc.add_paragraph()
                continue

            if "Задача" in text or "Выводы" in text:
                is_header = False
                current_mode = "main"
                self.add_styled_paragraph(paragraph.text, BlockKind.BODY)
                continue

            if "Исходный код" in text:
                is_header = False
                current_mode = "code"
                self.add_styled_paragraph(paragraph.text, BlockKind.BODY)
                continue

            if is_header:
                self.add_styled_paragraph(paragraph.text, BlockKind.TITLE)
            elif current_mode == "main":
                self.add_styled_paragraph(paragraph.text, BlockKind.BODY)
            elif current_mode == "code":
                self.add_styled_paragraph(paragraph.text, BlockKind.CODE)
            

    def save(self, file_path):
        self.doc.save(file_path)


if __name__ == "__main__":
    generator = ReportGenerator()

    file_name = "Практика01.docx"

    try:
        input_doc = docx.Document(file_name)

        generator.parse(input_doc, file_name)
        generator.save("FORMATTED.docx")
        print("Saved!")

    except FileNotFoundError:
        print("Error!")