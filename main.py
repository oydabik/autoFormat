import docx
from dataclasses import dataclass
from enum import Enum, auto
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
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


@dataclass
class Block:
    """Один блок документа (абзац, код, картинка)."""
    kind: BlockKind
    text: str = ""
    image: bytes | None = None


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

    def read_blocks(self, input_doc: docx.Document, docx_path: str) -> list[Block]:
        blocks = []
        is_header = True
        current_mode = "main"
        images = self._extract_images(docx_path)
        image_index = 0
        for paragraph in input_doc.paragraphs:
            text = paragraph.text
            if 'w:drawing' in paragraph._p.xml:
                if image_index < len(images):
                    blocks.append(Block(BlockKind.IMAGE, image=images[image_index]))
                    image_index += 1
                continue
                
            if not text.strip():
                continue

            if "Задача" in text or "Выводы" in text:
                is_header = False
                current_mode = "main"
                blocks.append(Block(BlockKind.BODY, text=text))
                continue

            if "Исходный код" in text:
                is_header = False
                current_mode = "code"
                blocks.append(Block(BlockKind.BODY, text=text))
                continue

            kind = self._classify(text, is_header, current_mode)
            blocks.append(Block(kind, text=text))

        return blocks

    def render_blocks(self, blocks: list[Block]) -> None:
        """Отрендерить список блоков в документ."""
        for block in blocks:
            if block.kind == BlockKind.IMAGE:
                img_stream = io.BytesIO(block.image)
                p_img = self.doc.add_paragraph()
                p_img.add_run().add_picture(img_stream, width=Cm(16.5))
            else:
                self.add_styled_paragraph(block.text, block.kind)

    def _extract_images(self, docx_path: str) -> list[bytes]:
        extracted_images = []
        try:
            with zipfile.ZipFile(docx_path, 'r') as archive:
                media_files = [f for f in archive.namelist() if f.startswith('word/media/')]
                media_files.sort()
                for file_name in media_files:
                    img_data = archive.read(file_name)
                    extracted_images.append(img_data)
        except Exception:
            print("cant extract media")
        return extracted_images

    def _classify(self, text: str, is_header: bool, current_mode: str) -> BlockKind:
        if is_header:
            return BlockKind.TITLE
        if current_mode == "code":
            return BlockKind.CODE
        return BlockKind.BODY

    def add_styled_paragraph(self, text: str, kind:BlockKind) -> None:
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
        rFonts = run._element.get_or_add_rPr().get_or_add_rFonts()
        rFonts.set(qn('w:eastAsia'), style.font_name)
        run.font.size = Pt(style.font_size)

    def parse(self, input_doc: docx.Document, docx_path: str) -> None:
        blocks = self.read_blocks(input_doc, docx_path)
        self.render_blocks(blocks)

    def save(self, file_path: str) -> None:
        self.doc.save(file_path)


if __name__ == "__main__":
    generator = ReportGenerator()
    file_name = "Практика01.docx"

    try:
        input_doc = docx.Document(file_name)
        blocks = generator.read_blocks(input_doc, file_name)
        for b in blocks:
            preview = b.text[:40] if b.text else f"<image {len(b.image) if b.image else 0} bytes>"
            print(f"{b.kind.name:12} | {preview}")
        generator.parse(input_doc, file_name)
        generator.save("FORMATTED.docx")
        print("Saved!")

    except FileNotFoundError:
        print("Error!")