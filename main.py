import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class ReportConfig:
    FONT_MAIN_STYLE = "Times New Roman"
    FONT_CODE_STYLE = "Courier New"
    FONT_MAIN_SIZE = 14
    FONT_CODE_SIZE = 12
    COLOR_BLACK = RGBColor(0, 0, 0)


class ReportGenerator:
    def __init__(self):
        self.doc = docx.Document()

    def header_text(self, src_paragraph):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.5

        run = p.add_run(src_paragraph.text)

        run.font.name = ReportConfig.FONT_MAIN_STYLE
        run.font.size = Pt(ReportConfig.FONT_MAIN_SIZE)

    def main_text(self, src_paragraph):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.5

        run = p.add_run(src_paragraph.text)
              
        run.font.name = ReportConfig.FONT_MAIN_STYLE
        run.font.size = Pt(ReportConfig.FONT_MAIN_SIZE)

    def code_text(self, src_paragraph):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0

        run = p.add_run(src_paragraph.text)
              
        run.font.name = ReportConfig.FONT_CODE_STYLE
        run.font.size = Pt(ReportConfig.FONT_CODE_SIZE)

    def parse(self, input_doc):
        is_header = True
        current_mode = "main"

        for paragraph in input_doc.paragraphs:
            text = paragraph.text

            if not text.strip():
                self.doc.add_paragraph()
                continue

            if "Задача" in text or "Выводы" in text:
                is_header = False
                current_mode = "main"
                self.main_text(paragraph)
                continue

            if "Исходный код" in text:
                is_header = False
                current_mode = "code"
                self.main_text(paragraph)
                continue

            if is_header:
                self.header_text(paragraph)
            elif current_mode == "main":
                self.main_text(paragraph)
            elif current_mode == "code":
                self.code_text(paragraph)
            

    def save(self, file_path):
        self.doc.save(file_path)

if __name__ == "__main__":
    generator = ReportGenerator()

    try:
        input_doc = docx.Document("Практика01.docx")

        generator.parse(input_doc)
        generator.save("FORMATTED.docx")
        print("Saved!")

    except FileNotFoundError:
        print("Error!")