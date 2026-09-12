import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


class ReportConfig:
    FONT_MAIN_STYLE = "Times New Roman"
    FONT_CODE_STYLE = "Courier New"
    FONT_MAIN_SIZE = 14
    FONT_CODE_SIZE = 12
    COLOR_BLACK = RGBColor(0, 0, 0)


class ReportGenerator:
    def __init__(self):
        self.doc = docx.Document()

    def header_text(self, text):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.5

        run = p.add_run(text)
        run.font.name = ReportConfig.FONT_MAIN_STYLE
        run.font.size = Pt(ReportConfig.FONT_MAIN_SIZE)

    def main_text(self, text):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.5
              
        run = p.add_run(text)
        run.font.name = ReportConfig.FONT_MAIN_STYLE
        run.font.size = Pt(ReportConfig.FONT_MAIN_SIZE)

    def code_text(self,text):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
              
        run = p.add_run(text)
        run.font.name = ReportConfig.FONT_CODE_STYLE
        run.font.size = Pt(ReportConfig.FONT_CODE_SIZE)

    def parse(self, raw_text):
        pass

    def save(self, file_path):
        self.doc.save(file_path)

if __name__ == "__main__":
    generator = ReportGenerator()
    generator.header_text("Отчет")
    generator.main_text("Задача 1. Изучить основы")
    generator.code_text("print('Hello world')")

    generator.save("text.docx")

    read_doc = docx.Document("text.docx")
    print("Прочитанные абзацы из файла:")
    for p in read_doc.paragraphs:
        print(f"- {p.text}")