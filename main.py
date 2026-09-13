import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import copy
import zipfile, io


class ReportConfig:
    FONT_MAIN_STYLE = "Times New Roman"
    FONT_CODE_STYLE = "Courier New"
    FONT_MAIN_SIZE = 14
    FONT_CODE_SIZE = 12


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

    file_name = "Практика01.docx"

    try:
        input_doc = docx.Document(file_name)

        generator.parse(input_doc, file_name)
        generator.save("FORMATTED.docx")
        print("Saved!")

    except FileNotFoundError:
        print("Error!")