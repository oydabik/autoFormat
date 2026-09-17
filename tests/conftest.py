import pytest
import docx


@pytest.fixture
def simple_doc(tmp_path):
    """Простой .docx с несколькими абзацами разных типов."""
    doc = docx.Document()
    doc.add_paragraph("Практика")
    doc.add_paragraph("Задача 1. Текст")
    doc.add_paragraph("Решение.")
    doc.add_paragraph("1. Текст")
    doc.add_paragraph("Текст")

    file_path = tmp_path / "test.docx"
    doc.save(file_path)

    return doc, str(file_path)