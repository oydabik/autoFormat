from generator import ReportGenerator
from models import BlockKind


def test_read_blocks_returns_list(simple_doc):
    """Возвращает список."""
    doc, path = simple_doc
    gen = ReportGenerator()
    blocks = gen.read_blocks(doc, path)
    assert isinstance(blocks, list)


def test_read_blocks_count(simple_doc):
    """Находит все 5 абзацев."""
    doc, path = simple_doc
    gen = ReportGenerator()
    blocks = gen.read_blocks(doc, path)
    assert len(blocks) == 5


def test_read_blocks_first_is_title(simple_doc):
    """Первый блок - TITLE."""
    doc, path = simple_doc
    gen = ReportGenerator()
    blocks = gen.read_blocks(doc, path)
    assert blocks[0].kind is BlockKind.TITLE


def test_read_blocks_has_section(simple_doc):
    """Среди блоков есть SECTION"""
    doc, path = simple_doc
    gen = ReportGenerator()
    blocks = gen.read_blocks(doc, path)
    kinds = [b.kind for b in blocks]
    assert BlockKind.SECTION in kinds