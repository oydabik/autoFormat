from generator import ReportGenerator
from models import BlockKind


def test_classify_section_from_task():
    """Текст, начинающийся с 'Задача' - SECTION."""
    gen = ReportGenerator()
    assert gen._classify("Задача 1. Текст", False, "main") is BlockKind.SECTION


def test_classify_subheader_from_screenshot():
    """Текст, начинающийся с 'Скриншоты' - SUBHEADER."""
    gen = ReportGenerator()
    assert gen._classify("Скриншоты", False, "main") is BlockKind.SUBHEADER


def test_classify_title_when_header():
    """Текст в начале - TITLE."""
    gen = ReportGenerator()
    assert gen._classify("Текст", True, "main") is BlockKind.TITLE


def test_classify_code_when_code_mode():
    """Текст с current_mode='code' - CODE."""
    gen = ReportGenerator()
    assert gen._classify("Текст", False, "code") is BlockKind.CODE


def test_classify_list_item_from_regex():
    """Текст, начинающийся с regex - LIST_ITEM."""
    gen = ReportGenerator()
    assert gen._classify("1. Текст", False, "main") is BlockKind.LIST_ITEM


def test_classify_body_from_plain_text():
    """Просто текст - BODY."""
    gen = ReportGenerator()
    assert gen._classify("Текст", False, "main") is BlockKind.BODY


def test_classify_section_overrides_header():
    """Текст, начинающийся с 'Задача 1.' при is_header=TRUE - SECTION."""
    gen = ReportGenerator()
    assert gen._classify("Задача 1. ", True, "main") is BlockKind.SECTION