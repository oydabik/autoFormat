"""Тесты для SpacingRules."""
from models import Block, BlockKind
from spacing import SpacingRules


def test_empty_after_image_group():
    """После последней картинки в группе - пустая строка."""
    prev = Block(kind=BlockKind.IMAGE, image=b"")
    nxt = Block(kind=BlockKind.SUBHEADER, text="Исходный код")
    assert SpacingRules.need_empty_after(prev, nxt) is True


def test_no_empty_inside_image_group():
    """Между двумя картинками - без пустой строки."""
    prev = Block(kind=BlockKind.IMAGE, image=b"")
    nxt = Block(kind=BlockKind.IMAGE, image=b"")
    assert SpacingRules.need_empty_after(prev, nxt) is False


def test_empty_after_code_group():
    """После последнего кода в группе - пустая строка."""
    prev = Block(kind=BlockKind.CODE, text="print('hi')")
    nxt = Block(kind=BlockKind.SECTION, text="Выводы")
    assert SpacingRules.need_empty_after(prev, nxt) is True


def test_no_empty_inside_code_group():
    """Между двумя абзацами кода - без пустой строки."""
    prev = Block(kind=BlockKind.CODE, text="line 1")
    nxt = Block(kind=BlockKind.CODE, text="line 2")
    assert SpacingRules.need_empty_after(prev, nxt) is False


def test_empty_after_list_item_group():
    """После последнего пункта списка - пустая строка."""
    prev = Block(kind=BlockKind.LIST_ITEM, text="1. Первый")
    nxt = Block(kind=BlockKind.SECTION, text="Задача 2.")
    assert SpacingRules.need_empty_after(prev, nxt) is True


def test_no_empty_inside_list_item_group():
    """Между двумя пунктами списка - без пустой строки."""
    prev = Block(kind=BlockKind.LIST_ITEM, text="1. Первый")
    nxt = Block(kind=BlockKind.LIST_ITEM, text="2. Второй")
    assert SpacingRules.need_empty_after(prev, nxt) is False


def test_no_empty_after_title():
    """После титульной части - без пустой строки."""
    prev = Block(kind=BlockKind.TITLE, text="Отчёт")
    nxt = Block(kind=BlockKind.SECTION, text="Задача 1.")
    assert SpacingRules.need_empty_after(prev, nxt) is False


def test_no_empty_when_next_is_none():
    """После последнего блока - без пустой строки."""
    prev = Block(kind=BlockKind.CODE, text="last line")
    assert SpacingRules.need_empty_after(prev, None) is False


def test_no_empty_after_body():
    """После обычного текста - без пустой строки."""
    prev = Block(kind=BlockKind.BODY, text="some text")
    nxt = Block(kind=BlockKind.BODY, text="more text")
    assert SpacingRules.need_empty_after(prev, nxt) is False