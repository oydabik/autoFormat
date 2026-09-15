from dataclasses import dataclass
from enum import Enum, auto

from docx.enum.text import WD_ALIGN_PARAGRAPH


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