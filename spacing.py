from models import Block, BlockKind


class SpacingRules:
    """Правила пустых строк между блоками."""

    @staticmethod
    def need_empty_after(prev_block: Block, next_block: Block | None) -> bool:
        if next_block is None:
            return False

        if prev_block.kind == BlockKind.IMAGE and next_block.kind != BlockKind.IMAGE:
            return True

        if prev_block.kind == BlockKind.CODE and next_block.kind != BlockKind.CODE:
            return True

        if prev_block.kind == BlockKind.LIST_ITEM and next_block.kind != BlockKind.LIST_ITEM:
            return True

        return False