import argparse
import logging
import sys
from pathlib import Path

import docx
from docx.opc.exceptions import PackageNotFoundError

from main import ReportGenerator


logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Автоформатирование отчётов .docx")
    parser.add_argument("input", help="Путь к исходному .docx файлу")
    parser.add_argument("output", help="Путь к выходному .docx файлу")
    parser.add_argument("-v", "--verbose", action="store_true", help="Показывать отладочные сообщения (уровень DEBUG)")
    parser.add_argument("-q", "--quiet", action="store_true", help="Показывать только предупреждения и ошибки (уровень WARNING)")
    return parser.parse_args()

def setup_logging(args: argparse.Namespace) -> None:
    if args.verbose:
        level = logging.DEBUG
    elif args.quiet:
        level = logging.WARNING
    else:
        level = logging.INFO
    logging.basicConfig(level=level, format="%(asctime)s | %(levelname)-8s | %(message)s", datefmt="%H:%M:%S",)

def main() -> None:
    args = parse_args()
    setup_logging(args)

    if not Path(args.input).exists():
        logger.error("Файл не найден: %s", args.input)
        sys.exit(1)

    try:
        input_doc = docx.Document(args.input)
    except PackageNotFoundError:
        logger.error("Файл не является .docx или повреждён: %s", args.input)
        sys.exit(1)

    generator = ReportGenerator()
    generator.parse(input_doc, args.input)
    generator.save(args.output)
    logger.info("Сохранено: %s", args.output)


if __name__ == "__main__":
    main()