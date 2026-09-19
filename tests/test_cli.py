import sys
from unittest.mock import patch

from cli import parse_args


def test_parse_args_positional():
    """Позиционные input и output"""
    with patch.object(sys, "argv", ["cli.py", "in.docx", "out.docx"]):
        args = parse_args()
    assert args.input == "in.docx"
    assert args.output == "out.docx"
    assert args.verbose is False
    assert args.quiet is False

def test_parse_args_v_true():
    """Verbose (-v) is True"""
    with patch.object(sys, "argv", ["cli.py", "in.docx", "out.docx", "-v"]):
        args = parse_args()
    assert args.input == "in.docx"
    assert args.output == "out.docx"
    assert args.verbose is True
    assert args.quiet is False

def test_parse_args_verbose_true():
    """Verbose (--verbose) is True"""
    with patch.object(sys, "argv", ["cli.py", "in.docx", "out.docx", "--verbose"]):
        args = parse_args()
    assert args.input == "in.docx"
    assert args.output == "out.docx"
    assert args.verbose is True
    assert args.quiet is False

def test_parse_args_q_true():
    """Quiet(-q) is True"""
    with patch.object(sys, "argv", ["cli.py", "in.docx", "out.docx", "-q"]):
        args = parse_args()
    assert args.input == "in.docx"
    assert args.output == "out.docx"
    assert args.verbose is False
    assert args.quiet is True

def test_parse_args_quiet_true():
    """Quiet (--quiet) is True"""
    with patch.object(sys, "argv", ["cli.py", "in.docx", "out.docx", "--quiet"]):
        args = parse_args()
    assert args.input == "in.docx"
    assert args.output == "out.docx"
    assert args.verbose is False
    assert args.quiet is True

def test_parse_args_all():
    """All is True"""
    with patch.object(sys, "argv", ["cli.py", "in.docx", "out.docx", "-v", "-q"]):
        args = parse_args()
    assert args.input == "in.docx"
    assert args.output == "out.docx"
    assert args.verbose is True
    assert args.quiet is True