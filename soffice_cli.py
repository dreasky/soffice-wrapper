#!/usr/bin/env python3
"""
Soffice CLI

Subcommands:
  convert   Convert a document to PDF
"""

import argparse
import sys
from pathlib import Path

scripts_dir = Path(__file__).parent
if str(scripts_dir) not in sys.path:
    sys.path.insert(0, str(scripts_dir))

from soffice_wrapper import SofficeWrapper


def cmd_convert(args):
    SofficeWrapper().convert(Path(args.input), Path(args.output))


def main():
    parser = argparse.ArgumentParser(prog="soffice_cli")
    sub = parser.add_subparsers(dest="command", required=True)

    p_conv = sub.add_parser("convert", help="Convert a document to PDF")
    p_conv.add_argument(
        "-i", "--input", required=True, dest="input", help="Input file path"
    )
    p_conv.add_argument(
        "-o", "--output", required=True, dest="output", help="Output directory"
    )

    args = parser.parse_args()

    try:
        cmd_convert(args)
        return 0
    except Exception as e:
        print(f"转换失败: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
