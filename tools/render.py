#!/usr/bin/env python3
"""Render a deck to one or more formats.

Quarto cannot render two LaTeX based formats of one document in a single run
(both use the same intermediate ``.tex`` file), and all PDF formats would be
written to the same ``<deck>.pdf``. This script renders one format per run and
renames the result to ``<deck>-<brand>.<ext>``.

    tools/render.py slides/example/deck.qmd
    tools/render.py slides/example/deck.qmd --to iis-pptx fau-beamer
    tools/render.py slides/example/deck.qmd --to iis-beamer -- -M lang:de
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

import yaml

PROJECT = Path(__file__).resolve().parent.parent
EXTENSIONS = {"beamer": "pdf", "pptx": "pptx"}
ALL_FORMATS = ["fau-beamer", "iis-beamer", "iis-pptx"]


def deck_formats(deck: Path) -> list[str]:
    """Formats listed in the YAML header of the deck."""
    match = re.match(r"---\n(.*?)\n---", deck.read_text(encoding="utf-8"), re.S)
    header = yaml.safe_load(match.group(1)) if match else {}
    formats = (header or {}).get("format") or {}
    if isinstance(formats, str):
        formats = {formats: None}
    return [name for name in formats if name in ALL_FORMATS] or ALL_FORMATS


def render(deck: Path, name: str, extra: list[str]) -> Path | None:
    brand, base = name.split("-", 1)
    extension = EXTENSIONS[base]
    produced = deck.with_suffix(f".{extension}")
    target = deck.with_name(f"{deck.stem}-{brand}.{extension}")
    print(f"render: {deck.name} -> {name}", file=sys.stderr)
    run = subprocess.run(["quarto", "render", str(deck), "--to", name, *extra], cwd=PROJECT)
    if run.returncode != 0 or not produced.is_file():
        print(f"render: {name} failed", file=sys.stderr)
        return None
    produced.replace(target)
    return target


def main() -> int:
    # everything after "--" goes to quarto unchanged
    argv = sys.argv[1:]
    extra: list[str] = []
    if "--" in argv:
        split = argv.index("--")
        argv, extra = argv[:split], argv[split + 1 :]

    parser = argparse.ArgumentParser(
        description=__doc__.split("\n")[0],
        usage="%(prog)s DECK [--to FORMAT ...] [-- QUARTO ARGUMENTS]",
    )
    parser.add_argument("deck", type=Path, help="path of the .qmd file")
    parser.add_argument(
        "--to",
        nargs="+",
        choices=ALL_FORMATS,
        metavar="FORMAT",
        help=f"formats to render, default: those of the deck ({', '.join(ALL_FORMATS)})",
    )
    args = parser.parse_args(argv)

    deck = args.deck.resolve()
    if not deck.is_file():
        parser.error(f"{args.deck} not found")
    results = {name: render(deck, name, extra) for name in args.to or deck_formats(deck)}
    for name, target in results.items():
        shown = target.relative_to(PROJECT) if target else "failed"
        print(f"{name:12} {shown}")
    return 0 if all(results.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
