#!/usr/bin/env python3
"""Post-render hook: replace pandoc's generic PPTX with the branded one.

``quarto render --to iis-pptx`` writes ``<output>.ast.json`` next to the PPTX
(see ``_extensions/shared/dump-ast.lua``). Every output that has such a file
is rebuilt from the official template by ``md2pptx``.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import md2pptx  # noqa: E402


def main() -> int:
    project = Path(os.environ.get("QUARTO_PROJECT_DIR", Path(__file__).resolve().parent.parent))
    outputs = os.environ.get("QUARTO_PROJECT_OUTPUT_FILES", "").split("\n")
    failed = 0
    for name in filter(None, (o.strip() for o in outputs)):
        output = (project / name).resolve()
        ast = output.with_name(output.name + ".ast.json")
        if output.suffix != ".pptx" or not ast.is_file():
            continue
        try:
            md2pptx.build(ast, output)
        except Exception as error:  # keep pandoc's file and report
            print(f"postrender: {output.name} failed: {error!r}", file=sys.stderr)
            failed += 1
            continue
        if not os.environ.get("KEEP_AST"):
            ast.unlink()
        print(f"postrender: branded {output.relative_to(project)}", file=sys.stderr)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
