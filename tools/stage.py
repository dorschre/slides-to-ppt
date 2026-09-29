#!/usr/bin/env python3
"""Pre-render hook: symlink brand assets into every deck directory.

The beamer themes resolve their files relative to the LaTeX working directory,
which for Quarto is the directory of the rendered document. Nothing inside
``assets/`` is modified.

The Fraunhofer beamer slides follow the PowerPoint template. Its logo and
background pictures are copied from the template into ``iis-pptx/``.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
IIS = PROJECT / "assets" / "fraunhofer" / "beamer theme IIS V1_05_2026"
FAU = PROJECT / "assets" / "fau" / "fau-beamer"

# link name in the deck directory -> target directory
LINKS = {
    "theme": IIS / "theme",
    "logos": IIS / "logos",
    "iis-images": IIS / "images",
    "styles": FAU / "styles",
    "template-art": FAU / "template-art",
}


def deck_dirs() -> list[Path]:
    inputs = os.environ.get("QUARTO_PROJECT_INPUT_FILES", "").split("\n")
    dirs = {(PROJECT / f).resolve().parent for f in inputs if f.strip()}
    if not dirs:
        dirs = {p.parent for p in (PROJECT / "slides").rglob("*.qmd")}
    return sorted(dirs)


def stage(deck: Path) -> None:
    for name, target in LINKS.items():
        link = deck / name
        if not target.is_dir():
            print(f"stage: missing asset directory {target}", file=sys.stderr)
            continue
        rel = os.path.relpath(target, deck)
        if link.is_symlink():
            if os.readlink(link) == rel:
                continue
            link.unlink()
        elif link.exists():
            print(f"stage: {link} exists and is not a symlink, skipped", file=sys.stderr)
            continue
        link.symlink_to(rel, target_is_directory=True)
        print(f"stage: {link.relative_to(PROJECT)} -> {rel}", file=sys.stderr)


# file name in iis-pptx/ -> (template language, layout number, shape name)
PPTX_DIR = PROJECT / "assets" / "fraunhofer"
PPTX_TEMPLATES = {
    "de": "Präsentation-Fraunhofer-IIS-Vorlage-mit-Font-DE.pptx",
    "en": "Präsentation-Fraunhofer-IIS-Vorlage-mit-Font-EN.pptx",
}
PPTX_MEDIA = {
    "logo-de.png": ("de", 1, "iis_rgb_modul_send_de"),
    "logo-en.png": ("en", 1, "iis_rgb_modul_send_en"),
    "chapter.jpg": ("en", 12, "Grafik 1"),
    "final.jpg": ("en", 29, "Grafik 1"),
}


def extract_media(target: Path) -> None:
    """Copies pictures of the PowerPoint layouts, once per template version."""
    newest = max((PPTX_DIR / n).stat().st_mtime for n in PPTX_TEMPLATES.values() if (PPTX_DIR / n).is_file())
    wanted = [target / name for name in PPTX_MEDIA]
    if all(f.is_file() and f.stat().st_mtime >= newest for f in wanted):
        return
    from pptx import Presentation

    namespace = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    target.mkdir(exist_ok=True)
    decks = {}
    for name, (lang, number, shape_name) in PPTX_MEDIA.items():
        if lang not in decks:
            decks[lang] = Presentation(str(PPTX_DIR / PPTX_TEMPLATES[lang]))
        layout = decks[lang].slide_layouts[number - 1]
        pictures = [s for s in layout.shapes if s.shape_type is not None and s._element.tag.endswith("}pic")]
        match = [s for s in pictures if s.name == shape_name] or pictures
        if not match:
            print(f"stage: no picture {shape_name!r} in layout {number}", file=sys.stderr)
            continue
        blip = next(e for e in match[0]._element.iter() if e.tag.endswith("}blip"))
        part = layout.part.related_part(blip.get(namespace))
        (target / name).write_bytes(part.blob)
        print(f"stage: {(target / name).relative_to(PROJECT)} <- {part.partname}", file=sys.stderr)


def main() -> None:
    for deck in deck_dirs():
        stage(deck)
        try:
            extract_media(deck / "iis-pptx")
        except Exception as error:
            print(f"stage: PowerPoint pictures not extracted: {error!r}", file=sys.stderr)


if __name__ == "__main__":
    main()
