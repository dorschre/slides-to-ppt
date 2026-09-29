#!/usr/bin/env python3
"""Build a branded PPTX from a pandoc AST.

The AST is written by ``_extensions/shared/dump-ast.lua`` during
``quarto render --to iis-pptx``. Slides are created from the layouts of the
official template; ``_extensions/iis/layouts.yaml`` maps logical layouts to
layout numbers and placeholder indices.

    md2pptx.py deck.iis.pptx.ast.json -o deck.iis.pptx
"""

from __future__ import annotations

import argparse
import base64
import copy
import datetime as dt
import io
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Cm, Emu, Pt

PROJECT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG = PROJECT / "_extensions" / "iis" / "layouts.yaml"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"a": A, "p": P, "r": R, "a14": A14, "p14": P14, "m": M, "mc": MC}

FOOTER_TYPES = {
    PP_PLACEHOLDER.DATE,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.SLIDE_NUMBER,
}
BULLET_TAGS = {f"{{{A}}}{t}" for t in ("buNone", "buChar", "buAutoNum", "buBlip")}
FILL_TAGS = {f"{{{A}}}{t}" for t in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")}
# child order required by the DrawingML schema
PARAGRAPH_ORDER = [
    "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts",
    "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr", "extLst",
]
RUN_ORDER = [
    "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst",
    "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
    "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
]
DASH = "—"
# classes for coloured text, also in _extensions/shared/layouts.lua
TEXT_COLOURS = {"ok": "179C7D", "warn": "F58220", "accent": "005B7F"}

# token type prefix -> (colour, bold, italic), first match wins
CODE_COLOURS = [
    ("Token.Comment", ("7F7F7F", False, True)),
    ("Token.Keyword", ("005B7F", True, False)),
    ("Token.Literal.String", ("179C7D", False, False)),
    ("Token.Literal.Number", ("008598", False, False)),
    ("Token.Name.Function", ("005B7F", False, False)),
    ("Token.Name.Class", ("005B7F", False, False)),
    ("Token.Name.Builtin", ("008598", False, False)),
    ("Token.Operator.Word", ("005B7F", True, False)),
]


def qn(tag: str) -> str:
    prefix, name = tag.split(":")
    return f"{{{NS[prefix]}}}{name}"


def warn(message: str) -> None:
    print(f"md2pptx: {message}", file=sys.stderr)


# --------------------------------------------------------------------------
# Document model
# --------------------------------------------------------------------------


@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    code: bool = False
    strike: bool = False
    underline: bool = False
    script: str | None = None  # "sub" or "super"
    link: str | None = None
    math: str | None = None  # TeX source
    display: bool = False
    br: bool = False
    colour: str | None = None


@dataclass
class Para:
    runs: list[Run]
    role: str = "text"  # text, subheading, heading, bullet, number
    depth: int = 0
    start: int | None = None
    align: str | None = None

    @property
    def is_display_math(self) -> bool:
        return len(self.runs) == 1 and self.runs[0].math is not None and self.runs[0].display


@dataclass
class Code:
    text: str
    lang: str | None = None


@dataclass
class Table:
    header: list[list[list[Run]]]
    rows: list[list[list[Run]]]
    aligns: list[str]
    caption: list[Run] = field(default_factory=list)


@dataclass
class Picture:
    data: bytes
    caption: list[Run] = field(default_factory=list)
    source: str = ""
    width: str | None = None  # "50%" of the area, or a length such as "5cm"
    height: str | None = None


@dataclass
class Quote:
    paras: list[Para]


@dataclass
class Columns:
    columns: list[list]


@dataclass
class Slide:
    kind: str  # content or section
    title: list[Run]
    attrs: dict[str, str]
    items: list = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def title_text(self) -> str:
        return plain(self.title)


def plain(runs: list[Run]) -> str:
    return "".join("\n" if r.br else (r.math if r.math is not None else r.text) for r in runs)


# --------------------------------------------------------------------------
# pandoc AST -> document model
# --------------------------------------------------------------------------


class Reader:
    def __init__(self, doc: dict):
        self.doc = doc
        self.meta = {k: self.meta_value(v) for k, v in doc.get("meta", {}).items()}

    # ---- metadata ----

    def meta_value(self, node):
        t, c = node["t"], node.get("c")
        if t == "MetaString":
            return c
        if t == "MetaBool":
            return c
        if t == "MetaInlines":
            return plain(self.inlines(c))
        if t == "MetaBlocks":
            return "\n".join(plain(p.runs) for p in self.blocks(c) if isinstance(p, Para))
        if t == "MetaList":
            return [self.meta_value(v) for v in c]
        if t == "MetaMap":
            return {k: self.meta_value(v) for k, v in c.items()}
        return None

    def meta_text(self, key: str, default: str = "") -> str:
        value = self.meta.get(key)
        if value is None or value is False:
            return default
        if isinstance(value, list):
            return ", ".join(str(v) for v in value if isinstance(v, str))
        if isinstance(value, dict):
            return default
        return str(value)

    # ---- inlines ----

    def inlines(self, nodes: list, style: Run | None = None) -> list[Run]:
        style = style or Run()
        out: list[Run] = []

        def styled(**changes) -> Run:
            return _replace(style, **changes)

        for node in nodes:
            t, c = node["t"], node.get("c")
            if t == "Str":
                out.append(styled(text=c))
            elif t in ("Space", "SoftBreak"):
                out.append(styled(text=" "))
            elif t == "LineBreak":
                out.append(Run(br=True))
            elif t == "Strong":
                out += self.inlines(c, styled(bold=True))
            elif t == "Emph":
                out += self.inlines(c, styled(italic=True))
            elif t == "Underline":
                out += self.inlines(c, styled(underline=True))
            elif t == "Strikeout":
                out += self.inlines(c, styled(strike=True))
            elif t == "Superscript":
                out += self.inlines(c, styled(script="super"))
            elif t == "Subscript":
                out += self.inlines(c, styled(script="sub"))
            elif t == "SmallCaps":
                out += self.inlines(c, style)
            elif t == "Code":
                out.append(styled(text=c[1], code=True))
            elif t == "Math":
                out.append(styled(math=c[1].strip(), display=c[0]["t"] == "DisplayMath"))
            elif t == "Link":
                out += self.inlines(c[1], styled(link=c[2][0]))
            elif t == "Quoted":
                marks = ("“", "”") if c[0]["t"] == "DoubleQuote" else ("‘", "’")
                out.append(styled(text=marks[0]))
                out += self.inlines(c[1], style)
                out.append(styled(text=marks[1]))
            elif t == "Span":
                # [text]{.ok}, {.warn}, {.accent}: coloured text
                colour = next((TEXT_COLOURS[k] for k in c[0][1] if k in TEXT_COLOURS), None)
                out += self.inlines(c[1], styled(colour=colour) if colour else style)
            elif t == "Cite":
                out += self.inlines(c[1], style)
            elif t == "Image":
                warn("image inside a text paragraph is not supported, skipped")
            # RawInline and Note carry nothing that can be shown on a slide
        return merge_runs(out)

    # ---- blocks ----

    def picture(self, image_node: dict, caption: list[Run] | None = None) -> Picture | None:
        attr, alt, (src, _title) = image_node["c"]
        source = dict(attr[2]).get("source", src)
        if not src.startswith("data:"):
            warn(f"image {src} is not embedded in the AST, skipped")
            return None
        header, payload = src.split(",", 1)
        data = base64.b64decode(payload)
        attributes = dict(attr[2])
        return Picture(data=data, caption=caption or [], source=source,
                       width=attributes.get("width"), height=attributes.get("height"))

    def paragraph(self, nodes: list, role: str, depth: int = 0) -> list:
        """A paragraph; display formulas are split off into paragraphs of their own."""
        if len(nodes) == 1 and nodes[0]["t"] == "Image":
            picture = self.picture(nodes[0])
            return [picture] if picture else []
        runs = self.inlines(nodes)
        if not runs:
            return []
        if role == "text" and all(r.bold for r in runs if r.text.strip() or r.math):
            role = "heading"
            runs = [_replace(r, bold=False) for r in runs]
        out, current = [], []
        for run in runs:
            if run.math is not None and run.display:
                if any(r.text.strip() or r.math for r in current):
                    out.append(Para(current, role, depth))
                out.append(Para([run], "text", depth, align="ctr"))
                current = []
            else:
                current.append(run)
        if any(r.text.strip() or r.math or r.br for r in current):
            out.append(Para(current, role, depth))
        return out

    def list_items(self, items: list, role: str, depth: int, start: int | None = None) -> list:
        out = []
        for number, item in enumerate(items):
            first = True
            for block in item:
                t, c = block["t"], block.get("c")
                if t in ("Para", "Plain"):
                    for element in self.paragraph(c, role if first else "text", depth):
                        if isinstance(element, Para) and first:
                            element.role = role
                            element.depth = depth
                            if role == "number" and number == 0:
                                element.start = start
                            first = False
                        out.append(element)
                elif t == "BulletList":
                    out += self.list_items(c, "bullet", depth + 1)
                elif t == "OrderedList":
                    out += self.list_items(c[1], "number", depth + 1, c[0][0])
                else:
                    out += self.blocks([block])
        return out

    def table(self, c: list) -> Table:
        _attr, caption, colspecs, head, bodies, _foot = c
        aligns = [{"AlignRight": "r", "AlignCenter": "ctr"}.get(spec[0]["t"], "l") for spec in colspecs]

        def row(node) -> list[list[Run]]:
            cells = []
            for cell in node[1]:
                runs: list[Run] = []
                for block in cell[4]:
                    if block["t"] in ("Para", "Plain"):
                        if runs:
                            runs.append(Run(br=True))
                        runs += self.inlines(block["c"])
                cells.append(runs)
            return cells

        header = [row(r) for r in head[1]]
        rows = []
        for body in bodies:
            rows += [row(r) for r in body[2]] + [row(r) for r in body[3]]
        caption_runs: list[Run] = []
        for block in caption[1]:
            if block["t"] in ("Para", "Plain"):
                caption_runs += self.inlines(block["c"])
        return Table(header, rows, aligns, caption_runs)

    def figure(self, c: list) -> list:
        _attr, caption, content = c
        caption_runs: list[Run] = []
        for block in caption[1]:
            if block["t"] in ("Para", "Plain"):
                caption_runs += self.inlines(block["c"])
        out = []
        for item in self.blocks(content):
            if isinstance(item, Picture) and not item.caption:
                item.caption = caption_runs
            out.append(item)
        return out

    def blocks(self, nodes: list, notes: list[str] | None = None) -> list:
        out: list = []
        for node in nodes:
            t, c = node["t"], node.get("c")
            if t in ("Para", "Plain"):
                out += self.paragraph(c, "text")
            elif t == "Header":
                runs = [_replace(r, bold=False) for r in self.inlines(c[2])]
                out.append(Para(runs, "subheading" if c[0] <= 3 else "heading"))
            elif t == "BulletList":
                out += self.list_items(c, "bullet", 0)
            elif t == "OrderedList":
                out += self.list_items(c[1], "number", 0, c[0][0])
            elif t == "CodeBlock":
                classes = c[0][1]
                out.append(Code(c[1], classes[0] if classes else None))
            elif t == "BlockQuote":
                paras = [p for p in self.blocks(c) if isinstance(p, Para)]
                out.append(Quote(paras))
            elif t == "Table":
                out.append(self.table(c))
            elif t == "Figure":
                out += self.figure(c)
            elif t == "LineBlock":
                runs: list[Run] = []
                for line in c:
                    if runs:
                        runs.append(Run(br=True))
                    runs += self.inlines(line)
                out.append(Para(runs))
            elif t == "DefinitionList":
                for term, definitions in c:
                    out.append(Para(self.inlines(term), "heading"))
                    for definition in definitions:
                        out += self.blocks(definition)
            elif t == "Div":
                classes = c[0][1]
                if "hidden" in classes:
                    continue
                if "notes" in classes:
                    if notes is not None:
                        notes += [plain(p.runs) for p in self.blocks(c[1]) if isinstance(p, Para)]
                elif "columns" in classes:
                    cols = [
                        self.blocks(b["c"][1], notes)
                        for b in c[1]
                        if b["t"] == "Div" and "column" in b["c"][0][1]
                    ]
                    if cols:
                        out.append(Columns(cols))
                else:
                    out += self.blocks(c[1], notes)
        return out

    # ---- slides ----

    def slides(self) -> list[Slide]:
        slides: list[Slide] = []
        pending: list[dict] = []
        current: Slide | None = None

        def flush() -> None:
            nonlocal pending
            if current is not None:
                current.items = self.blocks(pending, current.notes)
            elif any(self.blocks(pending)):
                slide = Slide("content", [], {})
                slide.items = self.blocks(pending, slide.notes)
                slides.append(slide)
            pending = []

        for node in self.doc["blocks"]:
            if node["t"] == "Header" and node["c"][0] <= 2:
                flush()
                level, attr, inlines = node["c"]
                current = Slide(
                    "section" if level == 1 else "content",
                    self.inlines(inlines),
                    dict(attr[2]),
                )
                slides.append(current)
            else:
                pending.append(node)
        flush()
        return slides


UNITS = {"cm": Cm(1), "mm": Cm(0.1), "in": Cm(2.54), "pt": Pt(1)}


def size_of(value: str | None, reference: int) -> int | None:
    """EMU for a size attribute: "50%" of ``reference`` or a length such as
    "5cm". Same rules as size_of() in _extensions/shared/layouts.lua."""
    if not value:
        return None
    match = re.fullmatch(r"\s*([\d.]+)\s*(%|cm|mm|in|pt)\s*", str(value))
    if not match:
        warn(f"picture size {value!r} not understood, use % or cm")
        return None
    number, unit = float(match.group(1)), match.group(2)
    if unit == "%":
        return int(reference * number / 100)
    return int(number * UNITS[unit])


def _replace(run: Run, **changes) -> Run:
    new = copy.copy(run)
    for key, value in changes.items():
        setattr(new, key, value)
    return new


def merge_runs(runs: list[Run]) -> list[Run]:
    out: list[Run] = []
    for run in runs:
        if (
            out
            and not run.br
            and run.math is None
            and not out[-1].br
            and out[-1].math is None
            and _replace(out[-1], text="") == _replace(run, text="")
        ):
            out[-1] = _replace(out[-1], text=out[-1].text + run.text)
        else:
            out.append(run)
    return out


# --------------------------------------------------------------------------
# Math: TeX -> OMML through pandoc
# --------------------------------------------------------------------------


def pandoc_command() -> list[str] | None:
    if shutil.which("quarto"):
        return ["quarto", "pandoc"]
    if shutil.which("pandoc"):
        return ["pandoc"]
    return None


def convert_math(formulas: list[tuple[str, bool]]) -> dict[tuple[str, bool], etree._Element]:
    """Returns the ``a14:m`` element for every formula pandoc could convert."""
    result: dict[tuple[str, bool], etree._Element] = {}
    command = pandoc_command()
    if not formulas or command is None:
        return result
    parts = []
    for number, (tex, display) in enumerate(formulas, 1):
        mark = "$$" if display else "$"
        parts.append(f"# f{number}\n\n{mark}{tex}{mark}\n")
    with tempfile.TemporaryDirectory() as tmp:
        source, target = Path(tmp) / "math.md", Path(tmp) / "math.pptx"
        source.write_text("\n".join(parts), encoding="utf-8")
        run = subprocess.run(
            [*command, str(source), "-f", "markdown", "-t", "pptx", "--slide-level=1", "-o", str(target)],
            capture_output=True,
            text=True,
        )
        if run.returncode != 0:
            warn(f"pandoc could not convert formulas: {run.stderr.strip()}")
            return result
        with zipfile.ZipFile(target) as archive:
            for number, key in enumerate(formulas, 1):
                try:
                    xml = archive.read(f"ppt/slides/slide{number}.xml")
                except KeyError:
                    continue
                found = etree.fromstring(xml).find(f".//{qn('a14:m')}")
                if found is not None:
                    element = etree.fromstring(etree.tostring(found))
                    result[key] = element
    return result


def render_math_png(tex: str) -> bytes | None:
    """Fallback for display formulas pandoc cannot convert."""
    if not (shutil.which("pdflatex") and shutil.which("pdftoppm")):
        return None
    document = (
        "\\documentclass[preview,border=2pt]{standalone}\n"
        "\\usepackage{amsmath,amssymb}\n\\begin{document}\n"
        f"$\\displaystyle {tex}$\n\\end{{document}}\n"
    )
    with tempfile.TemporaryDirectory() as tmp:
        (Path(tmp) / "f.tex").write_text(document, encoding="utf-8")
        run = subprocess.run(
            ["pdflatex", "-interaction=batchmode", "-halt-on-error", "f.tex"],
            cwd=tmp,
            capture_output=True,
        )
        if run.returncode != 0:
            return None
        subprocess.run(["pdftoppm", "-png", "-r", "300", "-singlefile", "f.pdf", "f"], cwd=tmp, check=False)
        png = Path(tmp) / "f.png"
        return png.read_bytes() if png.exists() else None


# --------------------------------------------------------------------------
# Geometry
# --------------------------------------------------------------------------


@dataclass
class Rect:
    x: int
    y: int
    w: int
    h: int

    @property
    def bottom(self) -> int:
        return self.y + self.h


@dataclass
class Region:
    rect: Rect
    placeholder: object | None
    sizes: dict[str, float]
    style: etree._Element  # a:lstStyle for generated text boxes


# --------------------------------------------------------------------------
# Builder
# --------------------------------------------------------------------------


class Builder:
    def __init__(self, reader: Reader, config: dict, deck_dir: Path):
        self.reader = reader
        self.config = config
        self.deck_dir = deck_dir
        lang = reader.meta_text("lang", "en")[:2].lower()
        self.lang = lang if lang in config["templates"] else "en"
        self.strings = config["strings"][self.lang]
        self.assets = PROJECT / config["assets"]
        self.prs = Presentation(str(self.assets / config["templates"][self.lang]))
        self.levels = config["levels"]
        self.fonts = config["fonts"]
        self.gap = Cm(config["gap-cm"])
        self.column_gap = Cm(config["column-gap-cm"])
        self.content_bottom = Cm(config["content-bottom-cm"])
        self.math: dict[tuple[str, bool], etree._Element] = {}
        self.math_source: dict[bytes, str] = {}
        self.sections = 0
        self.master_levels = self.prs.slide_master._element.find(".//p:txStyles/p:bodyStyle", NS)
        self.master_body_pr = None
        for ph in self.prs.slide_master.placeholders:
            if ph.placeholder_format.type == PP_PLACEHOLDER.BODY:
                self.master_body_pr = ph._element.find(".//a:bodyPr", NS)

    # ---- template handling ----

    def clear_slides(self) -> None:
        id_list = self.prs.slides._sldIdLst
        for slide_id in list(id_list):
            self.prs.part.drop_rel(slide_id.rId)
            id_list.remove(slide_id)
        # sections of the sample deck refer to the removed slides
        for section_list in self.prs.part._element.findall(f".//{qn('p14:sectionLst')}"):
            ext = section_list.getparent()
            ext.getparent().remove(ext)

    def layout(self, name: str):
        spec = self.config["layouts"][name]
        return spec, self.prs.slide_layouts[spec["layout"] - 1]

    def new_slide(self, name: str, footers: bool = True):
        spec, layout = self.layout(name)
        slide = self.prs.slides.add_slide(layout)
        if footers:
            self.add_footers(slide, layout)
        return spec, layout, slide

    def add_footers(self, slide, layout) -> None:
        for ph in layout.placeholders:
            kind = ph.placeholder_format.type
            if kind not in FOOTER_TYPES:
                continue
            element = copy.deepcopy(ph._element)
            # ids of the layout may collide with those of the slide's shapes
            used = [int(i) for i in slide.shapes._spTree.xpath("//p:cNvPr/@id") if i.isdigit()]
            element.find(f".//{qn('p:cNvPr')}").set("id", str(max(used, default=1) + 1))
            slide.shapes._spTree.append(element)
            if kind == PP_PLACEHOLDER.DATE:
                self.set_plain_text(element, self.date_text())
            elif kind == PP_PLACEHOLDER.FOOTER:
                self.set_plain_text(element, self.reader.meta_text("footer", self.strings["footer"]))

    def copy_pictures(self, slide, layout) -> None:
        """Pictures of the layout (logo) would be hidden by a full size picture."""
        for picture in layout.shapes._spTree.findall(qn("p:pic")):
            element = copy.deepcopy(picture)
            blip = element.find(f".//{qn('a:blip')}")
            if blip is None or blip.get(qn("r:embed")) is None:
                continue
            image = layout.part.related_part(blip.get(qn("r:embed")))
            blip.set(qn("r:embed"), slide.part.relate_to(image, RT.IMAGE))
            used = [int(i) for i in slide.shapes._spTree.xpath("//p:cNvPr/@id") if i.isdigit()]
            element.find(f".//{qn('p:cNvPr')}").set("id", str(max(used, default=1) + 1))
            slide.shapes._spTree.append(element)

    def classification(self) -> None:
        """Text of the classification label in the footer of the master."""
        if "classification" not in self.reader.meta:
            return
        text = self.reader.meta_text("classification")
        name = self.config.get("classification-shape")
        for shape in self.prs.slide_master.shapes:
            if shape.name != name or not shape.has_text_frame:
                continue
            runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
            if not runs:
                return
            runs[0].text = text
            for run in runs[1:]:
                run.text = ""
            return
        warn(f"classification label {name!r} not found in the slide master")

    def date_text(self) -> str:
        raw = self.reader.meta_text("date")
        try:
            return dt.date.fromisoformat(raw).strftime(self.strings["date-format"])
        except ValueError:
            return raw

    def set_plain_text(self, shape_element, text: str) -> None:
        body = shape_element.find(qn("p:txBody"))
        for paragraph in body.findall(qn("a:p")):
            body.remove(paragraph)
        paragraph = etree.SubElement(body, qn("a:p"))
        run = etree.SubElement(paragraph, qn("a:r"))
        etree.SubElement(run, qn("a:rPr"), lang=self.strings["locale"])
        etree.SubElement(run, qn("a:t")).text = text

    @staticmethod
    def placeholder(slide, idx: int):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                return ph
        return None

    @staticmethod
    def remove(shape) -> None:
        if shape is not None:
            element = shape._element
            element.getparent().remove(element)

    @staticmethod
    def rect_of(layout, idx: int) -> Rect:
        for ph in layout.placeholders:
            if ph.placeholder_format.idx == idx:
                return Rect(ph.left, ph.top, ph.width, ph.height)
        raise KeyError(f"layout {layout.name!r} has no placeholder idx {idx}")

    def list_style(self, layout, idx: int) -> etree._Element:
        """Text style of a body placeholder, for text boxes that imitate it."""
        style = etree.Element(qn("a:lstStyle"), nsmap={"a": A})
        levels = {}
        for level in self.master_levels:
            name = etree.QName(level).localname
            if name.startswith("lvl"):
                levels[name] = copy.deepcopy(level)
        for ph in layout.placeholders:
            if ph.placeholder_format.idx != idx:
                continue
            override = ph._element.find(".//p:txBody/a:lstStyle", NS)
            for level in override if override is not None else []:
                name = etree.QName(level).localname
                if name in levels:
                    merge_level(levels[name], level)
        for name in sorted(levels):
            style.append(levels[name])
        return style

    # ---- paragraphs ----

    def level_of(self, para: Para) -> int:
        if para.role in ("bullet", "number"):
            return self.levels["list"] + min(para.depth, 2)
        return self.levels[para.role]

    def run_properties(self, slide, run: Run, size: float | None = None, colour: str | None = None):
        props = etree.Element(qn("a:rPr"), lang=self.strings["locale"], dirty="0")
        if size:
            props.set("sz", str(int(round(size * 100))))
        if run.italic:
            props.set("i", "1")
        if run.underline:
            props.set("u", "sng")
        if run.strike:
            props.set("strike", "sngStrike")
        if run.script:
            props.set("baseline", "30000" if run.script == "super" else "-25000")
        colour = run.colour or colour
        if colour:
            fill = etree.SubElement(props, qn("a:solidFill"))
            etree.SubElement(fill, qn("a:srgbClr"), val=colour)
        if run.code:
            etree.SubElement(props, qn("a:latin"), typeface=self.fonts["code"])
            etree.SubElement(props, qn("a:cs"), typeface=self.fonts["code"])
        elif run.bold:
            bold = self.fonts.get("bold")
            if bold:
                etree.SubElement(props, qn("a:latin"), typeface=bold)
            else:
                props.set("b", "1")
        if run.link and slide is not None:
            rel = slide.part.relate_to(run.link, RT.HYPERLINK, is_external=True)
            etree.SubElement(props, qn("a:hlinkClick")).set(qn("r:id"), rel)
        return props

    def paragraph(
        self,
        slide,
        para: Para,
        level: int | None = None,
        size: float | None = None,
        colour: str | None = None,
        plain_list: bool = False,
    ):
        element = etree.Element(qn("a:p"))
        props = etree.SubElement(element, qn("a:pPr"))
        if level is not None:
            props.set("lvl", str(level))
        if para.align:
            props.set("algn", para.align)
        if para.role == "number" and not plain_list:
            indent = 360000
            props.set("marL", str(indent * (min(para.depth, 2) + 1)))
            props.set("indent", str(-indent))
            etree.SubElement(props, qn("a:buFont"), typeface="+mn-lt")
            number = etree.SubElement(props, qn("a:buAutoNum"), type="arabicPeriod")
            if para.start and para.start != 1:
                number.set("startAt", str(para.start))
        elif para.role == "text" and para.depth > 0 and level is not None:
            # continuation paragraph of a list item: indented, without bullet
            props.set("indent", "0")
            etree.SubElement(props, qn("a:buNone"))
        for run in para.runs:
            if run.br:
                etree.SubElement(element, qn("a:br"))
            elif run.math is not None:
                math_element = self.math.get((run.math, run.display))
                if math_element is not None:
                    element.append(copy.deepcopy(math_element))
                else:
                    node = etree.SubElement(element, qn("a:r"))
                    node.append(self.run_properties(slide, _replace(run, code=True), size, colour))
                    etree.SubElement(node, qn("a:t")).text = run.math
            elif run.text:
                node = etree.SubElement(element, qn("a:r"))
                node.append(self.run_properties(slide, run, size, colour))
                etree.SubElement(node, qn("a:t")).text = clean(run.text)
        return element

    def write(self, slide, body, paragraphs: list) -> None:
        """Replaces the paragraphs of a ``txBody`` element."""
        for paragraph in body.findall(qn("a:p")):
            body.remove(paragraph)
        for paragraph in paragraphs:
            body.append(paragraph)
        if not paragraphs:
            etree.SubElement(body, qn("a:p"))

    def write_paras(self, slide, shape, paras: list[Para]) -> None:
        body = shape._element.find(qn("p:txBody"))
        self.write(slide, body, [self.paragraph(slide, p, self.level_of(p)) for p in paras])

    def write_levels(self, slide, shape, lines: list[tuple[int, str]]) -> None:
        """Title like placeholders: (level, text) pairs."""
        body = shape._element.find(qn("p:txBody"))
        paragraphs = [self.paragraph(slide, Para([Run(text=text)]), level) for level, text in lines if text]
        self.write(slide, body, paragraphs)

    # ---- size estimates ----

    def text_height(self, paras: list[Para], width: int, sizes: dict[str, float]) -> int:
        total = 0.0
        width_pt = Emu(width).pt
        for para in paras:
            role = "list" if para.role in ("bullet", "number") else para.role
            size = sizes[role]
            if para.is_display_math:
                total += size * 3.2
                continue
            indent = 14 * (para.depth + 1) if role == "list" else 0
            per_line = max(8, int((width_pt - indent) / (size * 0.5)))
            text = plain(para.runs)
            lines = sum(max(1, math.ceil(len(part) / per_line)) for part in text.split("\n"))
            # line spacing 110 % of the line height; the master sets 19 pt
            # after subheadings and text paragraphs, nothing after list items
            total += lines * size * 1.32
            total += 19 if para.role in ("subheading", "text") else 0
        return Pt(total)

    def code_layout(self, code: Code, width: int, max_height: int | None = None) -> tuple[float, int]:
        """Font size and height of a code block; the font shrinks down to
        ``code-size-min`` to fit the width and, if given, the height."""
        lines = code.text.rstrip("\n").split("\n")
        longest = max((len(line.expandtabs(4)) for line in lines), default=1)
        inner_pt = Emu(width).pt - 16
        smallest = float(self.fonts["code-size-min"])
        size = float(self.fonts["code-size"])

        def height(size: float) -> int:
            per_line = max(1, int(inner_pt / (0.6 * size)))
            rows = sum(max(1, math.ceil(len(line.expandtabs(4)) / per_line)) for line in lines)
            return Pt(rows * size * 1.2 + 16)

        fitting = inner_pt / (0.6 * max(longest, 1))
        if fitting < size:
            size = max(smallest, math.floor(fitting * 2) / 2)
        while max_height is not None and height(size) > max_height and size > smallest:
            size = max(smallest, size - 0.5)
        return size, height(size)

    def table_height(self, table: Table) -> int:
        rows = len(table.header) + len(table.rows)
        height = Pt(rows * (self.fonts["table-size"] * 1.2 + 10))
        if table.caption:
            height += Pt(self.fonts["caption-size"] * 1.8)
        return height

    def natural_height(self, item, width: int, sizes: dict[str, float]) -> int | None:
        """Height of an element, or None when it adapts to the space left."""
        if isinstance(item, list):
            return self.text_height(item, width, sizes)
        if isinstance(item, Code):
            return self.code_layout(item, width)[1]
        if isinstance(item, Table):
            return self.table_height(item)
        if isinstance(item, Quote):
            return self.text_height(item.paras, width, sizes)
        if isinstance(item, Columns):
            count = len(item.columns)
            column_width = (width - self.column_gap * (count - 1)) // count
            heights = []
            for column in item.columns:
                segments = self.segments(column)
                parts = [self.natural_height(s, column_width, sizes) for s in segments]
                if any(p is None for p in parts):
                    return None
                heights.append(sum(parts) + self.gap * max(0, len(parts) - 1))
            return max(heights, default=0)
        return None  # Picture

    # ---- elements ----

    @staticmethod
    def segments(items: list) -> list:
        """Groups consecutive paragraphs; every other element stands alone."""
        out: list = []
        for item in items:
            if isinstance(item, Para):
                if out and isinstance(out[-1], list):
                    out[-1].append(item)
                else:
                    out.append([item])
            else:
                out.append(item)
        return out

    def text_box(self, slide, rect: Rect, style: etree._Element | None):
        shape = slide.shapes.add_textbox(rect.x, rect.y, rect.w, rect.h)
        body = shape._element.find(qn("p:txBody"))
        body_pr = body.find(qn("a:bodyPr"))
        for child in list(body_pr):
            body_pr.remove(child)
        for key in ("lIns", "tIns", "rIns", "bIns"):
            body_pr.set(key, "0")
        body_pr.set("wrap", "square")
        body_pr.set("rtlCol", "0")
        if style is not None:
            old = body.find(qn("a:lstStyle"))
            body.replace(old, copy.deepcopy(style))
        return shape

    def add_text(self, slide, region: Region, rect: Rect, paras: list[Para], use_placeholder: bool):
        if use_placeholder and region.placeholder is not None:
            shape = region.placeholder
            shape.left, shape.top, shape.width, shape.height = rect.x, rect.y, rect.w, rect.h
        else:
            shape = self.text_box(slide, rect, region.style)
        self.write_paras(slide, shape, paras)
        return shape

    def add_quote(self, slide, region: Region, rect: Rect, quote: Quote) -> None:
        shape = self.text_box(slide, rect, region.style)
        paras = [Para([_replace(r, italic=True) for r in p.runs], "text") for p in quote.paras]
        self.write_paras(slide, shape, paras)

    def add_code(self, slide, rect: Rect, code: Code) -> None:
        size, _height = self.code_layout(code, rect.w, rect.h)
        shape = slide.shapes.add_textbox(rect.x, rect.y, rect.w, rect.h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb("F2F2F2")
        body = shape._element.find(qn("p:txBody"))
        body_pr = body.find(qn("a:bodyPr"))
        for child in list(body_pr):
            body_pr.remove(child)
        for key in ("lIns", "tIns", "rIns", "bIns"):
            body_pr.set(key, str(Pt(8)))
        body_pr.set("wrap", "square")
        paragraphs = []
        for line in highlight(code):
            para = Para(line or [Run(text="")])
            element = self.paragraph(slide, para, None, size, "262626")
            if not line:
                end = etree.SubElement(element, qn("a:endParaRPr"), lang=self.strings["locale"])
                end.set("sz", str(int(size * 100)))
            paragraphs.append(element)
        self.write(slide, body, paragraphs)

    def add_table(self, slide, rect: Rect, table: Table) -> None:
        rows = table.header + table.rows
        if not rows:
            return
        columns = max(len(r) for r in rows)
        caption_height = Pt(self.fonts["caption-size"] * 1.8) if table.caption else 0
        lengths = [
            max([len(plain(r[c])) for r in rows if c < len(r)] + [3]) for c in range(columns)
        ]
        lengths = [min(max(n, 6), 40) for n in lengths]
        natural = Pt(sum(lengths) * self.fonts["table-size"] * 0.6 + columns * 16)
        width = min(rect.w, max(natural, rect.w // 3))
        height = min(rect.h - caption_height, self.table_height(table) - caption_height)
        frame = slide.shapes.add_table(len(rows), columns, rect.x, rect.y, width, height)
        grid = frame.table
        grid.first_row = bool(table.header)
        for index, column in enumerate(grid.columns):
            column.width = int(width * lengths[index] / sum(lengths))
        for r, row in enumerate(rows):
            head = r < len(table.header)
            for c in range(columns):
                runs = row[c] if c < len(row) else []
                if head:
                    runs = [_replace(run, bold=False) for run in runs]
                align = table.aligns[c] if c < len(table.aligns) else "l"
                para = Para(runs or [Run(text="")], align=None if align == "l" else align)
                body = grid.cell(r, c).text_frame._txBody
                element = self.paragraph(slide, para, None, self.fonts["table-size"])
                if head and self.fonts.get("bold"):
                    for props in element.iter(qn("a:rPr")):
                        if props.find(qn("a:latin")) is None:
                            etree.SubElement(props, qn("a:latin"), typeface=self.fonts["bold"])
                self.write(slide, body, [element])
        if table.caption:
            box = Rect(rect.x, rect.y + height + Pt(3), rect.w, caption_height)
            self.add_caption(slide, box, table.caption, "l")

    def add_caption(self, slide, rect: Rect, runs: list[Run], align: str = "ctr") -> None:
        shape = self.text_box(slide, rect, None)
        body = shape._element.find(qn("p:txBody"))
        para = Para(runs, align=align)
        self.write(slide, body, [self.paragraph(slide, para, None, self.fonts["caption-size"], "595959")])

    def add_picture(self, slide, rect: Rect, picture: Picture) -> None:
        caption_height = Pt(self.fonts["caption-size"] * 1.8) if picture.caption else 0
        data = raster(picture)
        if data is None:
            return
        with Image.open(io.BytesIO(data)) as image:
            pixel_width, pixel_height = image.size
        available_width = min(size_of(picture.width, rect.w) or rect.w, rect.w)
        available_height = max(rect.h - caption_height, Cm(1))
        scale = min(available_width / pixel_width, available_height / pixel_height)
        width, height = int(pixel_width * scale), int(pixel_height * scale)
        left = rect.x + (rect.w - width) // 2
        slide.shapes.add_picture(io.BytesIO(data), left, rect.y, width, height)
        if picture.caption:
            box = Rect(rect.x, rect.y + height + Pt(3), rect.w, caption_height)
            self.add_caption(slide, box, picture.caption)

    def add_columns(self, slide, region: Region, rect: Rect, item: Columns) -> None:
        count = len(item.columns)
        width = (rect.w - self.column_gap * (count - 1)) // count
        for index, column in enumerate(item.columns):
            box = Rect(rect.x + index * (width + self.column_gap), rect.y, width, rect.h)
            self.fill(slide, Region(box, None, region.sizes, region.style), column, "")

    # ---- regions ----

    def fill(self, slide, region: Region, items: list, title: str) -> None:
        """Stacks the elements of a region from top to bottom."""
        segments = self.segments(items)
        if not segments:
            self.remove(region.placeholder)
            return
        if len(segments) == 1 and isinstance(segments[0], list) and region.placeholder is not None:
            self.write_paras(slide, region.placeholder, segments[0])
            return

        rect = region.rect
        heights = [self.natural_height(s, rect.w, region.sizes) for s in segments]
        for index, segment in enumerate(segments):
            if isinstance(segment, Picture) and size_of(segment.height, rect.h):
                caption = Pt(self.fonts["caption-size"] * 1.8) if segment.caption else 0
                heights[index] = size_of(segment.height, rect.h) + caption
        gaps = self.gap * (len(segments) - 1)
        fixed = sum(h for h in heights if h is not None)
        flexible = [i for i, h in enumerate(heights) if h is None]
        if flexible:
            share = max((rect.h - fixed - gaps) // len(flexible), Cm(2))
            for index in flexible:
                heights[index] = share
        excess = sum(heights) + gaps - rect.h
        if excess > 0:
            # code blocks give way first: their font shrinks down to the minimum
            for index, segment in enumerate(segments):
                if isinstance(segment, Code) and excess > 0:
                    target = max(heights[index] - excess, Cm(1))
                    _size, smaller = self.code_layout(segment, rect.w, target)
                    excess -= heights[index] - smaller
                    heights[index] = smaller
        if sum(heights) + gaps > rect.h + Cm(0.3) and title:
            warn(f"slide {title!r}: content is higher than the slide, split it or shorten it")

        top = rect.y
        placeholder_free = region.placeholder is not None
        for segment, height in zip(segments, heights):
            box = Rect(rect.x, top, rect.w, height)
            if isinstance(segment, list):
                self.add_text(slide, region, box, segment, placeholder_free)
                placeholder_free = False
            elif isinstance(segment, Code):
                self.add_code(slide, box, segment)
            elif isinstance(segment, Table):
                self.add_table(slide, box, segment)
            elif isinstance(segment, Picture):
                self.add_picture(slide, box, segment)
            elif isinstance(segment, Quote):
                self.add_quote(slide, region, box, segment)
            elif isinstance(segment, Columns):
                self.add_columns(slide, region, box, segment)
            top += height + self.gap
        if placeholder_free:
            self.remove(region.placeholder)

    def region(self, slide, layout, spec: dict, idx: int, keep_placeholder: bool = True) -> Region:
        rect = self.rect_of(layout, idx)
        rect = Rect(rect.x, rect.y, rect.w, self.content_bottom - rect.y)
        placeholder = self.placeholder(slide, idx)
        if not keep_placeholder:
            self.remove(placeholder)
            placeholder = None
        return Region(rect, placeholder, spec["sizes"], self.list_style(layout, idx))

    # ---- slides ----

    def title_slide(self) -> None:
        meta = self.reader
        title = meta.meta_text("title")
        if not title:
            return
        spec, layout, slide = self.new_slide("title", footers=False)
        lines = [(0, title), (1, DASH), (2, meta.meta_text("subtitle"))]
        people = " | ".join(x for x in (meta.meta_text("author"), meta.meta_text("institute")) if x)
        lines += [(2, people), (2, self.date_text())]
        self.write_levels(slide, self.placeholder(slide, spec["text"]), lines)

        picture = self.placeholder(slide, spec["picture"])
        custom = meta.meta_text("title-image")
        if custom:
            path = self.deck_dir / custom
        else:
            variant = meta.meta_text("iis-titlepage", "1")
            images = {str(k): v for k, v in self.config["title-images"].items()}
            path = self.assets / images.get(variant, images["1"])
        if path.is_file():
            picture.insert_picture(str(path))
            self.copy_pictures(slide, layout)
        else:
            warn(f"title image {path} not found")
            self.remove(picture)

    def section_slide(self, source: Slide) -> None:
        self.sections += 1
        spec, _layout, slide = self.new_slide("section")
        lines = [(0, f"{self.sections:02d}"), (1, DASH), (2, source.title_text)]
        self.write_levels(slide, self.placeholder(slide, spec["text"]), lines)
        self.notes(slide, source)

    def closing_slide(self) -> None:
        spec, _layout, slide = self.new_slide("closing")
        text = self.reader.meta_text("closing-text", self.strings["closing"])
        self.write_levels(slide, self.placeholder(slide, spec["text"]), [(0, text), (1, DASH)])

    def quote_slide(self, source: Slide) -> bool:
        quote = next((i for i in source.items if isinstance(i, Quote)), None)
        paras = quote.paras if quote else [i for i in source.items if isinstance(i, Para)]
        if not paras:
            return False
        spec, _layout, slide = self.new_slide("quote")
        body = self.placeholder(slide, spec["text"])._element.find(qn("p:txBody"))
        elements = []
        for level, para in enumerate(paras[:3]):
            elements.append(self.paragraph(slide, Para(para.runs), level, plain_list=True))
        self.write(slide, body, elements)
        self.notes(slide, source)
        return True

    def heading(self, slide, spec: dict, source: Slide) -> None:
        title = self.placeholder(slide, spec["title"])
        subline = self.placeholder(slide, spec["subline"])
        if source.title:
            body = title._element.find(qn("p:txBody"))
            self.write(slide, body, [self.paragraph(slide, Para(source.title), None)])
        else:
            self.remove(title)
        text = source.attrs.get("subtitle", "")
        if text:
            self.write_levels(slide, subline, [(0, text)])
        else:
            self.remove(subline)

    def content_slide(self, source: Slide) -> None:
        name = source.attrs.get("layout", "content")
        items = source.items
        if name == "quote" and self.quote_slide(source):
            return

        picture = None
        if name in ("image-left", "image-right"):
            picture = next((i for i in items if isinstance(i, Picture)), None)
            if picture is None:
                warn(f"slide {source.title_text!r}: layout {name} without image")
                name = "content"
            else:
                items = [i for i in items if i is not picture]

        native_columns = None
        if name in ("content", "content-large", "columns") and len(items) == 1:
            only = items[0]
            if isinstance(only, Columns) and 2 <= len(only.columns) <= 4:
                native_columns = only
                name = f"columns-{len(only.columns)}"
        if name not in self.config["layouts"] or name in ("columns", "full-image", "blank", "quote"):
            blank = name == "blank"
            name = "content"
        else:
            blank = False

        spec, layout, slide = self.new_slide(name)
        if blank:
            source = Slide(source.kind, [], {}, source.items, source.notes)
        self.heading(slide, spec, source)

        if native_columns is not None:
            for idx, column in zip(spec["body"], native_columns.columns):
                self.fill(slide, self.region(slide, layout, spec, idx), column, source.title_text)
        else:
            region = self.region(slide, layout, spec, spec["body"][0])
            if blank:
                top = self.rect_of(layout, spec["title"]).y
                region.rect = Rect(region.rect.x, top, region.rect.w, self.content_bottom - top)
            self.fill(slide, region, items, source.title_text)

        if picture is not None:
            if "picture" in spec:
                data = raster(picture)
                placeholder = self.placeholder(slide, spec["picture"])
                if data is not None:
                    placeholder.insert_picture(io.BytesIO(data))
                else:
                    self.remove(placeholder)
            else:
                idx = spec["picture-region"]
                target = self.region(slide, layout, spec, idx, keep_placeholder=False)
                self.add_picture(slide, target.rect, picture)
        self.notes(slide, source)

    def notes(self, slide, source: Slide) -> None:
        if source.notes:
            slide.notes_slide.notes_text_frame.text = "\n".join(source.notes)

    # ---- math ----

    def prepare_math(self, slides: list[Slide]) -> None:
        formulas: list[tuple[str, bool]] = []

        def visit(items: list) -> None:
            for item in items:
                if isinstance(item, Para):
                    for run in item.runs:
                        if run.math is not None and (run.math, run.display) not in formulas:
                            formulas.append((run.math, run.display))
                elif isinstance(item, Quote):
                    visit(item.paras)
                elif isinstance(item, Columns):
                    for column in item.columns:
                        visit(column)

        for slide in slides:
            visit(slide.items)
        self.math = convert_math(formulas)
        self.math_source = {etree.tostring(e, method="c14n", exclusive=True): tex for (tex, _), e in self.math.items()}
        failed = [key for key in formulas if key not in self.math]
        for tex, display in failed:
            warn(f"formula could not be converted to a PowerPoint equation: {tex}")

        def replace(items: list) -> list:
            out = []
            for item in items:
                if isinstance(item, Para) and item.is_display_math:
                    key = (item.runs[0].math, True)
                    if key not in self.math:
                        png = render_math_png(item.runs[0].math)
                        if png:
                            out.append(Picture(png, source="formula"))
                            continue
                elif isinstance(item, Columns):
                    item.columns = [replace(column) for column in item.columns]
                out.append(item)
            return out

        if failed:
            for slide in slides:
                slide.items = replace(slide.items)

    def wrap_math(self, slide) -> None:
        """Equations need an mc:AlternateContent wrapper with a text fallback."""
        tree = slide.shapes._spTree
        for shape in list(tree):
            if shape.tag != qn("p:sp") or shape.find(f".//{qn('a14:m')}") is None:
                continue
            wrapper = etree.Element(qn("mc:AlternateContent"), nsmap={"mc": MC})
            choice = etree.SubElement(wrapper, qn("mc:Choice"), nsmap={"a14": A14})
            choice.set("Requires", "a14")
            fallback = etree.SubElement(wrapper, qn("mc:Fallback"))
            tree.replace(shape, wrapper)
            choice.append(shape)
            plain_shape = copy.deepcopy(shape)
            for equation in plain_shape.findall(f".//{qn('a14:m')}"):
                run = etree.Element(qn("a:r"))
                etree.SubElement(run, qn("a:rPr"), lang=self.strings["locale"], dirty="0")
                source = self.math_source.get(etree.tostring(equation, method="c14n", exclusive=True))
                etree.SubElement(run, qn("a:t")).text = source or "".join(equation.itertext())
                equation.getparent().replace(equation, run)
            fallback.append(plain_shape)

    # ---- entry point ----

    def build(self, output: Path) -> None:
        slides = self.reader.slides()
        self.prepare_math(slides)
        self.clear_slides()
        self.classification()
        self.title_slide()
        for source in slides:
            if source.kind == "section":
                self.section_slide(source)
            else:
                self.content_slide(source)
        if self.reader.meta.get("closing", True):
            self.closing_slide()
        for slide in self.prs.slides:
            self.wrap_math(slide)
        self.prs.core_properties.title = self.reader.meta_text("title")
        self.prs.core_properties.author = self.reader.meta_text("author")
        self.prs.save(str(output))


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def _rgb(value: str):
    from pptx.dml.color import RGBColor

    return RGBColor.from_string(value)


def clean(text: str) -> str:
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


def exclusive(tag: str) -> set[str]:
    """Tags that cannot stand next to each other."""
    for group in (BULLET_TAGS, FILL_TAGS):
        if tag in group:
            return group
    return {tag}


def overlay(base, override, order: list[str]) -> None:
    base.attrib.update(override.attrib)
    for child in override:
        if child.tag == qn("a:defRPr") and base.find(child.tag) is not None:
            overlay(base.find(child.tag), child, RUN_ORDER)
            continue
        for old in [e for e in base if e.tag in exclusive(child.tag)]:
            base.remove(old)
        base.append(copy.deepcopy(child))
    rank = {name: index for index, name in enumerate(order)}
    base[:] = sorted(base, key=lambda e: rank.get(etree.QName(e).localname, len(order)))


def merge_level(base, override) -> None:
    """Applies the paragraph level of a layout on top of the master's level."""
    overlay(base, override, PARAGRAPH_ORDER)


def raster(picture: Picture) -> bytes | None:
    """Image data python-pptx can embed; PDF figures are converted to PNG."""
    data = picture.data
    if data[:4] == b"%PDF":
        if not shutil.which("pdftoppm"):
            warn(f"pdftoppm missing, cannot convert {picture.source}")
            return None
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "f.pdf").write_bytes(data)
            subprocess.run(["pdftoppm", "-png", "-r", "200", "-singlefile", "f.pdf", "f"], cwd=tmp, check=False)
            png = Path(tmp) / "f.png"
            return png.read_bytes() if png.exists() else None
    try:
        with Image.open(io.BytesIO(data)) as image:
            image.verify()
    except Exception:
        warn(f"image format of {picture.source} is not supported in PPTX, skipped")
        return None
    return data


def highlight(code: Code) -> list[list[Run]]:
    """Lines of coloured runs."""
    text = code.text.rstrip("\n").expandtabs(4)
    tokens: list[tuple[str, str]] = [("", text)]
    if code.lang:
        try:
            from pygments import lex
            from pygments.lexers import get_lexer_by_name

            lexer = get_lexer_by_name(code.lang, stripnl=False, ensurenl=False)
            tokens = [(str(kind), value) for kind, value in lex(text, lexer)]
        except Exception:
            pass
    lines: list[list[Run]] = [[]]
    for kind, value in tokens:
        style = next((s for prefix, s in CODE_COLOURS if kind.startswith(prefix)), None)
        for index, part in enumerate(value.split("\n")):
            if index:
                lines.append([])
            if part:
                run = Run(text=part, code=True)
                if style:
                    run.colour, run.bold, run.italic = style
                    # bold is expressed by the font family elsewhere; code keeps its font
                    run.bold = False
                lines[-1].append(run)
    return lines


def build(ast_path: Path, output: Path, config_path: Path = DEFAULT_CONFIG) -> None:
    doc = json.loads(ast_path.read_text(encoding="utf-8"))
    config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    Builder(Reader(doc), config, ast_path.resolve().parent).build(output)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("ast", type=Path, help="pandoc AST as JSON")
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    args = parser.parse_args()
    build(args.ast, args.output, args.config)
    print(f"md2pptx: wrote {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
