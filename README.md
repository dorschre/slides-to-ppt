# slides-to-ppt
A small tool to create drafts for presentations in markdown.

Slides are written once as a Quarto document (`.qmd`) and rendered with the corporate
templates in `assets/`.

| Layout     | Format        | Output          | Template                                   |
|------------|---------------|-----------------|--------------------------------------------|
| FAU        | `fau-beamer`  | `deck-fau.pdf`  | `assets/fau/fau-beamer`                    |
| Fraunhofer | `iis-beamer`  | `deck-iis.pdf`  | `assets/fraunhofer/beamer theme IIS ...`   |
| Fraunhofer | `iis-pptx`    | `deck-iis.pptx` | `assets/fraunhofer/*-DE.pptx`, `*-EN.pptx` |

There is no PPTX output in the FAU layout, because `assets/` contains no FAU PowerPoint
template.

`assets/` is never modified and is excluded from git (proprietary fonts and logos).

## Installation of Requirements

- Quarto (version 1.4 or newer): download it from https://quarto.org/docs/get-started/
- Install it. E.g. via ```sudo apt install ./Downloads/quarto.deb```
- Python packages, in a virtual environment inside the project:

  ```sh
  uv venv .venv --python 3.12
  uv pip install --python .venv/bin/python -r requirements.txt
  ```

- LaTeX: `pdflatex` for FAU, `lualatex` for Fraunhofer. On Debian/Ubuntu LuaLaTeX can
  load fonts only with

  ```sh
  sudo apt install texlive-luatex
  ```

Used by the tool:

| Requirement   | Used for                                                      |
|---------------|---------------------------------------------------------------|
| Quarto        | reads the Markdown, runs code cells, bundles its own pandoc   |
| python3       | hooks and PPTX generator                                      |
| python-pptx   | writes the PPTX from the Fraunhofer template                  |
| jupyter, matplotlib | executable Python cells in a deck                       |
| pdflatex      | FAU beamer                                                    |
| lualatex      | Fraunhofer beamer                                             |
| pdftoppm      | converts PDF figures for PPTX                                 |

`latexmk`, `node` and a separately installed `pandoc` are not needed.

`_environment` points Quarto to the Python of `.venv`.

## How to run

Choose the layout and the export format through the format name:

```sh
tools/render.py slides/example/deck.qmd                    # all formats listed in the deck
tools/render.py slides/example/deck.qmd --to fau-beamer    # FAU, PDF
tools/render.py slides/example/deck.qmd --to iis-beamer    # Fraunhofer, PDF
tools/render.py slides/example/deck.qmd --to iis-pptx      # Fraunhofer, PPTX
tools/render.py slides/example/deck.qmd --to iis-pptx iis-beamer -- -M lang:de
```

Arguments after `--` are passed to Quarto, e.g. `-M lang:de`, `-M iis-titlepage:2` or
`-M keep-tex:true` (keeps the generated `.tex` file).

`tools/render.py` runs `quarto render <deck> --to <format>` once per format and renames
the result to `<deck>-<brand>.<ext>`. Calling `quarto render` directly works for a
single format and writes `deck.pdf` or `deck.pptx`. Rendering several formats in one
Quarto run does not work: Quarto uses the same intermediate `deck.tex` and the same
`deck.pdf` for both beamer formats.

A new deck is a directory below `slides/` with a `.qmd` file. Pictures of the deck go
into a subdirectory such as `figures/`. The names `theme`, `logos`, `iis-images`,
`iis-pptx`, `styles` and `template-art` are reserved: `tools/stage.py` links the brand assets
under these names into every deck directory before rendering.

## Markdown commands

| Element                           | Markdown                                                  |
|-----------------------------------|-----------------------------------------------------------|
| Section divider                   | `# Heading`                                               |
| Presentation slide main header    | `## Heading`                                              |
| Presentation slide sub header     | `## Heading {subtitle="Subline"}`                         |
| Subheading on a slide             | `### Heading` (beamer: block)                             |
| Paragraph heading                 | a paragraph that is bold as a whole: `**Heading**`        |
| Add image                         | `![Caption](figures/picture.png)`                         |
| Add table                         | pipe table, see the example below                         |
| Add bullet points                 | `- item`, nested by indentation, up to three levels       |
| Add numbered list                 | `1. item`                                                 |
| Add columns                       | `:::: columns` with one `::: column` per column, 1 to 5   |
| Add rows                          | elements written one after another are stacked top to bottom, also inside a column |
| Add code                          | fenced code block with language                           |
| Add math                          | `$...$` and `$$...$$`, editable equations in PowerPoint   |
| Add speaker notes                 | `::: notes`                                               |
| Add computed figure               | executable cell, e.g. ```` ```{python} ````               |
| Content for one format only       | `::: {.content-visible when-format="pptx"}` (`pptx`, `beamer`) |
| Coloured text                     | `[text]{.ok}` green, `[text]{.warn}` orange, `[text]{.accent}` blue |
| Picture caption                   | `![Caption](figures/picture.png)`, shown below the picture |
| Picture size                      | `![](figures/p.png){width=50%}` or `{height=30%}`: percent of the column width / content height; absolute lengths (`5cm`, `30mm`) refer to the PowerPoint page in the Fraunhofer formats |
| Lines without paragraph spacing   | line block: every line starts with `\| `                  |

### Slide layouts

Set with `## Heading {layout=...}`.

| `layout=`        | beamer                          | PPTX (layout number in the template)     |
|------------------|---------------------------------|-------------------------------------------|
| `content`        | default                         | 19, one column, small font                |
| `content-large`  | larger type                     | 20, one column, large font                |
| `image-left`     | first image left, text right    | 14, picture in the left column            |
| `image-right`    | text left, first image right    | 17, picture placeholder on the right      |
| `full-image`     | image fills the slide           | 19, picture fills the content area        |
| `quote`          | block quote                     | 24, quote / author / position             |
| `blank`          | no title                        | 19 without title                          |

Columns need no layout attribute. Two to four columns use the template layouts 14 to
16, five columns are generated as text boxes on layout 19.

### Configurations

Set in the YAML header of the deck.

| Key               | Formats        | Meaning                                              |
|-------------------|----------------|------------------------------------------------------|
| `title`, `subtitle`, `author`, `institute`, `date` | all | title slide and footer        |
| `lang`            | all            | `en` or `de`: texts, date format, PPTX template      |
| `format`          | all            | formats rendered by `tools/render.py` without `--to` |
| `fau-institute`   | `fau-beamer`   | `FAU`, `Tech`, `Nat`, `Phil`, `RW`, `Med`            |
| `iis-titlepage`   | `iis-*`        | `1` chip picture, `2` building picture               |
| `title-image`     | `iis-*`        | own title picture, path relative to the deck         |
| `closing`         | `iis-*`        | `false` removes the closing slide                    |
| `closing-text`    | `iis-*`        | text of the closing slide                            |
| `footer`          | `iis-*`        | footer text, default `© Fraunhofer IIS`              |
| `classification`  | `iis-*`        | text of the classification label in the footer, default `Restricted` |
| `shorttitle`, `shortauthor`, `shortinstitute` | beamer | short forms for the footline      |
| `toc`             | beamer         | table of contents slide                              |

## Example template

The complete example is `slides/example/deck.qmd`.

````markdown
---
title: "Example deck"
subtitle: "Markdown to beamer and PPTX"
author: "Jane Doe"
institute: "Fraunhofer IIS"
date: 2026-09-28
lang: en
format:
  fau-beamer: default
  iis-beamer: default
  iis-pptx: default
---

# Basics

## Text and lists {subtitle="Subline of the slide"}

Plain paragraph with **bold**, *italic* and `inline code`.

- First level
  - Second level
- Inline math $E = mc^2$

::: notes
Speaker notes for the first slide.
:::

## Table

| Layout  | Columns |
|---------|--------:|
| content |       1 |
| columns |     2–5 |

## Two columns

:::: columns
::: column
Left column
:::
::: column
Right column
:::
::::

## Image right {layout=image-right}

![](figures/placeholder.png)

- Text beside the image
````

## Example output

Render the example to get one file per format in `slides/example/`:

| Layout     | Format | File                          |
|------------|--------|-------------------------------|
| FAU        | PDF    | `slides/example/deck-fau.pdf`  |
| Fraunhofer | PDF    | `slides/example/deck-iis.pdf`  |
| Fraunhofer | PPTX   | `slides/example/deck-iis.pptx` |

## How it works

```
deck.qmd ── quarto ──┬─ beamer writer + template.latex + layouts.lua ── LaTeX ── PDF
                     └─ pptx: dump-ast.lua writes the AST ── md2pptx.py ── PPTX
```

- `tools/render.py`: renders one format per Quarto run and names the outputs.
- `_quarto.yml`: pre-render hook `tools/stage.py`, post-render hook `tools/postrender.py`.
- `_extensions/fau`, `_extensions/iis`: custom formats with pandoc templates.
- `_extensions/shared/layouts.lua`: slide layouts for beamer.
- `_extensions/shared/dump-ast.lua`: writes the final pandoc AST, images embedded.
- `tools/md2pptx.py`: builds the PPTX from the official template. Pandoc's own PPTX
  writer cannot address the Fraunhofer layouts.
- `_extensions/iis/layouts.yaml`: layout numbers, placeholder indices, texts, fonts.

`md2pptx.py` can be run on its own:

```sh
KEEP_AST=1 quarto render slides/example/deck.qmd --to iis-pptx
.venv/bin/python tools/md2pptx.py slides/example/deck.pptx.ast.json -o out.pptx
```

## Limits

- PowerPoint does not reflow text: `md2pptx` estimates heights, shrinks code blocks down
  to 8 pt and warns when a slide is still too full. Split such slides.
- Formulas pandoc cannot convert are inserted as pictures (display) or as TeX source
  (inline); a warning names the formula. Formulas inside table cells are always shown
  as TeX source.
- The Fraunhofer beamer theme supports 16:9 only.

## Fraunhofer: PDF and PPTX look alike

The PowerPoint template is the reference. The beamer output does not use the geometry of
the beamer theme of the institute; `_extensions/iis/template.latex` takes fonts and
colours from the theme and rebuilds the slides from the measures of the PowerPoint master:

- positions and font sizes are those of the PowerPoint page (33.867 cm x 19.05 cm),
  scaled by 0.4724 to the beamer page (16 cm x 9 cm), e.g. body text 14 pt -> 6.61 pt;
- title, chapter and closing slides follow the layouts 1, 12 and 29; their logo and
  background pictures are copied from the template to `iis-pptx/` by `tools/stage.py`;
- footer: page, date, footer text, classification label and logo at the positions of the
  master; every slide counts as a page, as in PowerPoint;
- columns, pictures and tables follow the rules of `tools/md2pptx.py`
  (`_extensions/shared/layouts.lua`).

Known differences:

- formulas: LaTeX in the PDF, native PowerPoint equations in the PPTX;
- code: different monospaced font (Latin Modern Mono, Consolas);
- text heights are estimated for PowerPoint, so stacked elements can be a few millimetres
  apart;
- the shipped `titlepage*.tex` and `thankyou.tex` of the beamer theme are not used. They
  also refer to files with wrong letter case (`IISLOgo.png`,
  `Titelseite2_Building_IIS.jpg`).
