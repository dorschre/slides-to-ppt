#!/usr/bin/env python3
"""Draws figures/pipeline.png: the CartoGraph pipeline.

    .venv/bin/python slides/cartograph/figures/make_pipeline.py
"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyArrowPatch, Rectangle

HERE = Path(__file__).resolve().parent
FONTS = HERE.parents[2] / "assets/fraunhofer/beamer theme IIS V1_05_2026/theme"

BLUE, GREEN, GREY, EDGE = "#005B7F", "#179C7D", "#F2F2F2", "#A6BBC8"

family = "DejaVu Sans"
light = FONTS / "Frutiger_ftl_ltc_.ttf"
if light.is_file():
    font_manager.fontManager.addfont(str(light))
    family = font_manager.FontProperties(fname=str(light)).get_name()
plt.rcParams.update({"font.family": family, "font.weight": 300, "mathtext.fontset": "dejavusans"})

# boxes: name -> (x, y, width, height, label, highlighted); units: cm of the source slide
BOXES = {
    "source": (0.0, 2.7, 4.0, 2.4, "Source graph\n$G_I$", False),
    "candidate": (6.5, 2.7, 4.0, 2.4, "Candidate\nelements\n$G_{E,i}$", False),
    "resolved": (13.3, 2.7, 4.0, 2.4, "Resolved\ncontent\n$G_{F,i}$", True),
    "image": (19.5, 5.0, 3.7, 1.7, "Map image\n$M_i$", False),
    "answer": (19.5, 0.7, 3.7, 2.1, "Reference\nanswer\n$A_{i,j}$", False),
}

fig, ax = plt.subplots(figsize=(23.4 / 2.54, 7.0 / 2.54), dpi=300)
ax.set_xlim(-0.1, 23.3)
ax.set_ylim(-0.2, 6.8)
ax.set_aspect("equal")
ax.axis("off")

for x, y, w, h, label, highlighted in BOXES.values():
    ax.add_patch(Rectangle((x, y), w, h, facecolor="#D9EEE8" if highlighted else GREY,
                           edgecolor=GREEN if highlighted else EDGE, linewidth=1.0))
    ax.text(x + w / 2, y + h / 2, label, ha="center", va="center", fontsize=10.5, linespacing=1.25)


def arrow(start, end, label, dx=0.0, dy=0.0, rotation=0.0):
    ax.add_patch(FancyArrowPatch(start, end, arrowstyle="-|>", mutation_scale=9, color=BLUE,
                                 linewidth=1.0, shrinkA=0, shrinkB=0))
    ax.text((start[0] + end[0]) / 2 + dx, (start[1] + end[1]) / 2 + dy, label, color=BLUE,
            ha="center", va="center", fontsize=8.5, rotation=rotation, linespacing=1.2)


arrow((4.0, 3.9), (6.5, 3.9), "Apply\nrules", dy=0.75)
arrow((10.5, 3.9), (13.3, 3.9), "Resolve\nelements", dy=0.75)
arrow((17.3, 3.9), (19.5, 5.8), "Render", dx=-0.45, dy=0.35, rotation=41)
arrow((17.3, 3.9), (19.5, 1.9), "Query $q_j$", dx=-0.45, dy=-0.4, rotation=-42)
ax.text(2.0, 1.9, "Lifted from\nsource data $D$", color=BLUE, ha="center", va="center",
        fontsize=8.5, linespacing=1.2)

fig.savefig(HERE / "pipeline.png", bbox_inches="tight", pad_inches=0.02)
print(HERE / "pipeline.png")
