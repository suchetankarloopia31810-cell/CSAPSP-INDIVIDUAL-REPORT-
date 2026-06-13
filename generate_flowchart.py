"""
Generates flowchart.png — a professional flowchart for the Hospital ED
Patient Triage Management System algorithm.

Shapes used:
  • Rounded rectangle (stadium) = START / END
  • Rectangle                   = Process step
  • Diamond                     = Decision
  • Arrows                      = Control flow
"""

import matplotlib
matplotlib.use("Agg")   # non-interactive backend (no display needed)
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

# ── colour palette ──────────────────────────────────────────────────────────
C_START   = "#1B4F72"   # dark navy   – terminal nodes
C_PROC    = "#2874A6"   # mid blue    – process boxes
C_DEC     = "#C0392B"   # red         – decision diamonds
C_FLAG    = "#117A65"   # dark teal   – critical flag box
C_TEXT    = "white"
C_ARROW   = "#2C3E50"   # near-black

FIG_W, FIG_H = 10, 22
fig, ax = plt.subplots(figsize=(FIG_W, FIG_H))
ax.set_xlim(0, FIG_W)
ax.set_ylim(0, FIG_H)
ax.axis("off")
fig.patch.set_facecolor("white")

# ── helper drawers ──────────────────────────────────────────────────────────

def draw_rounded(ax, cx, cy, w, h, color, text, fontsize=8.5, radius=0.35):
    """Draw a rounded rectangle (process / terminal) centred at (cx, cy)."""
    x = cx - w / 2
    y = cy - h / 2
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle=f"round,pad=0,rounding_size={radius}",
                         linewidth=1.4, edgecolor="white",
                         facecolor=color, zorder=3)
    ax.add_patch(box)
    ax.text(cx, cy, text, ha="center", va="center",
            fontsize=fontsize, color=C_TEXT, fontweight="bold",
            wrap=True, zorder=4,
            multialignment="center")


def draw_terminal(ax, cx, cy, w, h, color, text, fontsize=9.5):
    """Draw a stadium (fully rounded) terminal node."""
    x = cx - w / 2
    y = cy - h / 2
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle=f"round,pad=0,rounding_size={h/2}",
                         linewidth=1.6, edgecolor="white",
                         facecolor=color, zorder=3)
    ax.add_patch(box)
    ax.text(cx, cy, text, ha="center", va="center",
            fontsize=fontsize, color=C_TEXT, fontweight="bold", zorder=4)


def draw_diamond(ax, cx, cy, w, h, color, text, fontsize=8):
    """Draw a decision diamond centred at (cx, cy)."""
    dx, dy = w / 2, h / 2
    xs = [cx,       cx + dx, cx,       cx - dx, cx]
    ys = [cy + dy,  cy,      cy - dy,  cy,       cy + dy]
    ax.fill(xs, ys, color=color, zorder=3)
    ax.plot(xs, ys, color="white", linewidth=1.4, zorder=4)
    ax.text(cx, cy, text, ha="center", va="center",
            fontsize=fontsize, color=C_TEXT, fontweight="bold",
            multialignment="center", zorder=5)


def arrow(ax, x1, y1, x2, y2, label="", label_side="right"):
    """Draw a directed arrow; optionally label it."""
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", color=C_ARROW,
                                lw=1.5, mutation_scale=14),
                zorder=2)
    if label:
        mx = (x1 + x2) / 2
        my = (y1 + y2) / 2
        offset = 0.18 if label_side == "right" else -0.18
        ax.text(mx + offset, my, label, ha="center", va="center",
                fontsize=7.5, color=C_ARROW, fontstyle="italic", zorder=5)


def horiz_arrow(ax, x1, y, x2, label="", label_above=True):
    """Horizontal arrow from (x1,y) to (x2,y)."""
    ax.annotate("", xy=(x2, y), xytext=(x1, y),
                arrowprops=dict(arrowstyle="-|>", color=C_ARROW,
                                lw=1.5, mutation_scale=14),
                zorder=2)
    if label:
        mx = (x1 + x2) / 2
        offset = 0.18 if label_above else -0.18
        ax.text(mx, y + offset, label, ha="center", va="center",
                fontsize=7.5, color=C_ARROW, fontstyle="italic", zorder=5)


# ── layout constants ────────────────────────────────────────────────────────
CX   = 5.0      # main column centre x
LX   = 2.2      # left branch x  (Yes on left for "valid?" decision)
RX   = 7.8      # right branch x (No on right)
W    = 3.8      # standard box width
WD   = 3.6      # diamond width
H    = 0.65     # standard box height
HD   = 0.75     # diamond height
GAP  = 1.0      # vertical gap between nodes

# Y positions (top → bottom)
nodes = {
    "start":    21.4,
    "load":     20.2,
    "valid":    18.9,   # diamond
    "skip":     18.9,   # side box (No → skip)
    "store":    17.7,
    "init":     16.6,
    "timer_s":  15.5,
    "base":     14.3,   # diamond: len > 1 ?
    "divide":   13.1,
    "rec_l":    11.9,
    "rec_r":    10.8,
    "merge":     9.6,
    "timer_e":   8.5,
    "display":   7.4,
    "crit":      6.2,   # diamond: score ≤ 2 ?
    "flag":      5.1,   # Yes → flag box
    "stats":     3.9,
    "end":       2.8,
}

# ── draw nodes ──────────────────────────────────────────────────────────────

draw_terminal(ax, CX, nodes["start"], 2.6, 0.60, C_START, "START")

draw_rounded(ax, CX, nodes["load"],    W, H, C_PROC,
             "Load patient records\n(text file or generate test data)")

draw_diamond(ax, CX, nodes["valid"],   WD, HD, C_DEC,
             "Valid record?\n(string name + int score 1–5)")

# No branch (right)
draw_rounded(ax, RX, nodes["skip"],   2.5, H, "#7D3C98",
             "Skip / log\nmalformed record")

draw_rounded(ax, CX, nodes["store"],   W, H, C_PROC,
             "Store in list of dictionaries\n{'name': str, 'score': int}")

draw_rounded(ax, CX, nodes["init"],    W, H, C_PROC,
             "Initialise OpCounter\n(comparisons = swaps = arithmetic = 0)")

draw_rounded(ax, CX, nodes["timer_s"], W, H, C_PROC,
             "Start high-resolution timer\n(time.perf_counter)")

draw_diamond(ax, CX, nodes["base"],    WD, HD, C_DEC,
             "len(data) > 1?")

# No (base case) → right side text
ax.text(RX - 0.3, nodes["base"], "No →\nBase case\n(already sorted)",
        ha="left", va="center", fontsize=7.5, color=C_ARROW,
        style="italic")

draw_rounded(ax, CX, nodes["divide"],  W, H, C_PROC,
             "Divide: split into\nLEFT = data[:mid]  |  RIGHT = data[mid:]")

draw_rounded(ax, CX, nodes["rec_l"],   W, H, C_PROC,
             "Recursively sort\nLEFT half")

draw_rounded(ax, CX, nodes["rec_r"],   W, H, C_PROC,
             "Recursively sort\nRIGHT half")

draw_rounded(ax, CX, nodes["merge"],   W, H, C_PROC,
             "Merge: compare scores element-by-element\n→ rebuild sorted list; increment counters")

draw_rounded(ax, CX, nodes["timer_e"], W, H, C_PROC,
             "Stop timer\nelapsed_ms = (end − start) × 1000")

draw_rounded(ax, CX, nodes["display"], W, H, C_PROC,
             "Display sorted patient queue\n(ascending triage score)")

draw_diamond(ax, CX, nodes["crit"],    WD, HD, C_DEC,
             "Patient score ≤ 2?")

# Yes → left flag box
draw_rounded(ax, LX, nodes["flag"],   2.6, H, C_FLAG,
             "Flag as CRITICAL\n(Immediate / Very Urgent)")

draw_rounded(ax, CX, nodes["stats"],   W, H, C_PROC,
             "Print summary:\nrecord count · elapsed time · total ops")

draw_terminal(ax, CX, nodes["end"],   2.6, 0.60, C_START, "END")

# ── draw arrows (main vertical flow) ────────────────────────────────────────
flow = [
    ("start",   "load"),
    ("load",    "valid"),
    # valid → store (Yes, straight down)
    ("store",   "init"),
    ("init",    "timer_s"),
    ("timer_s", "base"),
    # base  → divide (Yes)
    ("divide",  "rec_l"),
    ("rec_l",   "rec_r"),
    ("rec_r",   "merge"),
    ("merge",   "timer_e"),
    ("timer_e", "display"),
    ("display", "crit"),
    # crit → stats (No)
    ("stats",   "end"),
]

def btm(node_key, box_h=H):
    """Bottom y of a box."""
    return nodes[node_key] - box_h / 2

def top(node_key, box_h=H):
    """Top y of a box."""
    return nodes[node_key] + box_h / 2

# straight vertical arrows along main spine
for a_key, b_key in flow:
    if a_key in ("start", "end"):
        ya = nodes[a_key] - 0.30
    else:
        ya = btm(a_key)
    if b_key in ("start", "end"):
        yb = nodes[b_key] + 0.30
    else:
        yb = top(b_key)
    arrow(ax, CX, ya, CX, yb)

# "valid?" → store  (Yes, down)
arrow(ax, CX, btm("valid", HD), CX, top("store"), label="Yes")

# "valid?" → skip (No, right horizontal)
horiz_arrow(ax, CX + WD/2, nodes["valid"], RX - 1.25, label="No", label_above=True)

# "base?" → divide (Yes, down)
arrow(ax, CX, btm("base", HD), CX, top("divide"), label="Yes")

# "crit?" → flag (Yes, left horizontal then down)
horiz_arrow(ax, CX - WD/2, nodes["crit"], LX + 1.3, label="Yes", label_above=True)
arrow(ax, LX, nodes["crit"] - HD/2, LX, top("flag", H))

# "crit?" → stats (No, down)
arrow(ax, CX, btm("crit", HD), CX, top("stats"), label="No")

# flag → stats (right then down — re-join main flow)
ax.annotate("", xy=(CX - W/2, nodes["stats"]),
            xytext=(LX + 1.3, nodes["flag"] - H/2),
            arrowprops=dict(arrowstyle="-|>", color=C_ARROW,
                            lw=1.5, mutation_scale=14,
                            connectionstyle="angle,angleA=0,angleB=90,rad=0"),
            zorder=2)

# ── title ────────────────────────────────────────────────────────────────────
ax.text(CX, FIG_H - 0.25,
        "Figure 1: Hospital ED Patient Triage System — Algorithm Flowchart",
        ha="center", va="top", fontsize=10, fontweight="bold", color="#1B2631")

# ── legend ───────────────────────────────────────────────────────────────────
legend_items = [
    mpatches.Patch(facecolor=C_START,   edgecolor="grey", label="Terminal (Start / End)"),
    mpatches.Patch(facecolor=C_PROC,    edgecolor="grey", label="Process Step"),
    mpatches.Patch(facecolor=C_DEC,     edgecolor="grey", label="Decision"),
    mpatches.Patch(facecolor=C_FLAG,    edgecolor="grey", label="Critical Flag Action"),
    mpatches.Patch(facecolor="#7D3C98", edgecolor="grey", label="Error Handling"),
]
ax.legend(handles=legend_items,
          loc="lower center", bbox_to_anchor=(0.5, 0.01),
          fontsize=7.5, framealpha=0.9, ncol=3,
          title="Shape Legend", title_fontsize=8)

plt.tight_layout(pad=0.4)
plt.savefig("flowchart.png", dpi=180, bbox_inches="tight",
            facecolor="white", edgecolor="none")
plt.close()
print("flowchart.png saved successfully.")
