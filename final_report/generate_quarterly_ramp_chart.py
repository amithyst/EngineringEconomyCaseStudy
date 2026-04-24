from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties


ROOT = Path(__file__).resolve().parent
FIGURE_DIR = ROOT / "figures"
FONT_DIR = ROOT / "fonts"
OUT = FIGURE_DIR / "quarterly_ramp_chart.png"


def font(name: str, size: int, weight: str = "normal") -> FontProperties:
    return FontProperties(fname=str(FONT_DIR / name), size=size, weight=weight)


FONT = font("NotoSerifCJKsc-Regular.otf", 11)
FONT_BOLD = font("NotoSerifCJKsc-Bold.otf", 11, "bold")
FONT_SMALL = font("NotoSerifCJKsc-Regular.otf", 9)
FONT_TITLE = font("NotoSerifCJKsc-Bold.otf", 15, "bold")


def main() -> None:
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)

    quarters = ["2026Q1", "2026Q2", "2026Q3", "2026Q4"]
    stage_labels = ["改造准备", "免租运营", "正式收租", "正式收租"]
    total_income = [0.00, 232.62, 1382.46, 1382.46]
    allocated_opex = [0.00, 509.11, 509.11, 509.11]
    pretax_cf = [0.00, -276.49, 873.35, 873.35]

    x = list(range(len(quarters)))
    width = 0.30

    blue = "#2E75B6"
    blue_light = "#AFC8DD"
    gray = "#8796A5"
    red = "#A23E48"
    axis = "#2F3A45"
    grid = "#D8DEE6"

    fig, ax = plt.subplots(figsize=(9.2, 4.8), dpi=240)
    ax.set_facecolor("white")

    ax.bar([i - width / 2 for i in x], total_income, width=width, label="经营收入", color=blue_light, edgecolor="white", linewidth=0.8)
    ax.bar([i + width / 2 for i in x], allocated_opex, width=width, label="O&M 成本", color=gray, edgecolor="white", linewidth=0.8)
    ax.plot(x, pretax_cf, color=red, marker="o", markersize=5.0, linewidth=2.2, label="税前经营现金流")

    for i, (inc, cost, cf) in enumerate(zip(total_income, allocated_opex, pretax_cf)):
        if inc > 0:
            ax.text(i - width / 2, inc + 40, f"{inc:.0f}", ha="center", va="bottom", fontproperties=FONT_SMALL, color=blue)
        if cost > 0:
            ax.text(i + width / 2, cost + 40, f"{cost:.0f}", ha="center", va="bottom", fontproperties=FONT_SMALL, color=axis)
        offset = 55 if cf >= 0 else -70
        va = "bottom" if cf >= 0 else "top"
        ax.text(i, cf + offset, f"{cf:.0f}", ha="center", va=va, fontproperties=FONT_BOLD, color=red)

    ax.annotate(
        "Q3 起租金计收，\n季度现金流转正",
        xy=(2, pretax_cf[2]),
        xytext=(2.35, 1120),
        textcoords="data",
        arrowprops={"arrowstyle": "->", "color": red, "lw": 1.0},
        fontproperties=FONT_SMALL,
        color=axis,
        ha="left",
        va="center",
    )

    ax.set_xticks(x)
    ax.set_xticklabels([])
    for i, (quarter, stage) in enumerate(zip(quarters, stage_labels)):
        label_box = {"facecolor": "white", "edgecolor": "none", "alpha": 1.0, "pad": 1.4}
        ax.text(i, -58, quarter, ha="center", va="top", fontproperties=FONT_SMALL, color=axis, bbox=label_box, zorder=6)
        ax.text(i, -138, stage, ha="center", va="top", fontproperties=FONT_SMALL, color=axis, bbox=label_box, zorder=6)
    ax.set_ylabel("金额（万元）", fontproperties=FONT, color=axis)
    ax.set_ylim(-520, 1580)
    ax.set_xlim(-0.55, 3.55)
    ax.grid(axis="y", linestyle="--", linewidth=0.8, color=grid, alpha=0.75)
    ax.set_axisbelow(True)

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color(axis)
    ax.spines["bottom"].set_position(("data", 0))
    ax.spines["bottom"].set_color(axis)
    ax.spines["bottom"].set_linewidth(1.0)
    ax.tick_params(axis="x", colors=axis, pad=5, length=4)

    handles, labels = ax.get_legend_handles_labels()
    legend = fig.legend(
        handles,
        labels,
        loc="upper center",
        bbox_to_anchor=(0.5, 0.845),
        ncol=3,
        frameon=False,
        prop=FONT_SMALL,
        handlelength=1.8,
        columnspacing=1.8,
    )
    for text in legend.get_texts():
        text.set_color(axis)

    fig.suptitle(
        "2026 年季度经营现金流改善",
        x=0.5,
        y=0.965,
        ha="center",
        va="top",
        fontproperties=FONT_TITLE,
        color=axis,
    )
    fig.text(
        0.5,
        0.895,
        "单位：万元；税前经营现金流 = 经营收入 - 当季分摊 O&M 成本",
        ha="center",
        va="top",
        fontproperties=FONT_SMALL,
        color="#5D6670",
    )

    fig.subplots_adjust(left=0.09, right=0.985, bottom=0.14, top=0.72)
    fig.savefig(OUT, bbox_inches="tight", facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    main()
