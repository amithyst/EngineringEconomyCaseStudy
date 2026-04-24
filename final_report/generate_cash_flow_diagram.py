from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties


ROOT = Path(__file__).resolve().parent
FIGURE_DIR = ROOT / "figures"
FONT_DIR = ROOT / "fonts"
OUT = FIGURE_DIR / "cash_flow_diagram.png"


def font(name: str, size: int, weight: str = "normal") -> FontProperties:
    return FontProperties(fname=str(FONT_DIR / name), size=size, weight=weight)


FONT_CN = font("NotoSerifCJKsc-Regular.otf", 9)
FONT_CN_BOLD = font("NotoSerifCJKsc-Bold.otf", 9, "bold")
FONT_SMALL = font("NotoSerifCJKsc-Regular.otf", 7)


def draw_arrow(ax, x, y0, y1, color, width=2.6):
    ax.annotate(
        "",
        xy=(x, y1),
        xytext=(x, y0),
        arrowprops={
            "arrowstyle": "-|>",
            "lw": width,
            "color": color,
            "mutation_scale": 16,
            "shrinkA": 0,
            "shrinkB": 0,
        },
    )


def main() -> None:
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)

    fig, ax = plt.subplots(figsize=(10.6, 5.6), dpi=240)
    ax.set_xlim(-0.75, 4.75)
    ax.set_ylim(-3.05, 3.95)
    ax.axis("off")

    axis_color = "#2F3A45"
    inflow = "#2E75B6"
    outflow = "#A23E48"
    terminal = "#1F5F8B"
    grid = "#D8DEE6"

    ax.plot([-0.05, 4.25], [0, 0], color=axis_color, lw=1.5)
    for x in range(5):
        ax.plot([x, x], [-0.07, 0.07], color=axis_color, lw=1.2)
        if x == 0:
            ax.text(x - 0.12, -0.23, "t=0", ha="right", va="top", fontproperties=FONT_SMALL, color=axis_color)
        else:
            ax.text(x, -0.23, f"t={x}", ha="center", va="top", fontproperties=FONT_SMALL, color=axis_color)

    year_labels = [
        "2026-01-01\nPeriod 0",
        "2026年末",
        "2027年末",
        "2028年末",
        "2029年末",
    ]
    for x, label in enumerate(year_labels):
        if x == 0:
            ax.text(x - 0.12, -0.55, label, ha="right", va="top", fontproperties=FONT_SMALL, color="#49515A")
        else:
            ax.text(x, -0.55, label, ha="center", va="top", fontproperties=FONT_SMALL, color="#49515A")

    # Initial outflow is expressed symbolically because purchase price is solved in Step 2.
    draw_arrow(ax, 0, 0, -1.95, outflow, width=3.0)
    ax.text(
        0,
        -2.22,
        "期初投资成本\nP×1.03 + 2,100万元",
        ha="center",
        va="top",
        fontproperties=FONT_CN_BOLD,
        color=outflow,
    )

    yearly = [
        (1, 0.82, "ATCF\n1,102.66万元"),
        (2, 1.12, "ATCF\n3,324.53万元"),
        (3, 1.15, "ATCF\n3,376.84万元"),
    ]
    for x, h, label in yearly:
        draw_arrow(ax, x, 0, h, inflow)
        ax.text(x, h + 0.12, label, ha="center", va="bottom", fontproperties=FONT_CN_BOLD, color=inflow)

    draw_arrow(ax, 4, 0, 1.45, terminal, width=3.2)
    ax.text(
        4,
        1.66,
        "ATCF 3,534.27万元",
        ha="center",
        va="bottom",
        fontproperties=FONT_CN_BOLD,
        color=terminal,
    )
    ax.text(
        4,
        1.96,
        "退出价值 44,178.36万元",
        ha="center",
        va="bottom",
        fontproperties=FONT_CN_BOLD,
        color=terminal,
    )

    ax.text(
        2.1,
        0.18,
        "年末现金流入",
        ha="center",
        va="bottom",
        fontproperties=FONT_SMALL,
        color="#69717A",
    )
    ax.text(
        4.34,
        0,
        "时间",
        ha="left",
        va="center",
        fontproperties=FONT_SMALL,
        color=axis_color,
    )

    ax.text(
        -0.48,
        3.72,
        "Cash Flow Diagram from Buyer's Viewpoint",
        ha="left",
        va="top",
        fontsize=12,
        fontweight="bold",
        family="Times New Roman",
        color="#243B53",
    )
    ax.text(
        -0.48,
        3.34,
        "向下箭头表示现金流出；向上箭头表示经营现金流和期末退出现金流。MARR=12% 的现值分析在 Step 2 中进行。",
        ha="left",
        va="top",
        fontproperties=FONT_SMALL,
        color="#4B5563",
    )

    ax.hlines(-2.72, -0.48, 4.50, color=grid, lw=0.8)
    fig.tight_layout(pad=0.5)
    fig.savefig(OUT, bbox_inches="tight", facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    main()
