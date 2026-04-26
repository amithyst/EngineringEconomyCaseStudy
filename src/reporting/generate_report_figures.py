#!/usr/bin/env python3

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.font_manager import FontProperties


REPO_ROOT = Path(__file__).resolve().parents[2]
RESULTS_PATH = REPO_ROOT / "outputs" / "case_study_results.json"
FIGURE_DIR = REPO_ROOT / "final_report" / "figures"
FONT_DIR = REPO_ROOT / "final_report" / "fonts"

# Unified colour palette
_P = {
    "bar1": "#B8D0E0",
    "bar2": "#4A7FA5",
    "bar3": "#5B8E7D",
    "bar4": "#C07050",
    "occ": "#A23E48",
    "rent": "#3A6E9E",
    "parking": "#9E6080",
    "prop_fee": "#4E7E5A",
    "grid": "#E4EAF0",
    "text": "#1A1A1A",
    "ann_bg": "#F4F6F8",
    "ann_edge": "#D0DCE8",
}

TITLE_SZ = 14
LABEL_SZ = 12
TICK_SZ = 10
DATA_SZ = 10
ANNOT_SZ = 10
LEGEND_SZ = 10


FONT_CN = FontProperties(fname=str(FONT_DIR / "NotoSerifCJKsc-Regular.otf"), size=LABEL_SZ)
FONT_CN_TITLE = FontProperties(fname=str(FONT_DIR / "NotoSerifCJKsc-Bold.otf"), size=TITLE_SZ, weight="bold")
FONT_CN_SMALL = FontProperties(fname=str(FONT_DIR / "NotoSerifCJKsc-Regular.otf"), size=LEGEND_SZ)


def setup_style() -> None:
    plt.rcParams.update(
        {
            "font.family": "serif",
            "font.serif": ["Times New Roman", "DejaVu Serif"],
            "font.sans-serif": [
                "Noto Serif SC",
                "Noto Sans SC",
                "Microsoft YaHei",
                "SimHei",
                "DejaVu Sans",
            ],
            "axes.unicode_minus": False,
            "axes.spines.top": False,
            "axes.spines.right": False,
        }
    )


def _spine_light(ax, color="#C8D8E4") -> None:
    for sp in ["left", "bottom"]:
        ax.spines[sp].set_color(color)


def make_cash_flow_chart(data: dict) -> None:
    annual = data["annual_results"]
    years = [str(item["year"]) for item in annual]
    cashflows = [item["after_tax_cash_flow"] / 1e6 for item in annual]
    purchase_price = data["summary"]["recommended_purchase_price"] / 1e8

    fig, ax = plt.subplots(figsize=(8.5, 4.2), constrained_layout=True)
    bar_colors = [_P["bar1"], _P["bar2"], _P["bar3"], _P["bar4"]]
    bars = ax.bar(years, cashflows, color=bar_colors, width=0.55, zorder=3)

    ax.set_title(
        "年度税后经营现金流",
        fontproperties=FONT_CN_TITLE, loc="left", pad=8, color=_P["text"],
    )
    ax.set_ylabel("金额（百万元）", fontproperties=FONT_CN, color=_P["text"])
    ax.tick_params(labelsize=TICK_SZ)
    ax.grid(axis="y", linestyle="--", alpha=0.45, color=_P["grid"], zorder=0)
    ax.set_axisbelow(True)
    _spine_light(ax)
    ax.set_ylim(0, max(cashflows) * 1.24)

    for bar, value in zip(bars, cashflows):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.5,
            f"{value:.1f}",
            ha="center", va="bottom",
            fontsize=DATA_SZ, fontweight="bold", color=_P["text"],
        )

    ax2 = ax.twinx()
    occ = [item["occupancy"] * 100 for item in annual]
    ax2.plot(years, occ, color=_P["occ"], marker="o", linewidth=2.2, markersize=5, zorder=4)
    ax2.set_ylabel("出租率（%）", fontproperties=FONT_CN, color=_P["occ"])
    ax2.tick_params(axis="y", colors=_P["occ"], labelsize=TICK_SZ)
    ax2.set_ylim(60, 100)
    ax2.spines["right"].set_visible(True)
    ax2.spines["right"].set_color(_P["occ"])
    ax2.spines["top"].set_visible(False)

    # Annotation in upper-left, visually below the chart title
    ax.text(
        0.02, 0.97,
        f"收购对价上限：{purchase_price:.2f} 亿元",
        transform=ax.transAxes,
        ha="left", va="top",
        fontproperties=FONT_CN_SMALL, color=_P["text"],
        bbox={"boxstyle": "round,pad=0.3", "facecolor": _P["ann_bg"], "edgecolor": _P["ann_edge"]},
    )

    fig.savefig(FIGURE_DIR / "cash_flow_chart.png", dpi=240, bbox_inches="tight")
    plt.close(fig)


def make_income_mix_chart(data: dict) -> None:
    annual = data["annual_results"]
    years = [str(item["year"]) for item in annual]
    rent = [item["rent_income"] / 1e6 for item in annual]
    parking = [item["parking_income"] / 1e6 for item in annual]
    property_fee = [item["property_fee_income"] / 1e6 for item in annual]

    fig, ax = plt.subplots(figsize=(7.5, 4.0), constrained_layout=True)
    ax.bar(years, rent, label="租金", color=_P["rent"], zorder=3)
    ax.bar(years, parking, bottom=rent, label="停车费", color=_P["parking"], zorder=3)
    bottoms = [a + b for a, b in zip(rent, parking)]
    ax.bar(years, property_fee, bottom=bottoms, label="物业费", color=_P["prop_fee"], zorder=3)

    ax.set_title("年度收入结构", fontproperties=FONT_CN_TITLE, loc="left", pad=8, color=_P["text"])
    ax.set_ylabel("金额（百万元）", fontproperties=FONT_CN, color=_P["text"])
    ax.tick_params(labelsize=TICK_SZ)
    ax.legend(frameon=False, ncol=3, loc="upper left", prop=FONT_CN_SMALL)
    ax.grid(axis="y", linestyle="--", alpha=0.45, color=_P["grid"], zorder=0)
    ax.set_axisbelow(True)
    _spine_light(ax)

    totals = [a + b + c for a, b, c in zip(rent, parking, property_fee)]
    ax.set_ylim(0, max(totals) * 1.18)
    for i, total in enumerate(totals):
        ax.text(i, total + 0.8, f"{total:.1f}", ha="center", va="bottom", fontsize=DATA_SZ, color=_P["text"])

    fig.savefig(FIGURE_DIR / "income_mix_chart.png", dpi=240, bbox_inches="tight")
    plt.close(fig)


def make_sensitivity_chart(data: dict) -> None:
    sensitivity = data["sensitivity"]
    growth_keys = [
        "rent_growth_neg03pct", "rent_growth_00pct",
        "rent_growth_03pct", "rent_growth_05pct", "rent_growth_08pct",
    ]
    growth_labels = ["-3%", "0%", "3%", "5%", "8%"]
    occ_labels = [f"{int(row['stabilized_occupancy'] * 100)}%" for row in sensitivity]
    matrix = [[row[key] / 1e8 for key in growth_keys] for row in sensitivity]

    cmap = LinearSegmentedColormap.from_list(
        "case_study", ["#E2EEF6", "#A6C8DC", "#5E97B8", "#2E6087"]
    )

    fig, ax = plt.subplots(figsize=(8.0, 5.0), constrained_layout=True)
    flat = [v for row in matrix for v in row]
    vmin, vmax = min(flat), max(flat)
    im = ax.imshow(matrix, cmap=cmap, aspect="auto", vmin=vmin, vmax=vmax)

    ax.set_title(
        "收购价格敏感性分析（亿元）",
        fontproperties=FONT_CN_TITLE, loc="left", pad=8, color=_P["text"],
    )
    ax.set_xticks(range(len(growth_labels)))
    ax.set_xticklabels(growth_labels, fontsize=TICK_SZ)
    ax.set_yticks(range(len(occ_labels)))
    ax.set_yticklabels(occ_labels, fontsize=TICK_SZ)
    ax.set_xlabel("2029 年租金增长率", fontproperties=FONT_CN, color=_P["text"])
    ax.set_ylabel("稳定期出租率", fontproperties=FONT_CN, color=_P["text"])

    norm_range = vmax - vmin
    for i, row in enumerate(matrix):
        for j, value in enumerate(row):
            brightness = (value - vmin) / norm_range if norm_range > 0 else 0.5
            txt_color = "white" if brightness > 0.52 else "#1A2633"
            ax.text(
                j, i, f"{value:.2f}",
                ha="center", va="center",
                color=txt_color, fontsize=DATA_SZ + 0.5, fontweight="bold",
            )

    cbar = fig.colorbar(im, ax=ax, shrink=0.85, pad=0.02)
    cbar.set_label("收购价格（亿元）", fontproperties=FONT_CN_SMALL, color=_P["text"])
    cbar.ax.tick_params(labelsize=TICK_SZ - 1)

    fig.savefig(FIGURE_DIR / "sensitivity_heatmap.png", dpi=240, bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    setup_style()
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    with RESULTS_PATH.open(encoding="utf-8") as fp:
        data = json.load(fp)
    make_cash_flow_chart(data)
    make_income_mix_chart(data)
    make_sensitivity_chart(data)


if __name__ == "__main__":
    main()
