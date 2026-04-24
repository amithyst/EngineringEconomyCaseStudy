#!/usr/bin/env python3

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap


REPO_ROOT = Path(__file__).resolve().parents[2]
RESULTS_PATH = REPO_ROOT / "outputs" / "case_study_results.json"
FIGURE_DIR = REPO_ROOT / "outputs" / "submission" / "figures"


def setup_fonts() -> None:
    plt.rcParams["font.sans-serif"] = [
        "Noto Sans CJK SC",
        "Source Han Sans SC",
        "Microsoft YaHei",
        "SimHei",
        "DejaVu Sans",
    ]
    plt.rcParams["axes.unicode_minus"] = False


def make_cash_flow_chart(data: dict) -> None:
    annual = data["annual_results"]
    years = [str(item["year"]) for item in annual]
    cashflows = [item["after_tax_cash_flow"] / 1e6 for item in annual]
    purchase_price = data["summary"]["recommended_purchase_price"] / 1e8

    fig, ax = plt.subplots(figsize=(7.2, 3.6), constrained_layout=True)
    colors = ["#C9D6DF", "#4F6D7A", "#5B8E7D", "#C67B5C"]
    bars = ax.bar(years, cashflows, color=colors, width=0.58)
    ax.set_title("Tax-Free Operating Cash Flow by Year", fontsize=13, fontweight="bold", loc="left")
    ax.set_ylabel("Million RMB")
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)
    for bar, value in zip(bars, cashflows):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.8, f"{value:.1f}", ha="center", va="bottom", fontsize=10, fontweight="bold")

    ax2 = ax.twinx()
    occ = [item["occupancy"] * 100 for item in annual]
    ax2.plot(years, occ, color="#A23E48", marker="o", linewidth=2.0)
    ax2.set_ylabel("Occupancy (%)", color="#A23E48")
    ax2.tick_params(axis="y", colors="#A23E48")
    ax2.set_ylim(70, 100)
    ax.text(
        0.03,
        0.88,
        f"Purchase price cap: {purchase_price:.2f} bn RMB",
        transform=ax.transAxes,
        ha="left",
        va="top",
        fontsize=9.5,
        color="#4B5563",
        bbox={"boxstyle": "round,pad=0.25", "facecolor": "#F4F6F8", "edgecolor": "#D9E2EC"},
    )
    fig.savefig(FIGURE_DIR / "cash_flow_chart.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def make_income_mix_chart(data: dict) -> None:
    annual = data["annual_results"]
    years = [str(item["year"]) for item in annual]
    rent = [item["rent_income"] / 1e6 for item in annual]
    parking = [item["parking_income"] / 1e6 for item in annual]
    property_fee = [item["property_fee_income"] / 1e6 for item in annual]

    fig, ax = plt.subplots(figsize=(7.0, 3.5), constrained_layout=True)
    ax.bar(years, rent, label="Rent", color="#406E8E")
    ax.bar(years, parking, bottom=rent, label="Parking", color="#8E5572")
    bottoms = [a + b for a, b in zip(rent, parking)]
    ax.bar(years, property_fee, bottom=bottoms, label="Property Fee", color="#7C9D54")
    ax.set_title("Revenue Mix by Year", fontsize=13, fontweight="bold", loc="left")
    ax.set_ylabel("Million RMB")
    ax.legend(frameon=False, ncol=3, loc="upper left")
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)
    fig.savefig(FIGURE_DIR / "income_mix_chart.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def make_quarterly_chart(data: dict) -> None:
    q = data["quarterly_2026"]
    labels = [row["quarter"] for row in q]
    income = [row["total_income"] / 1e6 for row in q]
    opex = [row["allocated_opex"] / 1e6 for row in q]
    pretax = [row["pre_tax_operating_cash_flow"] / 1e6 for row in q]

    fig, ax = plt.subplots(figsize=(7.0, 3.4), constrained_layout=True)
    x = range(len(labels))
    ax.bar([i - 0.18 for i in x], income, width=0.36, label="Operating income", color="#6B8E9F")
    ax.bar([i + 0.18 for i in x], opex, width=0.36, label="Allocated opex", color="#D9A066")
    ax.plot(list(x), pretax, color="#A23E48", marker="o", linewidth=2.2, label="Pre-tax operating CF")
    ax.axhline(0, color="#666", linewidth=0.8)
    ax.set_xticks(list(x), labels)
    ax.set_ylabel("Million RMB")
    ax.set_title("2026 Quarterly Operating Ramp-Up", fontsize=13, fontweight="bold", loc="left")
    ax.legend(frameon=False, ncol=3, loc="upper left")
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)
    fig.savefig(FIGURE_DIR / "quarterly_ramp_chart.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def make_sensitivity_chart(data: dict) -> None:
    sensitivity = data["sensitivity"]
    growth_keys = ["rent_growth_neg03pct", "rent_growth_00pct", "rent_growth_03pct", "rent_growth_05pct", "rent_growth_08pct"]
    growth_labels = ["-3%", "0%", "3%", "5%", "8%"]
    occ_labels = [f"{int(row['stabilized_occupancy'] * 100)}%" for row in sensitivity]
    matrix = [[row[key] / 1e8 for key in growth_keys] for row in sensitivity]
    cmap = LinearSegmentedColormap.from_list("case_study", ["#E9F1F7", "#A8C5D6", "#6E8FA8", "#34526F"])
    fig, ax = plt.subplots(figsize=(7.4, 4.8), constrained_layout=True)
    im = ax.imshow(matrix, cmap=cmap, aspect="auto")
    ax.set_title("Sensitivity of Purchase Price", fontsize=13, fontweight="bold", loc="left")
    ax.set_xticks(range(len(growth_labels)), growth_labels)
    ax.set_yticks(range(len(occ_labels)), occ_labels)
    ax.set_xlabel("2029 Rent Growth")
    ax.set_ylabel("Stabilized Occupancy")
    for i, row in enumerate(matrix):
        for j, value in enumerate(row):
            ax.text(j, i, f"{value:.2f}", ha="center", va="center", color="white", fontsize=9.5, fontweight="bold")
    cbar = fig.colorbar(im, ax=ax, shrink=0.9)
    cbar.set_label("Purchase Price (bn RMB)")
    fig.savefig(FIGURE_DIR / "sensitivity_heatmap.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    setup_fonts()
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    with RESULTS_PATH.open(encoding="utf-8") as fp:
        data = json.load(fp)
    make_cash_flow_chart(data)
    make_income_mix_chart(data)
    make_quarterly_chart(data)
    make_sensitivity_chart(data)


if __name__ == "__main__":
    main()
