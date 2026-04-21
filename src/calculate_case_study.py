#!/usr/bin/env python3

import csv
import json
from dataclasses import dataclass
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = REPO_ROOT / "outputs"


@dataclass(frozen=True)
class YearResult:
    year: int
    occupancy: float
    rent_growth_factor: float
    rent_income: float
    parking_income: float
    property_fee_income: float
    total_income: float
    personnel_cost: float
    marketing_cost: float
    maint_repair_cost: float
    property_mgmt_cost: float
    energy_cost: float
    admin_cost: float
    capital_reno_cost: float
    total_opex: float
    pre_tax_profit: float
    tax: float
    after_tax_cash_flow: float


FLOOR_RENTABLE_AREA = {
    "6F": 141.70,
    "5F": 391.45,
    "4F": 10535.92,
    "3F": 13052.43,
    "2F": 13212.02,
    "1F": 13128.34,
    "B1": 11751.42,
    "B2": 416.48,
}

FLOOR_RENT_PER_DAY = {
    "6F": 1.5,
    "5F": 1.5,
    "4F": 2.5,
    "3F": 2.5,
    "2F": 2.5,
    "1F": 3.0,
    "B1": 2.0,
    "B2": 1.0,
}

BUILDING_AREA = 107039.13
RENTABLE_AREA = sum(FLOOR_RENTABLE_AREA.values())
PARKING_SPACES = 1005

STAMP_DUTY_RATE = 0.03
OPENING_COST = 10_000_000.0
OTHER_TRANSACTION_COST = 11_000_000.0
INCOME_TAX_RATE = 0.25
MARR = 0.12
EXIT_CAP_RATE = 0.08

OPERATING_DAYS = {
    2026: 275,
    2027: 365,
    2028: 366,
    2029: 365,
}

RENT_COLLECTION_DAYS = {
    2026: 184,
    2027: 365,
    2028: 366,
    2029: 365,
}

PROPERTY_FEE_MONTHS = {
    2026: 9,
    2027: 12,
    2028: 12,
    2029: 12,
}

PARKING_USAGE = {
    2026: 0.50,
    2027: 0.60,
    2028: 0.70,
    2029: 0.80,
}

BASE_OCCUPANCY = {
    2026: 0.80,
    2027: 0.95,
    2028: 0.95,
    2029: 0.95,
}

BASE_RENT_GROWTH_FACTOR = {
    2026: 1.00,
    2027: 1.00,
    2028: 1.00,
    2029: 1.03,
}

BASE_OPEX = {
    "personnel": 5_000_000.0,
    "marketing": 5_000_000.0,
    "maint_repair": 15.0 * BUILDING_AREA,
    "property_mgmt": 30.0 * BUILDING_AREA,
    "energy": 40.0 * BUILDING_AREA,
    "admin": 500_000.0,
}

OPEX_GROWTH = {
    "personnel": {2027: 0.00, 2028: 0.03, 2029: 0.00},
    "marketing": {2027: -0.05, 2028: -0.05, 2029: -0.05},
    "maint_repair": {2027: 0.03, 2028: 0.03, 2029: 0.03},
    "property_mgmt": {2027: 0.03, 2028: 0.03, 2029: 0.03},
    "energy": {2027: 0.01, 2028: 0.01, 2029: 0.01},
    "admin": {2027: 0.03, 2028: 0.03, 2029: 0.03},
}


def annual_rent_income(year: int, occupancy: float, rent_growth_factor: float) -> float:
    return sum(
        FLOOR_RENTABLE_AREA[floor]
        * FLOOR_RENT_PER_DAY[floor]
        * RENT_COLLECTION_DAYS[year]
        * occupancy
        * rent_growth_factor
        for floor in FLOOR_RENTABLE_AREA
    )


def annual_parking_income(year: int) -> float:
    return PARKING_SPACES * 3.0 * 6.0 * PARKING_USAGE[year] * OPERATING_DAYS[year]


def annual_property_fee_income(year: int, occupancy: float) -> float:
    return RENTABLE_AREA * occupancy * 10.0 * PROPERTY_FEE_MONTHS[year]


def build_opex_schedule() -> dict[int, dict[str, float]]:
    schedule = {2026: {}}
    for name, value in BASE_OPEX.items():
        schedule[2026][name] = value * 9.0 / 12.0

    for year in (2027, 2028, 2029):
        schedule[year] = {}
        for name, base_value in BASE_OPEX.items():
            previous = base_value if year == 2027 else schedule[year - 1][name]
            schedule[year][name] = previous * (1.0 + OPEX_GROWTH[name][year])
    return schedule


def build_year_results(
    occupancy: dict[int, float],
    rent_growth_factor: dict[int, float],
) -> list[YearResult]:
    opex_schedule = build_opex_schedule()
    results = []

    for year in (2026, 2027, 2028, 2029):
        rent_income = annual_rent_income(year, occupancy[year], rent_growth_factor[year])
        parking_income = annual_parking_income(year)
        property_fee_income = annual_property_fee_income(year, occupancy[year])
        capital_reno_cost = 0.025 * rent_income
        total_opex = sum(opex_schedule[year].values()) + capital_reno_cost
        total_income = rent_income + parking_income + property_fee_income
        pre_tax_profit = total_income - total_opex
        tax = max(pre_tax_profit, 0.0) * INCOME_TAX_RATE
        after_tax_cash_flow = pre_tax_profit - tax

        results.append(
            YearResult(
                year=year,
                occupancy=occupancy[year],
                rent_growth_factor=rent_growth_factor[year],
                rent_income=rent_income,
                parking_income=parking_income,
                property_fee_income=property_fee_income,
                total_income=total_income,
                personnel_cost=opex_schedule[year]["personnel"],
                marketing_cost=opex_schedule[year]["marketing"],
                maint_repair_cost=opex_schedule[year]["maint_repair"],
                property_mgmt_cost=opex_schedule[year]["property_mgmt"],
                energy_cost=opex_schedule[year]["energy"],
                admin_cost=opex_schedule[year]["admin"],
                capital_reno_cost=capital_reno_cost,
                total_opex=total_opex,
                pre_tax_profit=pre_tax_profit,
                tax=tax,
                after_tax_cash_flow=after_tax_cash_flow,
            )
        )

    return results


def solve_purchase_price(results: list[YearResult]) -> tuple[float, float, float]:
    discounted_operating = sum(
        result.after_tax_cash_flow / ((1.0 + MARR) ** (result.year - 2025))
        for result in results
    )
    terminal_value = results[-1].after_tax_cash_flow / EXIT_CAP_RATE
    terminal_pv = terminal_value / ((1.0 + MARR) ** 4)
    purchase_price = (
        discounted_operating + terminal_pv - OPENING_COST - OTHER_TRANSACTION_COST
    ) / (1.0 + STAMP_DUTY_RATE)
    return purchase_price, terminal_value, terminal_pv


def build_quarterly_2026(result_2026: YearResult) -> list[dict[str, float | str]]:
    q2_days = 91
    q3_days = 92
    q4_days = 92

    q2_property_fee = RENTABLE_AREA * BASE_OCCUPANCY[2026] * 10.0 * 3.0
    q3_property_fee = q2_property_fee
    q4_property_fee = q2_property_fee

    q2_parking = PARKING_SPACES * 3.0 * 6.0 * PARKING_USAGE[2026] * q2_days
    q3_parking = PARKING_SPACES * 3.0 * 6.0 * PARKING_USAGE[2026] * q3_days
    q4_parking = PARKING_SPACES * 3.0 * 6.0 * PARKING_USAGE[2026] * q4_days

    q2_rent = 0.0
    q3_rent = annual_rent_income(2026, BASE_OCCUPANCY[2026], 1.0) / RENT_COLLECTION_DAYS[2026] * q3_days
    q4_rent = annual_rent_income(2026, BASE_OCCUPANCY[2026], 1.0) / RENT_COLLECTION_DAYS[2026] * q4_days

    q_opex = result_2026.total_opex / 3.0

    return [
        {
            "quarter": "2026Q1",
            "rent_income": 0.0,
            "parking_income": 0.0,
            "property_fee_income": 0.0,
            "total_income": 0.0,
            "allocated_opex": 0.0,
            "pre_tax_operating_cash_flow": 0.0,
        },
        {
            "quarter": "2026Q2",
            "rent_income": q2_rent,
            "parking_income": q2_parking,
            "property_fee_income": q2_property_fee,
            "total_income": q2_rent + q2_parking + q2_property_fee,
            "allocated_opex": q_opex,
            "pre_tax_operating_cash_flow": q2_rent + q2_parking + q2_property_fee - q_opex,
        },
        {
            "quarter": "2026Q3",
            "rent_income": q3_rent,
            "parking_income": q3_parking,
            "property_fee_income": q3_property_fee,
            "total_income": q3_rent + q3_parking + q3_property_fee,
            "allocated_opex": q_opex,
            "pre_tax_operating_cash_flow": q3_rent + q3_parking + q3_property_fee - q_opex,
        },
        {
            "quarter": "2026Q4",
            "rent_income": q4_rent,
            "parking_income": q4_parking,
            "property_fee_income": q4_property_fee,
            "total_income": q4_rent + q4_parking + q4_property_fee,
            "allocated_opex": q_opex,
            "pre_tax_operating_cash_flow": q4_rent + q4_parking + q4_property_fee - q_opex,
        },
    ]


def scenario_purchase_price(stabilized_occupancy: float, terminal_rent_growth: float) -> float:
    occupancy = {
        2026: stabilized_occupancy - 0.15,
        2027: stabilized_occupancy,
        2028: stabilized_occupancy,
        2029: stabilized_occupancy,
    }
    rent_growth_factor = {
        2026: 1.00,
        2027: 1.00,
        2028: 1.00,
        2029: 1.00 + terminal_rent_growth,
    }
    results = build_year_results(occupancy, rent_growth_factor)
    purchase_price, _, _ = solve_purchase_price(results)
    return purchase_price


def write_csv(path: Path, rows: list[list[object]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.writer(fp)
        writer.writerows(rows)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    base_results = build_year_results(BASE_OCCUPANCY, BASE_RENT_GROWTH_FACTOR)
    purchase_price, terminal_value, terminal_pv = solve_purchase_price(base_results)
    initial_outflow = purchase_price * (1.0 + STAMP_DUTY_RATE) + OPENING_COST + OTHER_TRANSACTION_COST
    quarterly_2026 = build_quarterly_2026(base_results[0])

    sensitivity_rows = []
    for stabilized_occupancy in (0.90, 0.95, 1.00):
        row = {"stabilized_occupancy": stabilized_occupancy}
        for terminal_rent_growth in (0.00, 0.03, 0.05):
            scenario_key = f"rent_growth_{int(terminal_rent_growth * 100):02d}pct"
            row[scenario_key] = scenario_purchase_price(stabilized_occupancy, terminal_rent_growth)
        sensitivity_rows.append(row)

    summary = {
        "recommended_purchase_price": purchase_price,
        "initial_total_investment": initial_outflow,
        "terminal_value_2029": terminal_value,
        "present_value_of_terminal_value": terminal_pv,
        "base_case_after_tax_cash_flows": {
            str(result.year): result.after_tax_cash_flow for result in base_results
        },
    }

    cash_flow_csv_rows = [
        [
            "Year",
            "Occupancy",
            "Rent Growth Factor",
            "Rent Income",
            "Parking Income",
            "Property Fee Income",
            "Total Income",
            "Personnel Cost",
            "Marketing Cost",
            "Maint Repair Cost",
            "Property Mgmt Cost",
            "Energy Cost",
            "Admin Cost",
            "Capital Reno Cost",
            "Total Opex",
            "Pre-tax Profit",
            "Tax",
            "After-tax Cash Flow",
        ]
    ]
    for result in base_results:
        cash_flow_csv_rows.append(
            [
                result.year,
                result.occupancy,
                result.rent_growth_factor,
                result.rent_income,
                result.parking_income,
                result.property_fee_income,
                result.total_income,
                result.personnel_cost,
                result.marketing_cost,
                result.maint_repair_cost,
                result.property_mgmt_cost,
                result.energy_cost,
                result.admin_cost,
                result.capital_reno_cost,
                result.total_opex,
                result.pre_tax_profit,
                result.tax,
                result.after_tax_cash_flow,
            ]
        )

    quarterly_csv_rows = [
        [
            "Quarter",
            "Rent Income",
            "Parking Income",
            "Property Fee Income",
            "Total Income",
            "Allocated Opex",
            "Pre-tax Operating Cash Flow",
        ]
    ]
    for row in quarterly_2026:
        quarterly_csv_rows.append(
            [
                row["quarter"],
                row["rent_income"],
                row["parking_income"],
                row["property_fee_income"],
                row["total_income"],
                row["allocated_opex"],
                row["pre_tax_operating_cash_flow"],
            ]
        )

    sensitivity_csv_rows = [
        [
            "Stabilized Occupancy",
            "Purchase Price @ 0% Rent Growth",
            "Purchase Price @ 3% Rent Growth",
            "Purchase Price @ 5% Rent Growth",
        ]
    ]
    for row in sensitivity_rows:
        sensitivity_csv_rows.append(
            [
                row["stabilized_occupancy"],
                row["rent_growth_00pct"],
                row["rent_growth_03pct"],
                row["rent_growth_05pct"],
            ]
        )

    workbook_rows = [
        ["Case Study Calculation Output"],
        [],
        ["Summary"],
        ["Recommended Purchase Price", purchase_price],
        ["Initial Total Investment", initial_outflow],
        ["Terminal Value at 2029-12-31", terminal_value],
        ["Present Value of Terminal Value", terminal_pv],
        [],
        ["Annual Cash Flow"],
    ]
    workbook_rows.extend(cash_flow_csv_rows)
    workbook_rows.extend([[], ["2026 Quarterly Operating Improvement"]])
    workbook_rows.extend(quarterly_csv_rows)
    workbook_rows.extend([[], ["Sensitivity Analysis"]])
    workbook_rows.extend(sensitivity_csv_rows)

    write_csv(OUTPUT_DIR / "cash_flow_table.csv", cash_flow_csv_rows)
    write_csv(OUTPUT_DIR / "quarterly_2026.csv", quarterly_csv_rows)
    write_csv(OUTPUT_DIR / "sensitivity_table.csv", sensitivity_csv_rows)
    write_csv(OUTPUT_DIR / "case_study_calculations.csv", workbook_rows)

    with (OUTPUT_DIR / "case_study_results.json").open("w", encoding="utf-8") as fp:
        json.dump(
            {
                "summary": summary,
                "annual_results": [result.__dict__ for result in base_results],
                "quarterly_2026": quarterly_2026,
                "sensitivity": sensitivity_rows,
            },
            fp,
            ensure_ascii=False,
            indent=2,
        )


if __name__ == "__main__":
    main()
