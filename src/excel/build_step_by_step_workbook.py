#!/usr/bin/env python3

import json
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape
from xml.sax.saxutils import escape


REPO_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = REPO_ROOT / "outputs" / "submission"
RESULTS_PATH = REPO_ROOT / "outputs" / "case_study_results.json"


def col_name(index: int) -> str:
    out = []
    while index > 0:
        index, rem = divmod(index - 1, 26)
        out.append(chr(65 + rem))
    return "".join(reversed(out))


def make_cell(ref: str, value=None, cell_type="n", style=0, formula=None) -> str:
    if value is None and formula is None:
        return ""
    if cell_type == "s":
        return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t>{escape(str(value))}</t></is></c>'
    if formula is not None:
        cached = "" if value is None else f"<v>{value}</v>"
        return f'<c r="{ref}" s="{style}"><f>{escape(formula)}</f>{cached}</c>'
    return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'


def make_sheet_xml(rows, widths=None) -> str:
    cols_xml = ""
    if widths:
        cols = []
        for i, width in enumerate(widths, 1):
            cols.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
        cols_xml = "<cols>" + "".join(cols) + "</cols>"

    row_xml = []
    for r_idx, row in enumerate(rows, 1):
        cells = []
        for c_idx, cell in enumerate(row, 1):
            ref = f"{col_name(c_idx)}{r_idx}"
            if isinstance(cell, dict):
                cells.append(make_cell(ref, cell.get("value"), cell.get("type", "n"), cell.get("style", 0), cell.get("formula")))
            elif isinstance(cell, str):
                cells.append(make_cell(ref, cell, "s", 0))
            elif cell is None:
                continue
            else:
                cells.append(make_cell(ref, cell, "n", 0))
        row_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"{cols_xml}<sheetData>{''.join(row_xml)}</sheetData></worksheet>"
    )


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
    <numFmt numFmtId="165" formatCode="0.00%"/>
    <numFmt numFmtId="166" formatCode="#,##0"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F1"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4F81BD"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def wrap_str(text, style=0):
    return {"type": "s", "value": text, "style": style}


def wrap_num(value, style=0):
    return {"type": "n", "value": value, "style": style}


def wrap_formula(formula, value, style=0):
    return {"formula": formula, "value": value, "style": style}


def build_workbook() -> dict:
    with RESULTS_PATH.open(encoding="utf-8") as fp:
        data = json.load(fp)
    annual = data["annual_results"]
    assumptions = data["assumptions"]
    quarterly = data["quarterly_2026"]
    sensitivity = data["sensitivity"]
    floor_rows = data["floor_rent_table"]
    opex = data["opex_schedule"]
    summary = data["summary"]

    sheets = []

    summary_rows = [
        [wrap_str("Case Study Summary", 2), None, None, None],
        [wrap_str("Metric", 1), wrap_str("Value", 1)],
        [wrap_str("Recommended Purchase Price"), wrap_num(summary["recommended_purchase_price"], 3)],
        [wrap_str("Initial Total Investment"), wrap_num(summary["initial_total_investment"], 3)],
        [wrap_str("Terminal Value"), wrap_num(summary["terminal_value_2029"], 3)],
        [wrap_str("PV of Terminal Value"), wrap_num(summary["present_value_of_terminal_value"], 3)],
        [wrap_str("MARR"), wrap_num(assumptions["marr"], 4)],
        [wrap_str("Exit Cap Rate"), wrap_num(assumptions["exit_cap_rate"], 4)],
    ]
    sheets.append(("Summary", summary_rows, [24, 18, 16, 16]))

    assumption_rows = [[wrap_str("Core Assumptions", 2)]]
    assumption_rows.append([wrap_str("Parameter", 1), wrap_str("Value", 1), wrap_str("Unit", 1)])
    for row in data["assumption_table"][1:]:
        style = 4 if row[2] == "%" else 3 if row[2] in {"RMB"} else 5 if row[2] in {"sqm", "spaces"} else 0
        assumption_rows.append([wrap_str(row[0]), wrap_num(row[1], style), wrap_str(row[2])])
    assumption_rows.extend([
        [],
        [wrap_str("Timeline Assumptions", 2)],
        [wrap_str("Year", 1), wrap_str("Operating Days", 1), wrap_str("Rent Collection Days", 1), wrap_str("Property Fee Months", 1), wrap_str("Occupancy", 1), wrap_str("Parking Usage", 1)],
    ])
    for year in (2026, 2027, 2028, 2029):
        assumption_rows.append([
            wrap_num(year, 5),
            wrap_num(assumptions["operating_days"][str(year)] if isinstance(next(iter(assumptions["operating_days"].keys())), str) else assumptions["operating_days"][year], 5),
            wrap_num(assumptions["rent_collection_days"][str(year)] if isinstance(next(iter(assumptions["rent_collection_days"].keys())), str) else assumptions["rent_collection_days"][year], 5),
            wrap_num(assumptions["property_fee_months"][str(year)] if isinstance(next(iter(assumptions["property_fee_months"].keys())), str) else assumptions["property_fee_months"][year], 5),
            wrap_num(assumptions["occupancy"][str(year)] if isinstance(next(iter(assumptions["occupancy"].keys())), str) else assumptions["occupancy"][year], 4),
            wrap_num(assumptions["parking_usage"][str(year)] if isinstance(next(iter(assumptions["parking_usage"].keys())), str) else assumptions["parking_usage"][year], 4),
        ])
    sheets.append(("Assumptions", assumption_rows, [28, 16, 16, 18, 12, 12]))

    floor_sheet_rows = [[wrap_str("Floor-Level Rent Build-Up", 2)]]
    floor_sheet_rows.append([wrap_str(v, 1) for v in floor_rows[0]])
    for row in floor_rows[1:]:
        floor_sheet_rows.append([wrap_str(row[0]), wrap_num(row[1], 3), wrap_num(row[2], 3) if row[2] != "" else None, wrap_num(row[3], 3)])
    sheets.append(("RentBuildUp", floor_sheet_rows, [12, 18, 22, 24]))

    revenue_rows = [
        [wrap_str("Revenue Build-Up", 2)],
        [wrap_str("Year", 1), wrap_str("Rent Income", 1), wrap_str("Parking Income", 1), wrap_str("Property Fee", 1), wrap_str("Total Income", 1)],
    ]
    for i, row in enumerate(annual, start=3):
        revenue_rows.append([
            wrap_num(row["year"], 5),
            wrap_num(row["rent_income"], 3),
            wrap_num(row["parking_income"], 3),
            wrap_num(row["property_fee_income"], 3),
            wrap_formula(f"SUM(B{i}:D{i})", row["total_income"], 3),
        ])
    revenue_rows.extend([
        [],
        [wrap_str("2026 Quarterly Ramp-Up", 2)],
        [wrap_str("Quarter", 1), wrap_str("Income", 1), wrap_str("Opex", 1), wrap_str("Pre-tax Operating CF", 1)],
    ])
    base = len(revenue_rows) + 1
    for idx, row in enumerate(quarterly):
        r = base + idx
        revenue_rows.append([
            wrap_str(row["quarter"]),
            wrap_num(row["total_income"], 3),
            wrap_num(row["allocated_opex"], 3),
            wrap_formula(f"B{r}-C{r}", row["pre_tax_operating_cash_flow"], 3),
        ])
    sheets.append(("Revenue", revenue_rows, [14, 18, 18, 20, 18]))

    cost_rows = [
        [wrap_str("Operating Cost, Tax, and Exit Value", 2)],
        [wrap_str("Year", 1), wrap_str("Personnel", 1), wrap_str("Marketing", 1), wrap_str("Maint/Repair", 1), wrap_str("Property Mgmt", 1), wrap_str("Energy", 1), wrap_str("Admin", 1), wrap_str("Capex Reno", 1), wrap_str("Total Opex", 1), wrap_str("Tax", 1), wrap_str("After-tax CF", 1)],
    ]
    for i, row in enumerate(annual, start=3):
        cost_rows.append([
            wrap_num(row["year"], 5),
            wrap_num(row["personnel_cost"], 3),
            wrap_num(row["marketing_cost"], 3),
            wrap_num(row["maint_repair_cost"], 3),
            wrap_num(row["property_mgmt_cost"], 3),
            wrap_num(row["energy_cost"], 3),
            wrap_num(row["admin_cost"], 3),
            wrap_num(row["capital_reno_cost"], 3),
            wrap_formula(f"SUM(B{i}:H{i})", row["total_opex"], 3),
            wrap_num(row["tax"], 3),
            wrap_num(row["after_tax_cash_flow"], 3),
        ])
    cost_rows.extend([
        [],
        [wrap_str("Terminal Value", 2)],
        [wrap_str("Metric", 1), wrap_str("Value", 1)],
        [wrap_str("2029 After-tax Cash Flow"), wrap_num(summary["base_case_after_tax_cash_flows"]["2029"], 3)],
        [wrap_str("Exit Cap Rate"), wrap_num(assumptions["exit_cap_rate"], 4)],
        [wrap_str("Terminal Value"), wrap_formula("B15/B16", summary["terminal_value_2029"], 3)],
    ])
    sheets.append(("CostTaxExit", cost_rows, [14, 16, 16, 16, 16, 16, 14, 16, 16, 16, 16]))

    valuation_rows = [
        [wrap_str("Valuation Walk-Through", 2)],
        [wrap_str("Line Item", 1), wrap_str("Formula", 1), wrap_str("Value", 1)],
        [wrap_str("PV of 2026 CF"), wrap_str("CF2026/(1+MARR)^1"), wrap_num(annual[0]["after_tax_cash_flow"] / (1.12), 3)],
        [wrap_str("PV of 2027 CF"), wrap_str("CF2027/(1+MARR)^2"), wrap_num(annual[1]["after_tax_cash_flow"] / (1.12**2), 3)],
        [wrap_str("PV of 2028 CF"), wrap_str("CF2028/(1+MARR)^3"), wrap_num(annual[2]["after_tax_cash_flow"] / (1.12**3), 3)],
        [wrap_str("PV of 2029 CF"), wrap_str("CF2029/(1+MARR)^4"), wrap_num(annual[3]["after_tax_cash_flow"] / (1.12**4), 3)],
        [wrap_str("PV of terminal value"), wrap_str("TV/(1+MARR)^4"), wrap_num(summary["present_value_of_terminal_value"], 3)],
        [wrap_str("Opening cost"), wrap_str("Given"), wrap_num(assumptions["opening_cost"], 3)],
        [wrap_str("Other transaction cost"), wrap_str("Given"), wrap_num(assumptions["other_transaction_cost"], 3)],
        [wrap_str("Purchase price"), wrap_str("(sum PV - fixed costs)/(1+stamp duty)"), wrap_num(summary["recommended_purchase_price"], 3)],
    ]
    sheets.append(("Valuation", valuation_rows, [24, 32, 18]))

    sensitivity_rows = [
        [wrap_str("Sensitivity Analysis", 2)],
        [wrap_str("Occupancy", 1), wrap_str("0% growth", 1), wrap_str("3% growth", 1), wrap_str("5% growth", 1)],
    ]
    for row in sensitivity:
        sensitivity_rows.append([
            wrap_num(row["stabilized_occupancy"], 4),
            wrap_num(row["rent_growth_00pct"], 3),
            wrap_num(row["rent_growth_03pct"], 3),
            wrap_num(row["rent_growth_05pct"], 3),
        ])
    sheets.append(("Sensitivity", sensitivity_rows, [14, 18, 18, 18]))

    return {"sheets": sheets}


def write_xlsx(output_path: Path) -> None:
    workbook = build_workbook()
    sheets = workbook["sheets"]

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            + "".join(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for i in range(1, len(sheets) + 1)
            )
            + "</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>",
        )
        zf.writestr(
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>EE2025 Case Study Calculations</dc:title><dc:creator>Codex</dc:creator></cp:coreProperties>',
        )
        zf.writestr(
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            f"<Application>Codex</Application><TitlesOfParts><vt:vector size=\"{len(sheets)}\" baseType=\"lpstr\">"
            + "".join(f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name, _, _ in sheets)
            + f"</vt:vector></TitlesOfParts><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheets)}</vt:i4></vt:variant></vt:vector></HeadingPairs></Properties>",
        )
        workbook_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<sheets>"
            + "".join(
                f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                for i, (name, _, _) in enumerate(sheets, 1)
            )
            + "</sheets></workbook>"
        )
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(
                f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                for i in range(1, len(sheets) + 1)
            )
            + f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            + "</Relationships>",
        )
        zf.writestr("xl/styles.xml", styles_xml())
        for i, (_, rows, widths) in enumerate(sheets, 1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", make_sheet_xml(rows, widths))


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    write_xlsx(OUTPUT_DIR / "EE2025_Case_Study_Calculations.xlsx")


if __name__ == "__main__":
    main()
