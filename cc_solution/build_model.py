"""
宁波怡丰汇商业项目 – 工程经济学 2025 Case Study
完整财务模型 + Word 报告生成脚本
"""

import math
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              numbers)
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import docx.oxml.ns as qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────────────────────────────────────
# 1.  PARAMETERS
# ─────────────────────────────────────────────────────────────────────────────

# --- Area (sqm) ---
floors = {
    6:   {'gfa': 1021.56,  'gca': 1021.56,  'leasable': 141.7,    'owned_spots': 0,   'other_spots': 186},
    5:   {'gfa': 7855.27,  'gca': 7855.27,  'leasable': 391.45,   'owned_spots': 148, 'other_spots': 72},
    4:   {'gfa': 11560.31, 'gca': 11560.31, 'leasable': 10535.92, 'owned_spots': 0,   'other_spots': 0},
    3:   {'gfa': 14291.80, 'gca': 14076.81, 'leasable': 13052.43, 'owned_spots': 0,   'other_spots': 0},
    2:   {'gfa': 14237.26, 'gca': 14237.26, 'leasable': 13212.02, 'owned_spots': 0,   'other_spots': 0},
    1:   {'gfa': 14876.55, 'gca': 14331.43, 'leasable': 13128.34, 'owned_spots': 0,   'other_spots': 38},
    'B1':{'gfa': 43196.38, 'gca': 11751.42, 'leasable': 11751.42, 'owned_spots': 0,   'other_spots': 69},
    'B2':{'gfa': 0,        'gca': 2174.06,  'leasable': 416.48,   'owned_spots': 134, 'other_spots': 358},
}
# Rent rate: yuan / day / sqm (base year = 2026)
rent_rate = {6: 1.5, 5: 1.5, 4: 2.5, 3: 2.5, 2: 2.5, 1: 3.0, 'B1': 2.0, 'B2': 1.0}

TOTAL_GFA      = sum(f['gfa'] for f in floors.values())   # 107,039.13
TOTAL_LEASABLE = sum(f['leasable'] for f in floors.values())  # 62,629.76
TOTAL_SPOTS    = sum(f['owned_spots'] + f['other_spots'] for f in floors.values())  # 1,005

# --- Timeline ---
# Period 0: 2026-01-01   Period 1=2026  Period 2=2027  Period 3=2028  Period 4=2029
OPEN_DAYS_2026      = 275   # Apr 1 – Dec 31
FREE_RENT_DAYS_2026 = 91    # Apr 1 – Jun 30  (no rent)
RENT_DAYS_2026      = 184   # Jul 1 – Dec 31
MONTHS_2026         = 9     # operating months in 2026

# Quarterly breakdown 2026 (operating days)
Q2_DAYS = 91   # Apr–Jun
Q3_DAYS = 92   # Jul–Sep
Q4_DAYS = 92   # Oct–Dec

# --- Leasing assumptions ---
#   key = year, value = (occupancy, rent_growth_vs_prev_year)
lease_params = {
    2026: {'occ': 0.80, 'rent_g': 0.00},
    2027: {'occ': 0.95, 'rent_g': 0.00},
    2028: {'occ': 0.95, 'rent_g': 0.00},
    2029: {'occ': 0.95, 'rent_g': 0.03},
}

# --- Parking ---
parking = {
    2026: {'rate': 3, 'daily_hrs': 6, 'util': 0.50},
    2027: {'rate': 3, 'daily_hrs': 6, 'util': 0.60},
    2028: {'rate': 3, 'daily_hrs': 6, 'util': 0.70},
    2029: {'rate': 3, 'daily_hrs': 6, 'util': 0.80},
}

# --- Property management fee ---
PROP_FEE_PER_SQM_MONTH = 10  # yuan / sqm / month

# --- Operating cost base (full year) ---
PERSONNEL_BASE  = 5_000_000
MARKETING_BASE  = 5_000_000
MAINT_RATE      = 15   # yuan / sqm(GFA) / year  (basic maintenance)
CAPEX_RATIO     = 0.025   # 2.5% of rent income (capital improvement)
PROP_MGMT_RATE  = 30   # yuan / sqm(GFA) / year
UTIL_RATE       = 40   # yuan / sqm(GFA) / year
ADMIN_BASE      = 500_000

# Cost growth rates (vs previous full-year amount)
cost_growth = {
    # item: {2027: g, 2028: g, 2029: g}  (2026 = base, no growth)
    'personnel': {2027: 0.00, 2028: 0.03, 2029: 0.00},
    'marketing':  {2027:-0.05, 2028:-0.05, 2029:-0.05},
    'maint_basic':{2027: 0.03, 2028: 0.03, 2029: 0.03},
    'prop_mgmt':  {2027: 0.03, 2028: 0.03, 2029: 0.03},
    'utilities':  {2027: 0.01, 2028: 0.01, 2029: 0.01},
    'admin':      {2027: 0.03, 2028: 0.03, 2029: 0.03},
}

# --- Financial parameters ---
DEED_TAX_RATE       = 0.03
OTHER_FEES          = 11_000_000
OPENING_COST        = 10_000_000
CORP_TAX_RATE       = 0.25
MARR                = 0.12
EXIT_CAP_RATE       = 0.08

# ─────────────────────────────────────────────────────────────────────────────
# 2.  CORE CALCULATION FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def base_daily_rent():
    """Sum over all floors: leasable_area × daily_rent_rate   (yuan/day @ 100% occ)"""
    return sum(floors[k]['leasable'] * rent_rate[k] for k in floors)

BASE_DAILY_RENT = base_daily_rent()   # ~ 156,105 yuan/day

def rent_income(year, occ, rent_multiplier, days):
    """yuan"""
    return BASE_DAILY_RENT * rent_multiplier * occ * days

def parking_income(year, days):
    p = parking[year]
    return TOTAL_SPOTS * p['rate'] * p['daily_hrs'] * p['util'] * days

def property_fee_income(occ, months):
    """Property fee: collected on leased area even during rent-free period"""
    return TOTAL_LEASABLE * occ * PROP_FEE_PER_SQM_MONTH * months

def build_cost_schedule():
    """Returns dict[year][item] = annual cost (full year rates with growth)."""
    base = {
        'personnel':  PERSONNEL_BASE,
        'marketing':  MARKETING_BASE,
        'maint_basic': MAINT_RATE * TOTAL_GFA,
        'prop_mgmt':  PROP_MGMT_RATE * TOTAL_GFA,
        'utilities':  UTIL_RATE * TOTAL_GFA,
        'admin':      ADMIN_BASE,
    }
    schedule = {2026: dict(base)}
    prev = dict(base)
    for yr in [2027, 2028, 2029]:
        curr = {}
        for item in base:
            g = cost_growth[item][yr]
            curr[item] = prev[item] * (1 + g)
        schedule[yr] = curr
        prev = curr
    return schedule

COST_SCHEDULE = build_cost_schedule()

def opex_for_year(year, rent_inc):
    """Total operating costs for a year.
    2026: prorated by MONTHS_2026/12.  2027-2029: full year.
    Capex = 2.5% of rent income (rent_inc passed in).
    """
    c = COST_SCHEDULE[year]
    if year == 2026:
        factor = MONTHS_2026 / 12
    else:
        factor = 1.0
    opex = (c['personnel'] * factor
          + c['marketing']  * factor
          + c['maint_basic']* factor
          + c['prop_mgmt']  * factor
          + c['utilities']  * factor
          + c['admin']      * factor
          + rent_inc * CAPEX_RATIO)   # capex: no time proration needed (already in rent)
    return opex, {
        'personnel':  c['personnel'] * factor,
        'marketing':  c['marketing'] * factor,
        'maint_basic':c['maint_basic']* factor,
        'capex':      rent_inc * CAPEX_RATIO,
        'prop_mgmt':  c['prop_mgmt'] * factor,
        'utilities':  c['utilities'] * factor,
        'admin':      c['admin'] * factor,
    }

# ─────────────────────────────────────────────────────────────────────────────
# 3.  ANNUAL CASH FLOW MODEL
# ─────────────────────────────────────────────────────────────────────────────

def annual_cash_flows():
    """Returns list of dicts for years 2026-2029."""
    rent_mult = 1.0   # cumulative rent multiplier
    cfs = []
    for year in [2026, 2027, 2028, 2029]:
        occ = lease_params[year]['occ']
        rent_mult *= (1 + lease_params[year]['rent_g'])

        if year == 2026:
            rent_days, park_days, prop_months = RENT_DAYS_2026, OPEN_DAYS_2026, MONTHS_2026
        else:
            park_days_map = {2027: 365, 2028: 365, 2029: 365}
            rent_days = park_days = park_days_map[year]
            prop_months = 12

        rent  = rent_income(year, occ, rent_mult, rent_days)
        park  = parking_income(year, park_days)
        prop  = property_fee_income(occ, prop_months)
        total_rev = rent + park + prop

        total_opex, opex_detail = opex_for_year(year, rent)

        op_profit = total_rev - total_opex
        tax = max(0, op_profit * CORP_TAX_RATE)
        net_cf = op_profit - tax

        cfs.append({
            'year': year,
            'rent': rent, 'parking': park, 'prop_fee': prop,
            'total_rev': total_rev,
            **{'opex_' + k: v for k, v in opex_detail.items()},
            'total_opex': total_opex,
            'op_profit': op_profit,
            'tax': tax,
            'net_cf': net_cf,
            'rent_mult': rent_mult,
            'occ': occ,
        })
    return cfs

ANNUAL_CFS = annual_cash_flows()

# Exit value (perpetuity on 2029 after-tax operating CF)
NET_CF_2029 = ANNUAL_CFS[-1]['net_cf']
EXIT_VALUE  = NET_CF_2029 / EXIT_CAP_RATE
CF4_TOTAL   = NET_CF_2029 + EXIT_VALUE   # 2029 total CF incl. exit

# ─────────────────────────────────────────────────────────────────────────────
# 4.  ACQUISITION PRICE  (NPV = 0 at MARR = 12%)
# ─────────────────────────────────────────────────────────────────────────────

def pv_operating_cfs():
    pv = 0
    for i, cf in enumerate(ANNUAL_CFS):
        t = i + 1
        c = cf['net_cf']
        if i == 3:
            c += EXIT_VALUE   # add exit value to 2029
        pv += c / (1 + MARR) ** t
    return pv

PV_OPS = pv_operating_cfs()
# NPV = -(P*1.03 + OTHER_FEES + OPENING_COST) + PV_OPS = 0
# P*1.03 = PV_OPS - OTHER_FEES - OPENING_COST
# P = (PV_OPS - OTHER_FEES - OPENING_COST) / 1.03
ACQ_PRICE = (PV_OPS - OTHER_FEES - OPENING_COST) / (1 + DEED_TAX_RATE)
TOTAL_INV  = ACQ_PRICE * (1 + DEED_TAX_RATE) + OTHER_FEES + OPENING_COST

# ─────────────────────────────────────────────────────────────────────────────
# 5.  QUARTERLY CASH FLOWS  (2026 only, pre-tax for analysis)
# ─────────────────────────────────────────────────────────────────────────────

def quarterly_2026():
    """Returns quarterly breakdown for 2026 (Q1-Q4)."""
    occ = lease_params[2026]['occ']
    rent_mult = 1.0

    quarters = []
    # Q1: no operations
    quarters.append({'q': 'Q1 2026\n(1月–3月)', 'days': 0,
                     'rent': 0, 'parking': 0, 'prop_fee': 0,
                     'total_rev': 0, 'total_opex': 0,
                     'op_profit': 0, 'note': '收购/改造期，尚未开业'})

    # Q2: open, free-rent
    park_q2 = parking_income(2026, Q2_DAYS)
    prop_q2 = property_fee_income(occ, 3)   # 3 months
    rev_q2  = park_q2 + prop_q2
    opex_q2 = sum(COST_SCHEDULE[2026][k] * (Q2_DAYS / 365)
                  for k in ['personnel','marketing','maint_basic','prop_mgmt','utilities','admin'])
    op_profit_q2 = rev_q2 - opex_q2
    quarters.append({'q': 'Q2 2026\n(4月–6月)', 'days': Q2_DAYS,
                     'rent': 0, 'parking': park_q2, 'prop_fee': prop_q2,
                     'total_rev': rev_q2, 'total_opex': opex_q2,
                     'op_profit': op_profit_q2, 'note': '免租期：停车费+物业费'})

    # Q3: rent starts
    rent_q3  = BASE_DAILY_RENT * rent_mult * occ * Q3_DAYS
    park_q3  = parking_income(2026, Q3_DAYS)
    prop_q3  = property_fee_income(occ, 3)
    rev_q3   = rent_q3 + park_q3 + prop_q3
    opex_q3  = sum(COST_SCHEDULE[2026][k] * (Q3_DAYS / 365)
                   for k in ['personnel','marketing','maint_basic','prop_mgmt','utilities','admin'])
    opex_q3 += rent_q3 * CAPEX_RATIO
    op_profit_q3 = rev_q3 - opex_q3
    quarters.append({'q': 'Q3 2026\n(7月–9月)', 'days': Q3_DAYS,
                     'rent': rent_q3, 'parking': park_q3, 'prop_fee': prop_q3,
                     'total_rev': rev_q3, 'total_opex': opex_q3,
                     'op_profit': op_profit_q3, 'note': '租金正式收取'})

    # Q4: same structure as Q3
    rent_q4  = rent_q3
    park_q4  = park_q3
    prop_q4  = prop_q3
    rev_q4   = rev_q3
    opex_q4  = opex_q3
    op_profit_q4 = op_profit_q3
    quarters.append({'q': 'Q4 2026\n(10月–12月)', 'days': Q4_DAYS,
                     'rent': rent_q4, 'parking': park_q4, 'prop_fee': prop_q4,
                     'total_rev': rev_q4, 'total_opex': opex_q4,
                     'op_profit': op_profit_q4, 'note': '稳定运营'})

    # Annual tax on total 2026 profit
    total_2026_profit = max(0, op_profit_q2 + op_profit_q3 + op_profit_q4)
    annual_tax_2026   = total_2026_profit * CORP_TAX_RATE
    return quarters, annual_tax_2026

QUARTERS_2026, TAX_2026 = quarterly_2026()

# ─────────────────────────────────────────────────────────────────────────────
# 6.  SENSITIVITY ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

def compute_acq_price(stable_occ, rent_growth_2029):
    """Recompute acquisition price under different assumptions."""
    rent_mult = 1.0
    cfs_local = []
    for i, year in enumerate([2026, 2027, 2028, 2029]):
        if year == 2026:
            occ   = 0.80
            g_r   = 0.0
        else:
            occ   = stable_occ
            g_r   = rent_growth_2029 if year == 2029 else 0.0
        rent_mult *= (1 + g_r)

        if year == 2026:
            r_days, p_days, p_months = RENT_DAYS_2026, OPEN_DAYS_2026, MONTHS_2026
        else:
            r_days = p_days = 365; p_months = 12

        rent  = BASE_DAILY_RENT * rent_mult * occ * r_days
        park  = parking_income(year, p_days)
        prop  = property_fee_income(occ, p_months)
        rev   = rent + park + prop
        opex, _ = opex_for_year(year, rent)
        op_pr = rev - opex
        tax   = max(0, op_pr * CORP_TAX_RATE)
        net   = op_pr - tax
        cfs_local.append(net)

    exit_v = cfs_local[-1] / EXIT_CAP_RATE
    pv = sum((c + (exit_v if j == 3 else 0)) / (1 + MARR) ** (j + 1)
             for j, c in enumerate(cfs_local))
    return (pv - OTHER_FEES - OPENING_COST) / (1 + DEED_TAX_RATE)

OCC_LIST        = [0.70, 0.75, 0.80, 0.85, 0.90, 0.95]
RENT_GROWTH_LIST= [-0.03, 0.00, 0.03, 0.05, 0.08]

sens_table = []
for occ in OCC_LIST:
    row = []
    for rg in RENT_GROWTH_LIST:
        row.append(compute_acq_price(occ, rg))
    sens_table.append(row)

# ─────────────────────────────────────────────────────────────────────────────
# 7.  HELPER  –  EXCEL STYLES
# ─────────────────────────────────────────────────────────────────────────────

NAVY   = "1F3864"
BLUE   = "2E75B6"
LBLUE  = "BDD7EE"
YELLOW = "FFF2CC"
GREEN  = "E2EFDA"
ORANGE = "FCE4D6"
GREY   = "F2F2F2"
WHITE  = "FFFFFF"

def hdr_style(ws, cell, txt, bold=True, bg=NAVY, fg=WHITE, size=11, wrap=False, halign='center'):
    c = ws[cell] if isinstance(cell, str) else cell
    c.value = txt
    c.font  = Font(bold=bold, color=fg, size=size, name='微软雅黑')
    c.fill  = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal=halign, vertical='center', wrap_text=wrap)

def num_cell(ws, cell, val, fmt='#,##0', bold=False, bg=WHITE, color='000000'):
    c = ws[cell] if isinstance(cell, str) else cell
    c.value       = val
    c.number_format = fmt
    c.font        = Font(bold=bold, color=color, name='微软雅黑', size=10)
    c.fill        = PatternFill('solid', fgColor=bg)
    c.alignment   = Alignment(horizontal='right', vertical='center')

def txt_cell(ws, cell, txt, bold=False, bg=WHITE, halign='left', size=10, color='000000', wrap=True):
    c = ws[cell] if isinstance(cell, str) else cell
    c.value       = txt
    c.font        = Font(bold=bold, color=color, name='微软雅黑', size=size)
    c.fill        = PatternFill('solid', fgColor=bg)
    c.alignment   = Alignment(horizontal=halign, vertical='center', wrap_text=wrap)

def thin_border(ws, min_row, max_row, min_col, max_col):
    thin = Side(style='thin', color='BFBFBF')
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

FMT_WAN = '#,##0.00'   # 万元
FMT_INT = '#,##0'

def yuan_to_wan(v):
    return v / 10_000

# ─────────────────────────────────────────────────────────────────────────────
# 8.  BUILD EXCEL
# ─────────────────────────────────────────────────────────────────────────────

wb = Workbook()
wb.remove(wb.active)   # remove default sheet

# ── Sheet 0: 封面与目录 ──────────────────────────────────────────────────────
ws0 = wb.create_sheet("封面")
ws0.sheet_view.showGridLines = False
ws0.column_dimensions['A'].width = 5
ws0.column_dimensions['B'].width = 60
ws0.row_dimensions[1].height = 40

ws0.merge_cells('B2:B3')
c = ws0['B2']
c.value = "宁波怡丰汇商业项目"
c.font  = Font(bold=True, size=22, color=NAVY, name='微软雅黑')
c.alignment = Alignment(horizontal='center', vertical='center')

ws0.merge_cells('B4:B4')
c = ws0['B4']
c.value = "工程经济学 2025 Case Study  –  投资分析报告"
c.font  = Font(size=13, color=BLUE, name='微软雅黑')
c.alignment = Alignment(horizontal='center', vertical='center')

ws0['B6']  = "本工作簿包含以下内容："
ws0['B6'].font = Font(size=11, name='微软雅黑')

toc = [
    ("Step 1", "现金流模型（年度）", "年度现金流"),
    ("Step 1+", "季度现金流（2026年）", "季度现金流"),
    ("Step 2", "收购价格估算", "收购价格"),
    ("Step 3", "敏感性分析", "敏感性分析"),
    ("附录",   "主要参数汇总", "主要假设"),
]
for i, (step, desc, sheet) in enumerate(toc, start=8):
    ws0[f'B{i}'] = f'  {step}：{desc}（见"{sheet}"工作表）'
    ws0[f'B{i}'].font = Font(size=10, name='微软雅黑', color=BLUE)

# ── Sheet 1: 主要假设 ─────────────────────────────────────────────────────────
ws1 = wb.create_sheet("主要假设")
ws1.sheet_view.showGridLines = False

cols = ['B','C','D','E','F','G','H']
widths = [28, 16, 16, 16, 16, 16, 16]
for col, w in zip(cols, widths):
    ws1.column_dimensions[col].width = w

# Title
ws1.merge_cells('B1:H1')
hdr_style(ws1, 'B1', '宁波怡丰汇商业项目 – 主要参数汇总', size=13, bg=NAVY)

r = 3
# Section 1: Areas
ws1.merge_cells(f'B{r}:H{r}')
hdr_style(ws1, f'B{r}', '1. 项目面积（平方米）', bg=BLUE, size=11)
r += 1
for col, h in zip(['B','C','D','E','F','G'], ['楼层','建筑面积','产权面积','可租面积','产权车位','无产权车位']):
    hdr_style(ws1, f'{col}{r}', h, bg=LBLUE, fg=NAVY, size=10)
r += 1
for fl, d in floors.items():
    for col, val in zip(['B','C','D','E','F','G'],
                        [str(fl), d['gfa'], d['gca'], d['leasable'], d['owned_spots'], d['other_spots']]):
        if col == 'B':
            txt_cell(ws1, f'{col}{r}', val, halign='center')
        else:
            num_cell(ws1, f'{col}{r}', val if val else 0)
    r += 1
# Totals
hdr_style(ws1, f'B{r}', '合计', bg=GREY, fg=NAVY, size=10)
num_cell(ws1, f'C{r}', TOTAL_GFA, bold=True, bg=GREY)
num_cell(ws1, f'D{r}', sum(f['gca'] for f in floors.values()), bold=True, bg=GREY)
num_cell(ws1, f'E{r}', TOTAL_LEASABLE, bold=True, bg=GREY)
num_cell(ws1, f'F{r}', sum(f['owned_spots'] for f in floors.values()), bold=True, bg=GREY)
num_cell(ws1, f'G{r}', sum(f['other_spots'] for f in floors.values()), bold=True, bg=GREY)
r += 2

# Section 2: Rent rates
ws1.merge_cells(f'B{r}:H{r}')
hdr_style(ws1, f'B{r}', '2. 租金假设（首年平均日租金，元/天/平方米）', bg=BLUE, size=11)
r += 1
for fl in floors:
    txt_cell(ws1, f'B{r}', f'{fl}层', halign='center')
    num_cell(ws1, f'C{r}', rent_rate[fl], fmt='0.0')
    r += 1
r += 1

# Section 3: Leasing params
ws1.merge_cells(f'B{r}:H{r}')
hdr_style(ws1, f'B{r}', '3. 租赁参数', bg=BLUE, size=11)
r += 1
for col, h in zip(['B','C','D','E','F'], ['年份','出租率','租金增长率（vs上年）','租金收取天数','运营天数']):
    hdr_style(ws1, f'{col}{r}', h, bg=LBLUE, fg=NAVY, size=10)
r += 1
days_by_year = {2026: (RENT_DAYS_2026, OPEN_DAYS_2026), 2027: (365,365), 2028: (365,365), 2029: (365,365)}
for yr in [2026,2027,2028,2029]:
    lp = lease_params[yr]
    txt_cell(ws1, f'B{r}', str(yr), halign='center')
    num_cell(ws1, f'C{r}', lp['occ'], fmt='0%')
    num_cell(ws1, f'D{r}', lp['rent_g'], fmt='0%')
    num_cell(ws1, f'E{r}', days_by_year[yr][0])
    num_cell(ws1, f'F{r}', days_by_year[yr][1])
    r += 1
r += 1

# Section 4: Parking
ws1.merge_cells(f'B{r}:H{r}')
hdr_style(ws1, f'B{r}', '4. 停车费假设', bg=BLUE, size=11)
r += 1
for col, h in zip(['B','C','D','E'], ['年份','时费（元/时/位）','日均停车时长（时）','车位使用率']):
    hdr_style(ws1, f'{col}{r}', h, bg=LBLUE, fg=NAVY, size=10)
r += 1
for yr in [2026,2027,2028,2029]:
    p = parking[yr]
    txt_cell(ws1, f'B{r}', str(yr), halign='center')
    num_cell(ws1, f'C{r}', p['rate'], fmt='0')
    num_cell(ws1, f'D{r}', p['daily_hrs'], fmt='0')
    num_cell(ws1, f'E{r}', p['util'], fmt='0%')
    r += 1
r += 1

# Section 5: Opex
ws1.merge_cells(f'B{r}:H{r}')
hdr_style(ws1, f'B{r}', '5. 运营成本假设（首年全年基准，元）', bg=BLUE, size=11)
r += 1
items = [('人事费用', PERSONNEL_BASE, '元/年'),
         ('市场推广费', MARKETING_BASE, '元/年'),
         ('维保维修（基本）', MAINT_RATE, '元/㎡(GFA)/年'),
         ('资本性改造', f'{CAPEX_RATIO:.1%}', '×租金收入'),
         ('物业管理费', PROP_MGMT_RATE, '元/㎡(GFA)/年'),
         ('能耗费用', UTIL_RATE, '元/㎡(GFA)/年'),
         ('行政费用', ADMIN_BASE, '元/年')]
for name, val, unit in items:
    txt_cell(ws1, f'B{r}', name)
    if isinstance(val, str):
        txt_cell(ws1, f'C{r}', val, halign='right')
    else:
        num_cell(ws1, f'C{r}', val)
    txt_cell(ws1, f'D{r}', unit)
    r += 1
r += 1

# Section 6: Cost growth
ws1.merge_cells(f'B{r}:H{r}')
hdr_style(ws1, f'B{r}', '6. 费用增长率（相对上一年）', bg=BLUE, size=11)
r += 1
for col, h in zip(['B','C','D','E','F'], ['费项','2026','2027','2028','2029']):
    hdr_style(ws1, f'{col}{r}', h, bg=LBLUE, fg=NAVY, size=10)
r += 1
item_names = {'personnel':'人事费用','marketing':'市场推广费','maint_basic':'维保维修',
              'prop_mgmt':'物业管理','utilities':'能耗费用','admin':'行政费用'}
for k, name in item_names.items():
    txt_cell(ws1, f'B{r}', name)
    num_cell(ws1, f'C{r}', 0, fmt='0%')
    for j, yr in enumerate([2027,2028,2029], start=1):
        num_cell(ws1, f'{chr(ord("C")+j)}{r}', cost_growth[k][yr], fmt='0%')
    r += 1
r += 1

# Section 7: Financial params
ws1.merge_cells(f'B{r}:H{r}')
hdr_style(ws1, f'B{r}', '7. 财务参数', bg=BLUE, size=11)
r += 1
fp = [('交易契税税率', DEED_TAX_RATE, '3%'),
      ('其他交易费用', OTHER_FEES, '元'),
      ('商场开业成本', OPENING_COST, '元'),
      ('企业所得税税率', CORP_TAX_RATE, '25%'),
      ('MARR（折现率）', MARR, '12%'),
      ('退出资本化率', EXIT_CAP_RATE, '8%'),
      ('免租期', 3, '个月')]
for name, val, note in fp:
    txt_cell(ws1, f'B{r}', name)
    num_cell(ws1, f'C{r}', val, fmt='0%' if isinstance(val, float) and val < 1 else '#,##0')
    txt_cell(ws1, f'D{r}', note)
    r += 1

thin_border(ws1, 1, r, 2, 8)

# ── Sheet 2: 年度现金流 ───────────────────────────────────────────────────────
ws2 = wb.create_sheet("年度现金流")
ws2.sheet_view.showGridLines = False

col_w = [3, 36, 16, 16, 16, 16, 16]
for i, w in enumerate(col_w, start=1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# Title
ws2.merge_cells('B1:G1')
hdr_style(ws2, 'B1', 'Step 1：年度经营现金流模型（万元）', size=13, bg=NAVY)

ws2.merge_cells('B2:G2')
txt_cell(ws2, 'B2', '注：期初（Period 0）= 2026年1月1日；2026年为9个月运营期；所有金额单位为万元。',
         size=9, color='595959')

# Header row
r = 4
headers = ['项目', '2026年\n（改造期）', '2027年\n（稳定期）', '2028年\n（稳定期）', '2029年\n（稳定期+退出）', '备注']
for i, h in enumerate(headers, start=2):
    hdr_style(ws2, f'{get_column_letter(i)}{r}', h, bg=BLUE, size=10, wrap=True)
ws2.row_dimensions[r].height = 36

r += 1

def add_row(label, vals, bold=False, bg=WHITE, fmt=FMT_WAN, indent=0):
    global r
    txt_cell(ws2, f'B{r}', ('  ' * indent) + label, bold=bold, bg=bg)
    for j, v in enumerate(vals, start=3):
        num_cell(ws2, f'{get_column_letter(j)}{r}', yuan_to_wan(v) if v is not None else None,
                 fmt=fmt, bold=bold, bg=bg)
    r += 1

def add_sep(label, bg=LBLUE):
    global r
    ws2.merge_cells(f'B{r}:G{r}')
    hdr_style(ws2, f'B{r}', label, bg=bg, fg=NAVY, size=10)
    r += 1

cf = ANNUAL_CFS

add_sep('▶ 一、营业收入', bg=GREEN)
add_row('租金收入', [c['rent'] for c in cf], indent=1)
add_row('停车费收入', [c['parking'] for c in cf], indent=1)
add_row('物业费收入', [c['prop_fee'] for c in cf], indent=1)
add_row('营业收入合计', [c['total_rev'] for c in cf], bold=True, bg=GREEN)

add_sep('▶ 二、运营成本', bg=ORANGE)
add_row('人事费用', [c['opex_personnel'] for c in cf], indent=1)
add_row('市场推广费', [c['opex_marketing'] for c in cf], indent=1)
add_row('维保维修（基本）', [c['opex_maint_basic'] for c in cf], indent=1)
add_row('资本性改造（占租金2.5%）', [c['opex_capex'] for c in cf], indent=1)
add_row('物业管理费', [c['opex_prop_mgmt'] for c in cf], indent=1)
add_row('能耗费用', [c['opex_utilities'] for c in cf], indent=1)
add_row('行政费用', [c['opex_admin'] for c in cf], indent=1)
add_row('运营成本合计', [c['total_opex'] for c in cf], bold=True, bg=ORANGE)

add_sep('▶ 三、营业利润与税', bg=YELLOW)
add_row('营业利润（税前）', [c['op_profit'] for c in cf], bold=True)
add_row('企业所得税（25%）', [c['tax'] for c in cf], indent=1)
add_row('税后经营净现金流', [c['net_cf'] for c in cf], bold=True, bg=YELLOW)

add_sep('▶ 四、退出价值（永续年金 i=8%）', bg=LBLUE)
exit_row = [None, None, None, EXIT_VALUE]
add_row('退出价值（不计税）', exit_row, bold=True, bg=LBLUE)

add_sep('▶ 五、年度总现金流', bg=GREEN)
total_cf = [c['net_cf'] for c in cf]
total_cf[3] += EXIT_VALUE
add_row('年度总现金流（含退出）', total_cf, bold=True, bg=GREEN)

# Key metrics summary
r += 1
ws2.merge_cells(f'B{r}:G{r}')
hdr_style(ws2, f'B{r}', '关键运营指标', bg=NAVY, size=11)
r += 1
add_row('出租率', [c['occ'] for c in cf], fmt='0.0%')
add_row('实际运营天数', [OPEN_DAYS_2026, 365, 365, 365], fmt='#,##0')
add_row('租金收入/收入合计', [c['rent']/c['total_rev'] for c in cf], fmt='0.0%')
add_row('税前净利润率', [c['op_profit']/c['total_rev'] for c in cf], fmt='0.0%')

thin_border(ws2, 1, r+1, 2, 7)

# ── Sheet 3: 季度现金流 ───────────────────────────────────────────────────────
ws3 = wb.create_sheet("季度现金流")
ws3.sheet_view.showGridLines = False

col_w3 = [3, 30, 16, 16, 16, 16]
for i, w in enumerate(col_w3, start=1):
    ws3.column_dimensions[get_column_letter(i)].width = w

ws3.merge_cells('B1:F1')
hdr_style(ws3, 'B1', 'Step 1（补充）：2026年季度现金流分析（万元）', size=13, bg=NAVY)
ws3.merge_cells('B2:F2')
txt_cell(ws3, 'B2', '说明：Q1为收购期无运营；Q2免租期仅有停车费+物业费；Q3起租金开始收取。税收按年度汇总计算。',
         size=9, color='595959')

r3 = 4
hdrs3 = ['项目', 'Q1 (1-3月)', 'Q2 (4-6月)', 'Q3 (7-9月)', 'Q4 (10-12月)']
for i, h in enumerate(hdrs3, start=2):
    hdr_style(ws3, f'{get_column_letter(i)}{r3}', h, bg=BLUE, size=10, wrap=True)
ws3.row_dimensions[r3].height = 30
r3 += 1

def add_q_row(label, vals, bold=False, bg=WHITE, fmt=FMT_WAN):
    global r3
    txt_cell(ws3, f'B{r3}', label, bold=bold, bg=bg)
    for j, v in enumerate(vals, start=3):
        num_cell(ws3, f'{get_column_letter(j)}{r3}',
                 yuan_to_wan(v) if v is not None else '-',
                 fmt=fmt, bold=bold, bg=bg)
    r3 += 1

q = QUARTERS_2026

ws3.merge_cells(f'B{r3}:F{r3}')
hdr_style(ws3, f'B{r3}', '收入', bg=GREEN, fg=NAVY, size=10); r3 += 1
add_q_row('租金收入', [0, 0, q[2]['rent'], q[3]['rent']], bg=GREEN)
add_q_row('停车费收入', [0, q[1]['parking'], q[2]['parking'], q[3]['parking']])
add_q_row('物业费收入', [0, q[1]['prop_fee'], q[2]['prop_fee'], q[3]['prop_fee']])
add_q_row('收入小计', [0, q[1]['total_rev'], q[2]['total_rev'], q[3]['total_rev']], bold=True, bg=GREEN)

ws3.merge_cells(f'B{r3}:F{r3}')
hdr_style(ws3, f'B{r3}', '运营成本', bg=ORANGE, fg=NAVY, size=10); r3 += 1
add_q_row('运营成本小计', [0, q[1]['total_opex'], q[2]['total_opex'], q[3]['total_opex']], bg=ORANGE)

ws3.merge_cells(f'B{r3}:F{r3}')
hdr_style(ws3, f'B{r3}', '利润（税前，季度口径）', bg=YELLOW, fg=NAVY, size=10); r3 += 1
add_q_row('营业利润（含季度分摊）', [0, q[1]['op_profit'], q[2]['op_profit'], q[3]['op_profit']],
          bold=True, bg=YELLOW)

r3 += 1
ws3.merge_cells(f'B{r3}:F{r3}')
hdr_style(ws3, f'B{r3}', '2026全年汇总（与年度现金流匹配）', bg=NAVY, size=10)
r3 += 1
cf26 = ANNUAL_CFS[0]
add_q_row('2026年度税后净现金流（年度口径）',
          [None, None, None, cf26['net_cf']], bold=True, bg=LBLUE)

r3 += 2
ws3.merge_cells(f'B{r3}:F{r3}')
hdr_style(ws3, f'B{r3}', '季度现金流改善分析', bg=NAVY, size=10)
r3 += 1
txt_cell(ws3, f'B{r3}', 'Q1 → Q2：运营启动（仅停车费+物业费），税前利润由0→负（因分摊的人事/管理/能耗成本较高）',
         size=9)
r3 += 1
txt_cell(ws3, f'B{r3}', 'Q2 → Q3：免租期结束，租金收入涌入，季度税前利润大幅转正（约+' +
         f'{yuan_to_wan(q[2]["op_profit"]-q[1]["op_profit"]):.0f}万元）',
         size=9)
r3 += 1
txt_cell(ws3, f'B{r3}', 'Q3 = Q4（2026）：同等运营条件，现金流基本持平',
         size=9)
r3 += 1
txt_cell(ws3, f'B{r3}', f'改造期年度净CF（2026）≈ {yuan_to_wan(cf26["net_cf"]):.0f}万元  vs  '
         f'稳定期（2027）≈ {yuan_to_wan(ANNUAL_CFS[1]["net_cf"]):.0f}万元，相差约'
         f'{yuan_to_wan(ANNUAL_CFS[1]["net_cf"]-cf26["net_cf"]):.0f}万元（'
         f'{(ANNUAL_CFS[1]["net_cf"]/cf26["net_cf"]-1)*100:.0f}%增幅）',
         size=9, bold=True, color=NAVY)

thin_border(ws3, 1, r3+1, 2, 6)

# ── Sheet 4: 收购价格 ─────────────────────────────────────────────────────────
ws4 = wb.create_sheet("收购价格")
ws4.sheet_view.showGridLines = False

for i, w in enumerate([3,32,20,20,20,20,20], start=1):
    ws4.column_dimensions[get_column_letter(i)].width = w

ws4.merge_cells('B1:G1')
hdr_style(ws4, 'B1', 'Step 2：收购价格估算（MARR = 12%）', size=13, bg=NAVY)

r4 = 3

ws4.merge_cells(f'B{r4}:G{r4}')
hdr_style(ws4, f'B{r4}', '一、未来现金流现值计算（折现率12%）', bg=BLUE, size=11)
r4 += 1

hdrs4 = ['时间节点', '期间', '年度总CF（万元）', '折现系数', '现值（万元）']
for i, h in enumerate(hdrs4, start=2):
    hdr_style(ws4, f'{get_column_letter(i)}{r4}', h, bg=LBLUE, fg=NAVY, size=10)
r4 += 1

total_cfs4 = []
for i, cf in enumerate(ANNUAL_CFS):
    t = i + 1
    total_cf_yr = cf['net_cf'] + (EXIT_VALUE if i == 3 else 0)
    disc_factor = 1 / (1 + MARR) ** t
    pv = total_cf_yr * disc_factor
    total_cfs4.append((t, cf['year'], total_cf_yr, disc_factor, pv))

    yr_label = f"{cf['year']}年"
    if i == 0:
        note = '改造期（9个月运营）'
    elif i == 3:
        note = '稳定期 + 退出价值'
    else:
        note = '稳定期'

    txt_cell(ws4, f'B{r4}', yr_label, halign='center')
    txt_cell(ws4, f'C{r4}', note)
    num_cell(ws4, f'D{r4}', yuan_to_wan(total_cf_yr), fmt=FMT_WAN)
    num_cell(ws4, f'E{r4}', disc_factor, fmt='0.0000')
    num_cell(ws4, f'F{r4}', yuan_to_wan(pv), fmt=FMT_WAN, bold=True)
    r4 += 1

pv_sum = sum(t[4] for t in total_cfs4)
hdr_style(ws4, f'B{r4}', '未来CF现值合计', bg=GREEN, fg=NAVY, size=10)
num_cell(ws4, f'F{r4}', yuan_to_wan(pv_sum), bold=True, bg=GREEN)
r4 += 2

ws4.merge_cells(f'B{r4}:G{r4}')
hdr_style(ws4, f'B{r4}', '二、收购价格推算', bg=BLUE, size=11)
r4 += 1

items4 = [
    ('未来CF现值合计', pv_sum, '万元', False),
    ('  减：其他交易费用', -OTHER_FEES, '万元', False),
    ('  减：商场开业成本', -OPENING_COST, '万元', False),
    ('  = 可用于支付收购对价×(1+3%)', pv_sum - OTHER_FEES - OPENING_COST, '万元', True),
    ('  ÷ (1 + 契税税率 3%)', 1+DEED_TAX_RATE, '倍', False),
    ('★ 建议收购对价（P）', ACQ_PRICE, '万元', True),
    ('  对应契税', ACQ_PRICE * DEED_TAX_RATE, '万元', False),
    ('  其他交易费用', OTHER_FEES, '万元', False),
    ('  商场开业成本', OPENING_COST, '万元', False),
    ('★ 总投资额（P×1.03 + 2100万）', TOTAL_INV, '万元', True),
]
for label, val, unit, bold in items4:
    txt_cell(ws4, f'B{r4}', label, bold=bold, bg=YELLOW if bold else WHITE)
    if label.startswith('  ÷'):
        num_cell(ws4, f'C{r4}', val, fmt='0.00', bold=bold, bg=YELLOW if bold else WHITE)
    else:
        num_cell(ws4, f'C{r4}', yuan_to_wan(val) if unit=='万元' else val,
                 fmt=FMT_WAN if unit=='万元' else '0.00',
                 bold=bold, bg=YELLOW if bold else WHITE)
    txt_cell(ws4, f'D{r4}', unit, bg=YELLOW if bold else WHITE)
    r4 += 1

r4 += 1
ws4.merge_cells(f'B{r4}:G{r4}')
hdr_style(ws4, f'B{r4}', '三、投资回报验证（按建议收购价格）', bg=BLUE, size=11)
r4 += 1

npv_check = -TOTAL_INV + pv_sum
irr_approx = MARR  # by construction when NPV=0
equity_mult = (pv_sum) / TOTAL_INV

metrics4 = [
    ('NPV（验证，应≈0）', npv_check, '万元'),
    ('IRR（理论值）', irr_approx, '= MARR = 12%'),
    ('总投资回报倍数（含退出价值）', equity_mult, '倍'),
    ('退出价值 / 总现金流入', EXIT_VALUE / pv_sum, '%'),
    ('总CF现值 / 总投资', pv_sum / TOTAL_INV, '倍'),
]
for label, val, unit in metrics4:
    txt_cell(ws4, f'B{r4}', label)
    if unit == '万元':
        num_cell(ws4, f'C{r4}', yuan_to_wan(val), fmt=FMT_WAN, bold=True)
    elif unit in ('倍', '%'):
        num_cell(ws4, f'C{r4}', val, fmt='0.00', bold=True)
    else:
        num_cell(ws4, f'C{r4}', val, fmt='0%', bold=True)
    txt_cell(ws4, f'D{r4}', unit)
    r4 += 1

thin_border(ws4, 1, r4+1, 2, 7)

# ── Sheet 5: 敏感性分析 ───────────────────────────────────────────────────────
ws5 = wb.create_sheet("敏感性分析")
ws5.sheet_view.showGridLines = False

for i, w in enumerate([3,22,16,16,16,16,16], start=1):
    ws5.column_dimensions[get_column_letter(i)].width = w

ws5.merge_cells('B1:G1')
hdr_style(ws5, 'B1', 'Step 3：敏感性分析 – 建议收购对价（万元，MARR=12%）', size=13, bg=NAVY)
ws5.merge_cells('B2:G2')
txt_cell(ws5, 'B2', '分析变量：稳定期（2027-2029年）出租率 × 2029年租金增长率；底色越深代表收购价格越高。',
         size=9, color='595959')

r5 = 4
# Col headers: rent growth
hdr_style(ws5, f'B{r5}', '出租率 \\ 租金增长率', bg=NAVY, size=10)
for j, rg in enumerate(RENT_GROWTH_LIST, start=3):
    hdr_style(ws5, f'{get_column_letter(j)}{r5}', f'{rg:.0%}', bg=BLUE, size=10)
r5 += 1

# Find range for color scale
all_prices = [sens_table[i][j] for i in range(len(OCC_LIST)) for j in range(len(RENT_GROWTH_LIST))]
min_p, max_p = min(all_prices), max(all_prices)

def price_color(p):
    ratio = (p - min_p) / (max_p - min_p) if max_p > min_p else 0.5
    # green (low) → yellow → red (high)
    if ratio < 0.5:
        r_ = int(255 * ratio * 2)
        g_ = 200
    else:
        r_ = 200
        g_ = int(200 * (1 - ratio))
    return f'{r_:02X}{g_:02X}80'

for i, occ in enumerate(OCC_LIST):
    hdr_style(ws5, f'B{r5}', f'{occ:.0%}', bg=LBLUE, fg=NAVY, size=10)
    for j, rg in enumerate(RENT_GROWTH_LIST):
        p = sens_table[i][j]
        col_bg = price_color(p)
        is_base = (occ == 0.95 and rg == 0.03)
        num_cell(ws5, f'{get_column_letter(j+3)}{r5}',
                 yuan_to_wan(p), fmt=FMT_WAN,
                 bold=is_base, bg=YELLOW if is_base else col_bg)
    r5 += 1

r5 += 1

# One-variable analysis: occupancy
ws5.merge_cells(f'B{r5}:G{r5}')
hdr_style(ws5, f'B{r5}', '单变量分析①：出租率变化对收购价格的影响（租金增长率固定3%）', bg=BLUE, size=11)
r5 += 1
for col, h in zip(['B','C','D'], ['出租率', '收购价格（万元）', 'vs基准（95%, 3%）']):
    hdr_style(ws5, f'{col}{r5}', h, bg=LBLUE, fg=NAVY, size=10)
r5 += 1
base_p = compute_acq_price(0.95, 0.03)
for occ in OCC_LIST:
    p = compute_acq_price(occ, 0.03)
    txt_cell(ws5, f'B{r5}', f'{occ:.0%}', halign='center')
    num_cell(ws5, f'C{r5}', yuan_to_wan(p), fmt=FMT_WAN, bold=(occ == 0.95))
    num_cell(ws5, f'D{r5}', (p - base_p) / abs(base_p), fmt='0.0%',
             color='C00000' if p < base_p else '375623')
    r5 += 1

r5 += 1

# One-variable analysis: rent growth
ws5.merge_cells(f'B{r5}:G{r5}')
hdr_style(ws5, f'B{r5}', '单变量分析②：2029年租金增长率变化对收购价格的影响（出租率固定95%）', bg=BLUE, size=11)
r5 += 1
for col, h in zip(['B','C','D'], ['租金增长率', '收购价格（万元）', 'vs基准（95%, 3%）']):
    hdr_style(ws5, f'{col}{r5}', h, bg=LBLUE, fg=NAVY, size=10)
r5 += 1
for rg in RENT_GROWTH_LIST:
    p = compute_acq_price(0.95, rg)
    txt_cell(ws5, f'B{r5}', f'{rg:.0%}', halign='center')
    num_cell(ws5, f'C{r5}', yuan_to_wan(p), fmt=FMT_WAN, bold=(rg == 0.03))
    num_cell(ws5, f'D{r5}', (p - base_p) / abs(base_p), fmt='0.0%',
             color='C00000' if p < base_p else '375623')
    r5 += 1

r5 += 2
ws5.merge_cells(f'B{r5}:G{r5}')
hdr_style(ws5, f'B{r5}', '敏感性结论', bg=NAVY, size=10)
r5 += 1
txt_cell(ws5, f'B{r5}',
    f'基准假设（出租率95%，租金增长率3%）下建议收购价约 {yuan_to_wan(ACQ_PRICE):.0f} 万元（'
    f'{ACQ_PRICE/1e8:.2f}亿元）。出租率每下降5%，收购对价降低约'
    f'{yuan_to_wan(compute_acq_price(0.90,0.03)-compute_acq_price(0.85,0.03)):.0f}万元；'
    f'若出租率仅达80%（最悲观），合理价格约{yuan_to_wan(compute_acq_price(0.80,0.00)):.0f}万元。',
    size=10, wrap=True)
ws5.row_dimensions[r5].height = 48

thin_border(ws5, 1, r5+1, 2, 7)

# ─────────────────────────────────────────────────────────────────────────────
# 9.  SAVE EXCEL
# ─────────────────────────────────────────────────────────────────────────────

excel_path = '/home/wangyy/CaseStudy/宁波怡丰汇投资分析.xlsx'
wb.save(excel_path)
print(f"Excel saved: {excel_path}")

# ─────────────────────────────────────────────────────────────────────────────
# 10.  WORD DOCUMENT  (完整报告)
# ─────────────────────────────────────────────────────────────────────────────

def set_run_fmt(run, bold=False, size=11, color=None):
    run.font.name = '微软雅黑'
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))

def add_heading(doc, text, level=1, color=NAVY):
    p = doc.add_heading(level=level)
    p.clear()
    run = p.add_run(text)
    set_run_fmt(run, bold=True, size={1:14,2:12,3:11}.get(level,11), color=color)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    return p

def add_para(doc, text, bold=False, size=11, indent=0, space_before=2, space_after=2):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_fmt(run, bold=bold, size=size)
    p.paragraph_format.left_indent  = Cm(indent)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    return p

def add_table(doc, headers, rows, col_widths=None):
    """Create a simple formatted table."""
    table = doc.add_table(rows=1+len(rows), cols=len(headers))
    table.style = 'Table Grid'
    # Header row
    hdr_row = table.rows[0]
    for j, h in enumerate(headers):
        cell = hdr_row.cells[j]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.font.bold = True
            run.font.name = '微软雅黑'
            run.font.size = Pt(10)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shading = OxmlElement('w:shd')
        shading.set(qn.qn('w:fill'), '2E75B6')
        shading.set(qn.qn('w:color'), 'FFFFFF')
        shading.set(qn.qn('w:val'), 'clear')
        cell._tc.get_or_add_tcPr().append(shading)
    # Data rows
    for i, row_data in enumerate(rows):
        row = table.rows[i+1]
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.text = str(val)
            for run in cell.paragraphs[0].runs:
                run.font.name = '微软雅黑'
                run.font.size = Pt(10)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
    if col_widths:
        for j, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[j].width = Cm(w)
    return table

doc = Document()
# Page margins
section = doc.sections[0]
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(3)
section.right_margin  = Cm(2.5)

# ── Cover ──
doc.add_paragraph()
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_title.add_run('宁波怡丰汇商业项目投资分析报告')
set_run_fmt(r, bold=True, size=18, color=NAVY)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_sub.add_run('工程经济学 2025 Case Study')
set_run_fmt(r, size=13, color=BLUE)

doc.add_paragraph()

# ── Step 1: 现金流模型 ─────────────────────────────────────────────────────────
add_heading(doc, 'Step 1：建立现金流模型')

add_para(doc, '基于Excel文件提供的基础假设，本报告以2026年1月1日为期初（Period 0），构建了以下结构的年度现金流模型。')

# 1.1 Timeline
add_heading(doc, '1.1  项目时间线与现金流结构', level=2)
add_para(doc, '【期初 Period 0】2026年1月1日：一次性支付全部投资成本（收购对价、契税、其他费用、开业成本）')
add_para(doc, '【改造期 2026年】：4月1日开业，设3个月免租期（4–6月），7月起正式收取租金；运营时间合计9个月（275天）', indent=0.5)
add_para(doc, '【稳定期 2027–2029年】：全年运营，出租率95%，2029年租金增长3%', indent=0.5)
add_para(doc, '【退出 2029年12月31日】：以永续年金模型（i=8%）计算退出价值（不计税）', indent=0.5)

# 1.2 Annual CF table
add_heading(doc, '1.2  年度现金流汇总（单位：万元）', level=2)
cf_table_headers = ['项目', '2026年', '2027年', '2028年', '2029年']
cf_table_rows = [
    ['出租率', f'{ANNUAL_CFS[0]["occ"]:.0%}', f'{ANNUAL_CFS[1]["occ"]:.0%}',
     f'{ANNUAL_CFS[2]["occ"]:.0%}', f'{ANNUAL_CFS[3]["occ"]:.0%}'],
    ['── 收入 ──', '', '', '', ''],
    ['租金收入', f'{yuan_to_wan(ANNUAL_CFS[0]["rent"]):.0f}',
                 f'{yuan_to_wan(ANNUAL_CFS[1]["rent"]):.0f}',
                 f'{yuan_to_wan(ANNUAL_CFS[2]["rent"]):.0f}',
                 f'{yuan_to_wan(ANNUAL_CFS[3]["rent"]):.0f}'],
    ['停车费收入', f'{yuan_to_wan(ANNUAL_CFS[0]["parking"]):.0f}',
                  f'{yuan_to_wan(ANNUAL_CFS[1]["parking"]):.0f}',
                  f'{yuan_to_wan(ANNUAL_CFS[2]["parking"]):.0f}',
                  f'{yuan_to_wan(ANNUAL_CFS[3]["parking"]):.0f}'],
    ['物业费收入', f'{yuan_to_wan(ANNUAL_CFS[0]["prop_fee"]):.0f}',
                  f'{yuan_to_wan(ANNUAL_CFS[1]["prop_fee"]):.0f}',
                  f'{yuan_to_wan(ANNUAL_CFS[2]["prop_fee"]):.0f}',
                  f'{yuan_to_wan(ANNUAL_CFS[3]["prop_fee"]):.0f}'],
    ['营业收入合计', f'{yuan_to_wan(ANNUAL_CFS[0]["total_rev"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[1]["total_rev"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[2]["total_rev"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[3]["total_rev"]):.0f}'],
    ['── 成本 ──', '', '', '', ''],
    ['运营成本合计', f'{yuan_to_wan(ANNUAL_CFS[0]["total_opex"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[1]["total_opex"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[2]["total_opex"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[3]["total_opex"]):.0f}'],
    ['── 利润 ──', '', '', '', ''],
    ['营业利润（税前）', f'{yuan_to_wan(ANNUAL_CFS[0]["op_profit"]):.0f}',
                        f'{yuan_to_wan(ANNUAL_CFS[1]["op_profit"]):.0f}',
                        f'{yuan_to_wan(ANNUAL_CFS[2]["op_profit"]):.0f}',
                        f'{yuan_to_wan(ANNUAL_CFS[3]["op_profit"]):.0f}'],
    ['企业所得税（25%）', f'{yuan_to_wan(ANNUAL_CFS[0]["tax"]):.0f}',
                          f'{yuan_to_wan(ANNUAL_CFS[1]["tax"]):.0f}',
                          f'{yuan_to_wan(ANNUAL_CFS[2]["tax"]):.0f}',
                          f'{yuan_to_wan(ANNUAL_CFS[3]["tax"]):.0f}'],
    ['税后净现金流', f'{yuan_to_wan(ANNUAL_CFS[0]["net_cf"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[1]["net_cf"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[2]["net_cf"]):.0f}',
                    f'{yuan_to_wan(ANNUAL_CFS[3]["net_cf"]):.0f}'],
    ['退出价值（永续年金）', '—', '—', '—', f'{yuan_to_wan(EXIT_VALUE):.0f}'],
    ['年度总CF（含退出）', f'{yuan_to_wan(ANNUAL_CFS[0]["net_cf"]):.0f}',
                           f'{yuan_to_wan(ANNUAL_CFS[1]["net_cf"]):.0f}',
                           f'{yuan_to_wan(ANNUAL_CFS[2]["net_cf"]):.0f}',
                           f'{yuan_to_wan(ANNUAL_CFS[3]["net_cf"]+EXIT_VALUE):.0f}'],
]
add_table(doc, cf_table_headers, cf_table_rows, col_widths=[5, 2.5, 2.5, 2.5, 2.5])

# 1.3 Quarterly
add_heading(doc, '1.3  2026年季度现金流分析', level=2)
add_para(doc, '问题一：现金流能否逐季度改善？')
add_para(doc, '是的，2026年运营现金流呈现显著的季度递增趋势：', bold=True)
q_rows = [
    ['Q1（1–3月）', '改造收购期，无运营收入', '—', '—', '0'],
    ['Q2（4–6月）',
     f'开业+免租期，停车费{yuan_to_wan(QUARTERS_2026[1]["parking"]):.0f}万+物业费{yuan_to_wan(QUARTERS_2026[1]["prop_fee"]):.0f}万',
     f'{yuan_to_wan(QUARTERS_2026[1]["total_rev"]):.0f}',
     f'{yuan_to_wan(QUARTERS_2026[1]["total_opex"]):.0f}',
     f'{yuan_to_wan(QUARTERS_2026[1]["op_profit"]):.0f}（税前，含年度税后调整）'],
    ['Q3（7–9月）',
     f'租金开始收取，收入大幅提升',
     f'{yuan_to_wan(QUARTERS_2026[2]["total_rev"]):.0f}',
     f'{yuan_to_wan(QUARTERS_2026[2]["total_opex"]):.0f}',
     f'{yuan_to_wan(QUARTERS_2026[2]["op_profit"]):.0f}'],
    ['Q4（10–12月）',
     f'同Q3，平稳运营',
     f'{yuan_to_wan(QUARTERS_2026[3]["total_rev"]):.0f}',
     f'{yuan_to_wan(QUARTERS_2026[3]["total_opex"]):.0f}',
     f'{yuan_to_wan(QUARTERS_2026[3]["op_profit"]):.0f}'],
]
add_table(doc, ['季度','主要特征','收入(万元)','成本(万元)','税前利润(万元)'],
          q_rows, col_widths=[2.5, 6, 2, 2, 3])

add_para(doc,
    f'免租期（Q2）由于成本已按比例摊入但租金为零，税前经营利润为负（约{yuan_to_wan(QUARTERS_2026[1]["op_profit"]):.0f}万元）。'
    f'进入Q3后，租金收入贡献约{yuan_to_wan(QUARTERS_2026[2]["rent"]):.0f}万元/季度，现金流大幅改善。',
    size=10, space_before=4)

add_para(doc, '问题二：改造期与稳定期现金流差异')
add_para(doc,
    f'改造期2026年税后净现金流约 {yuan_to_wan(ANNUAL_CFS[0]["net_cf"]):.0f} 万元，'
    f'稳定期2027年约 {yuan_to_wan(ANNUAL_CFS[1]["net_cf"]):.0f} 万元，'
    f'差距约 {yuan_to_wan(ANNUAL_CFS[1]["net_cf"]-ANNUAL_CFS[0]["net_cf"]):.0f} 万元（约{(ANNUAL_CFS[1]["net_cf"]/ANNUAL_CFS[0]["net_cf"]-1)*100:.0f}% 增幅）。'
    f'主要原因：①2026年仅9个月运营且含3个月免租，②出租率从80%提升至95%，③停车费利用率提高。',
    size=10)

# ── Step 2: 收购价格 ───────────────────────────────────────────────────────────
add_heading(doc, 'Step 2：收购价格估算')
add_para(doc,
    f'以MARR=12%为基准，令NPV=0，反推合理收购对价。计算公式为：')
add_para(doc,
    f'NPV = -[P×(1+3%) + 1,100万 + 1,000万] + Σ CFₜ/(1.12)ᵗ = 0',
    bold=True, indent=0.5)

# PV table
pv_rows = []
for t, yr, cf_yr, df, pv in total_cfs4:
    pv_rows.append([f't={t}（{yr}年）',
                    f'{yuan_to_wan(cf_yr):.0f}',
                    f'{df:.4f}',
                    f'{yuan_to_wan(pv):.0f}'])
pv_rows.append(['未来CF现值合计', '', '', f'{yuan_to_wan(pv_sum):.0f}'])
add_table(doc, ['时间节点', '年度总CF（万元）', '折现系数(12%)', '现值（万元）'],
          pv_rows, col_widths=[4, 3.5, 3, 3])

doc.add_paragraph()
add_para(doc,
    f'由此推导：P = (现值合计 − 其他费用 − 开业成本) / 1.03 = '
    f'({yuan_to_wan(pv_sum):.0f} − 1,100 − 1,000) / 1.03 ≈ '
    f'**{yuan_to_wan(ACQ_PRICE):.0f} 万元（约 {ACQ_PRICE/1e8:.2f} 亿元）**',
    bold=True)

add_para(doc, f'总投资额（含契税、其他费用和开业成本）：约 {yuan_to_wan(TOTAL_INV):.0f} 万元（{TOTAL_INV/1e8:.2f} 亿元）')
add_para(doc,
    f'建议收购价格区间：{yuan_to_wan(ACQ_PRICE*0.92):.0f}–{yuan_to_wan(ACQ_PRICE*1.08):.0f}万元（±8%弹性空间，对应不同风险溢价与谈判空间）。',
    size=10)

# ── Step 3: 敏感性分析 ─────────────────────────────────────────────────────────
add_heading(doc, 'Step 3：敏感性分析')
add_para(doc,
    '商业地产的核心不确定性来自出租率和租金水平。以下分析改变稳定期出租率（2027–2029）和2029年租金增长率，观察收购价格的变化。')

# Sensitivity table in word
sens_hdr = ['出租率 / 租金增长率'] + [f'{rg:.0%}' for rg in RENT_GROWTH_LIST]
sens_rows_word = []
for i, occ in enumerate(OCC_LIST):
    row_w = [f'{occ:.0%}']
    for j in range(len(RENT_GROWTH_LIST)):
        row_w.append(f'{yuan_to_wan(sens_table[i][j]):.0f}')
    sens_rows_word.append(row_w)
add_table(doc, sens_hdr, sens_rows_word)

add_para(doc, '注：表中数字为建议收购对价（万元）；★标记基准情形（95%, 3%）。', size=9)

add_heading(doc, '3.1  关键结论', level=2)

occ_impact = (compute_acq_price(0.90, 0.03) - compute_acq_price(0.85, 0.03))
rg_impact   = (compute_acq_price(0.95, 0.05) - compute_acq_price(0.95, 0.00))

bullets = [
    f'出租率对价格影响最为显著：从70%提升至95%，建议收购价差距超过{yuan_to_wan(compute_acq_price(0.95,0.03)-compute_acq_price(0.70,0.03)):.0f}万元（约{(compute_acq_price(0.95,0.03)/compute_acq_price(0.70,0.03)-1)*100:.0f}%）。',
    f'出租率每变动5%，收购价格约变动{yuan_to_wan(occ_impact):.0f}万元。',
    f'租金增长率对价格影响相对较小：从0%升至5%，价格仅增约{yuan_to_wan(compute_acq_price(0.95,0.05)-compute_acq_price(0.95,0.00)):.0f}万元（{(compute_acq_price(0.95,0.05)/compute_acq_price(0.95,0.00)-1)*100:.0f}%），因其仅在2029年生效且经8%资本化率放大。',
    f'最悲观情形（出租率70%，租金不增长）：建议收购价约{yuan_to_wan(compute_acq_price(0.70,0.00)):.0f}万元，比基准低约{yuan_to_wan(compute_acq_price(0.95,0.03)-compute_acq_price(0.70,0.00)):.0f}万元。',
    f'投资风险点：由于退出价值（永续年金）占总现值的{EXIT_VALUE/pv_sum:.0%}，退出资本化率（i=8%）的变动对估值影响极大。若市场出清时i升至10%，退出价值将下降20%，建议收购价格需相应下调。',
]
for b in bullets:
    p = doc.add_paragraph(style='List Bullet')
    r_b = p.add_run(b)
    set_run_fmt(r_b, size=10)

# ── Step 4: 开放性思考 ─────────────────────────────────────────────────────────
add_heading(doc, 'Step 4：开放性思考')

add_heading(doc, '4.1  项目面临的潜在风险', level=2)

risks = [
    ('招商执行风险',
     f'商场当前出租率仅30%，历史上基本处于半停业状态，品牌信心不足。模型假设开业即达80%出租率，但实际招商周期可能延长2–4个月，导致首年现金流不足。建议在谈判中将收购价格与招商完成度挂钩。'),
    ('主力店流失风险',
     '影院（如万达影城、CGV）、健身（如超级猩猩、Keep线下店）为典型主力店，能带动全场客流，但经营波动大：疫情或消费降级时租金拖欠概率高。建议单一主力租户的面积占比控制在15%以内，并在合同中约定保底租金。'),
    ('区域竞争加剧',
     '周边已有凯德、九龙仓、华润等成熟商业体，江北区3公里内未来可能还有新竞品入市。社区型商场必须保持差异化定位，否则客流难以支撑目标出租率。'),
    ('退出流动性风险',
     f'模型中退出价值占现值的{EXIT_VALUE/pv_sum:.0%}，极度依赖2029年能否找到合适买方。若届时大宗商业物业市场流动性差，或资本化率由8%升至10%，退出价值将下降约{yuan_to_wan(EXIT_VALUE - ANNUAL_CFS[-1]["net_cf"]/0.10):.0f}万元，直接影响IRR。'),
    ('地铁通车不确定性',
     '8号线预计2026年通车，若延期，开业初期客流支撑力减弱，直接影响80%出租率假设。'),
    ('物业改造成本超支',
     '模型中资本性改造按租金收入的2.5%计提，但若前期招商需大规模重新分割楼层或更换设施，实际capex可能大幅超支，影响首年现金流。'),
]
for title, body in risks:
    p = doc.add_paragraph()
    r_title = p.add_run(f'【{title}】')
    set_run_fmt(r_title, bold=True, size=10, color=NAVY)
    r_body = p.add_run(f' {body}')
    set_run_fmt(r_body, size=10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(0.3)

add_heading(doc, '4.2  运营策略与业态选择建议', level=2)

add_para(doc, '若由我负责商场运营与招商，将按以下逻辑构建业态组合，兼顾稳定现金流与客流支撑：')

strategy = [
    ('B1/B2（超市+餐饮集合）',
     '引入知名生鲜超市（如盒马、大润发奥莱）作为全楼流量发动机，同时在B1设置小吃/快餐集合区（日式、中式、东南亚）。社区超市具有极高出行频率，可为全楼提供稳定基础客流；且其通常愿意签订5–10年长期合同，保证了现金流稳定性。'),
    ('1层（商业价值最高）',
     '主打具有IP属性的餐饮旗舰（如网红咖啡、烘焙甜品）和快时尚品牌，通过高曝光率提升形象与租金。优先选择知名连锁（需支付保底），避免过度依赖单一餐饮品牌。'),
    ('2–3层（体验消费核心层）',
     '设置儿童教育综合体（早教+素质拓展）和亲子零售（玩具+服装），精准对接周边高密度家庭客群，且该业态受电商冲击较小。可搭配美甲、美容等生活美学类，形成"周末家庭日"场景。'),
    ('4层（运动健康）',
     '引入1家中等规模专业健身中心（1,500–2,000㎡），搭配运动品牌零售（LULULEMON、Decathlon等）。注意控制单一健身品牌面积，避免过度集中风险。'),
    ('5–6层（文化娱乐/可持续）',
     '5层保留给小型影院或体验式展览空间，5–6层可作为弹性办公或创业孵化器（低租金但稳定，且能带来日常消费人流）。顶层露天区打造特色市集/快闪，增加社交媒体曝光。'),
    ('核心招商原则',
     '①锁定长约（主力店≥5年），优化现金流可预测性；②控制单一业态集中度（不超过20%面积）；③优先选择线下不可替代的业态（体验、服务、生鲜）；④保留10%灵活铺位用于快闪和试验性品牌，提升话题热度。'),
]
for title, body in strategy:
    p = doc.add_paragraph()
    r_title = p.add_run(f'▶ {title}：')
    set_run_fmt(r_title, bold=True, size=10, color=BLUE)
    r_body = p.add_run(body)
    set_run_fmt(r_body, size=10)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.left_indent = Cm(0.3)

add_heading(doc, '4.3  综合投资可行性判断', level=2)
add_para(doc,
    f'在基准假设下（出租率95%、2029年租金增长3%），建议收购对价约 {yuan_to_wan(ACQ_PRICE):.0f} 万元（{ACQ_PRICE/1e8:.2f}亿元）。'
    f'项目具备以下优势：区位成熟、地铁赋能确定、周边人口密度高、当前估值处于历史低位（出租率仅30%，上行空间明显）。'
    f'主要风险在于招商执行难度大、首年现金流较薄（仅约{yuan_to_wan(ANNUAL_CFS[0]["net_cf"]):.0f}万元），以及退出高度依赖资本市场环境。',
    size=10)
add_para(doc,
    '综合判断：若能以低于建议价约5–10%完成交割（以80%出租率情景作为保守定价基准），结合专业商管团队的运营改善，该项目具备合理的风险收益比，建议在尽调确认产权清晰、无重大租户纠纷后，按敏感性分析下限作为谈判锚定价格推进交易。',
    size=10, bold=True)

# Footer note
doc.add_paragraph()
p_note = doc.add_paragraph()
r_note = p_note.add_run(
    '注：本报告所有计算基于Excel文件"宁波怡丰汇投资分析.xlsx"中的财务模型，详细计算步骤见对应工作表。')
set_run_fmt(r_note, size=9, color='595959')

# Save Word
word_path = '/home/wangyy/CaseStudy/投资分析报告.docx'
doc.save(word_path)
print(f"Word saved: {word_path}")

# ─────────────────────────────────────────────────────────────────────────────
# 11.  PRINT SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

print("\n" + "="*60)
print("   宁波怡丰汇商业项目 – 财务模型计算结果摘要")
print("="*60)
print(f"\n总可租面积:     {TOTAL_LEASABLE:,.0f} sqm")
print(f"总建筑面积:     {TOTAL_GFA:,.0f} sqm")
print(f"总车位数:       {TOTAL_SPOTS} 个")
print(f"基准日租金收入: {BASE_DAILY_RENT:,.0f} 元/天 (100%出租率)")

print("\n── 年度税后净现金流（万元）──")
for c in ANNUAL_CFS:
    print(f"  {c['year']}: 收入 {yuan_to_wan(c['total_rev']):,.0f}  "
          f"成本 {yuan_to_wan(c['total_opex']):,.0f}  "
          f"税后CF {yuan_to_wan(c['net_cf']):,.0f}")

print(f"\n退出价值 (CF2029/{EXIT_CAP_RATE:.0%}): {yuan_to_wan(EXIT_VALUE):,.0f} 万元")
print(f"PV_ops (MARR=12%):  {yuan_to_wan(PV_OPS):,.0f} 万元")
print(f"\n★ 建议收购价格:  {yuan_to_wan(ACQ_PRICE):,.0f} 万元 ({ACQ_PRICE/1e8:.3f} 亿元)")
print(f"  总投资额:        {yuan_to_wan(TOTAL_INV):,.0f} 万元 ({TOTAL_INV/1e8:.3f} 亿元)")
print(f"  NPV验证 (≈0):   {yuan_to_wan(-TOTAL_INV + PV_OPS):,.1f} 万元")
