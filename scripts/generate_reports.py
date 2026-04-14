from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side
import os
from config import TEMPLATE_DIR, OUTPUT_DIR


# =========================
# MAIN ENTRY
# =========================

def generate_reports(results, data):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for result in results:
        generate_single_report(result, data)


# =========================
# SINGLE REPORT
# =========================

def generate_single_report(result, data):
    TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "billing_template.xlsx")
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    tenant = get_tenant(result["tenant_id"], data)
    building = data["building"]
    unit = get_unit(result["unit_id"], data)

    # =========================
    # HEADER (FIXED CELLS)
    # =========================

    ws["B2"] = tenant.get("name", "")
    ws["B3"] = unit.get("unit_label", "")
    ws["B4"] = building.get("name", "")

    # =========================
    # COST TABLE
    # =========================

    start_row = 20
    current_row = start_row

    for line in result["lines"]:
        ws[f"A{current_row}"] = line.get("cost_type", "")
        ws[f"B{current_row}"] = line.get("allocation", "")
        ws[f"C{current_row}"] = line.get("share", "")
        ws[f"D{current_row}"] = line.get("amount", 0)

        current_row += 1

    # =========================
    # BORDER BEFORE TOTAL
    # =========================

    draw_bottom_border(ws, current_row)

    current_row += 1

    # =========================
    # TOTAL COSTS
    # =========================

    ws[f"C{current_row}"] = "Total costs"
    ws[f"D{current_row}"] = result["total_costs"]

    current_row += 1

    draw_bottom_border(ws, current_row)

    # =========================
    # PREPAYMENT
    # =========================

    prepayment = tenant.get("yearly_prepayment", 0)

    ws[f"C{current_row}"] = "Prepayment"
    ws[f"D{current_row}"] = prepayment

    current_row += 1

    # =========================
    # FINAL BALANCE
    # =========================

    balance = result["total_costs"] - prepayment

    ws[f"C{current_row}"] = "Balance (credit / debit)"
    ws[f"D{current_row}"] = balance

    # apply dotted fill from template row 20
    apply_template_fill(ws, source_row=20, target_row=current_row)

    # make final row bold
    ws[f"C{current_row}"].font = Font(bold=True)
    ws[f"D{current_row}"].font = Font(bold=True)

    # =========================
    # SAVE FILE
    # =========================

    filename = f"{OUTPUT_DIR}/tenant_{result['tenant_id']}.xlsx"
    wb.save(filename)


# =========================
# HELPERS
# =========================

def draw_bottom_border(ws, row):
    border = Border(bottom=Side(style="thin"))

    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{row}"].border = border


def apply_template_fill(ws, source_row, target_row):
    for col in ["A", "B", "C", "D"]:
        ws[f"{col}{target_row}"].fill = ws[f"{col}{source_row}"].fill


def get_tenant(tenant_id, data):
    for t in data["tenants"]:
        if t["tenant_id"] == tenant_id:
            return t
    raise ValueError("Tenant not found")


def get_unit(unit_id, data):
    for u in data["units"]:
        if u["unit_id"] == unit_id:
            return u
    raise ValueError("Unit not found")