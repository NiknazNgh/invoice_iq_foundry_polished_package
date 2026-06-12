from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch

OUT_DIR = Path("sample_invoices")
PROVIDER = "Texas GLO/State Power Program"

OUTPUT_COLUMNS = [
    "Provider",
    "Production Month",
    "From",
    "To",
    "Invoice Date",
    "Power Factor",
    "Load Factor",
    "Actual Demand (KW)",
    "Billing Demand (KW)",
    "4CP Charges Qty (KW)",
    "4CP Charges Rate ($/KW)",
    "4CP Charges ($)",
    "Usage - Actual KWH",
    "UOM",
    "Energy Charge",
    "Nodal Congestion Charge",
    "Market Securitization (Debt) Financing - Default Charge",
    "Prior Period Pass Through Charge",
    "ERCOT Cont Reserve Serv (ECRS)",
    "Firm Fuel Supply Service",
    "Firm Fuel Supply Service - Backbill",
    "Market Securitization - Uplift Charge",
    "TX-ERCOT Admin Fees - CIL",
    "Transmission Charges",
    "Taxes & PUC Assessment Charge",
    "Ancilliary Service Obligation Adjustment",
    "Other Taxes",
    "Bill Total",
]

BILL_COMPONENT_COLUMNS = [
    "Energy Charge",
    "Nodal Congestion Charge",
    "Market Securitization (Debt) Financing - Default Charge",
    "Prior Period Pass Through Charge",
    "ERCOT Cont Reserve Serv (ECRS)",
    "Firm Fuel Supply Service",
    "Firm Fuel Supply Service - Backbill",
    "Market Securitization - Uplift Charge",
    "TX-ERCOT Admin Fees - CIL",
    "Transmission Charges",
    "Taxes & PUC Assessment Charge",
    "Ancilliary Service Obligation Adjustment",
    "Other Taxes",
]

ACCOUNTS = [
    "VCWRF Main Plant",
    "VCWRF Dryer Facility A",
    "VCWRF Dryer Facility B",
    "North Service Pump Station",
    "South Service Pump Station",
    "Village Creek Lift Station",
    "Westside Water Plant",
    "Rolling Hills Water Plant",
    "North Holly Water Plant",
    "South Holly Water Plant",
    "Eagle Mountain Water Plant",
    "Biosolids Operations",
    "Large Meter Operations",
    "Remote Tank Site A",
    "Remote Tank Site B",
]


def money(value: float) -> str:
    return f"${value:,.2f}"


def number(value: float) -> str:
    return f"{value:,.2f}"


def pct(value: float) -> str:
    return f"{value:.1f}%"


def month_start(base_year: int, base_month: int, offset: int) -> date:
    month_num = base_month + offset
    year = base_year + (month_num - 1) // 12
    month = ((month_num - 1) % 12) + 1
    return date(year, month, 1)


def last_day(d: date) -> date:
    return date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def build_invoice(index: int) -> Tuple[Dict[str, object], List[Tuple[str, float, float, float]], str]:
    d_from = month_start(2025, 10, index)
    d_to = last_day(d_from)
    invoice_date = d_to + timedelta(days=5)

    account = ACCOUNTS[index]
    usage = 1_850_000 + index * 86_750 + (index % 3) * 42_300
    actual_demand = 3_900 + index * 145 + (index % 4) * 62
    billing_demand = actual_demand + 70 + (index % 5) * 15
    cp4_qty = billing_demand - 35 + (index % 4) * 9
    cp4_rate = 3.85 + (index % 6) * 0.215
    cp4_amount = round(cp4_qty * cp4_rate, 2)

    energy_rate = 0.0475 + (index % 5) * 0.0011
    energy_charge = round(usage * energy_rate, 2)
    nodal = round(950 + index * 83.25, 2)
    debt = round(usage * 0.00018, 2)
    prior = round(0 if index % 4 else 420 + index * 8.5, 2)
    ecrs = round(usage * 0.00105, 2)
    firm_fuel = round(usage * 0.00062, 2)
    backbill = round(0 if index % 5 else usage * 0.00011, 2)
    uplift = round(usage * 0.00027, 2)
    tx_ercot = round(85 + index * 4.75, 2)
    transmission = round(billing_demand * (7.95 + (index % 3) * 0.28), 2)
    pre_tax = energy_charge + nodal + debt + prior + ecrs + firm_fuel + backbill + uplift + tx_ercot + transmission
    taxes = round(pre_tax * 0.0045, 2)
    ancillary = round(usage * 0.00138, 2)
    other_taxes = round(0.00 if index % 6 else 37.50, 2)

    bill_total = round(
        energy_charge
        + nodal
        + debt
        + prior
        + ecrs
        + firm_fuel
        + backbill
        + uplift
        + tx_ercot
        + transmission
        + taxes
        + ancillary
        + other_taxes,
        2,
    )

    note = "PASS sample"

    # Add two controlled exceptions so the Flags/Exceptions tab is useful in the live demo.
    if index == 13:
        bill_total = round(bill_total + 127.43, 2)
        note = "Intentional Bill Total mismatch for demo flag"
    if index == 14:
        cp4_amount = round(cp4_amount + 99.99, 2)
        note = "Intentional 4CP mismatch for demo flag"

    row: Dict[str, object] = {
        "Provider": PROVIDER,
        "Production Month": d_from,
        "From": d_from,
        "To": d_to,
        "Invoice Date": invoice_date,
        "Power Factor": 0.945 + (index % 5) * 0.007,
        "Load Factor": 0.68 + (index % 6) * 0.024,
        "Actual Demand (KW)": actual_demand,
        "Billing Demand (KW)": billing_demand,
        "4CP Charges Qty (KW)": cp4_qty,
        "4CP Charges Rate ($/KW)": cp4_rate,
        "4CP Charges ($)": cp4_amount,
        "Usage - Actual KWH": usage,
        "UOM": "KWH",
        "Energy Charge": energy_charge,
        "Nodal Congestion Charge": nodal,
        "Market Securitization (Debt) Financing - Default Charge": debt,
        "Prior Period Pass Through Charge": prior,
        "ERCOT Cont Reserve Serv (ECRS)": ecrs,
        "Firm Fuel Supply Service": firm_fuel,
        "Firm Fuel Supply Service - Backbill": backbill,
        "Market Securitization - Uplift Charge": uplift,
        "TX-ERCOT Admin Fees - CIL": tx_ercot,
        "Transmission Charges": transmission,
        "Taxes & PUC Assessment Charge": taxes,
        "Ancilliary Service Obligation Adjustment": ancillary,
        "Other Taxes": other_taxes,
        "Bill Total": bill_total,
    }

    line_items = []
    for col in BILL_COMPONENT_COLUMNS:
        amount = float(row[col])
        if amount == 0:
            continue
        qty = float(row["Usage - Actual KWH"]) if "Energy" in col or "ERCOT" in col or "Fuel" in col or "Securitization" in col or "Ancilliary" in col else 1.0
        rate = amount / qty if qty else amount
        line_items.append((col, qty, rate, amount))

    return row, line_items, f"Account: {account} | {note}"


def draw_invoice(pdf_path: Path, row: Dict[str, object], line_items: List[Tuple[str, float, float, float]], note: str, index: int) -> None:
    c = canvas.Canvas(str(pdf_path), pagesize=landscape(letter))
    width, height = landscape(letter)

    left = 0.55 * inch
    y = height - 0.45 * inch
    line_height = 10.2

    c.setFont("Courier-Bold", 15)
    c.drawString(left, y, "SYNTHETIC ELECTRICITY INVOICE")
    y -= 18
    c.setFont("Courier", 8.2)
    c.drawString(left, y, f"Demo ID: SYN-2026-{index + 1:04d} | {note}")
    y -= 16

    for col in OUTPUT_COLUMNS:
        value = row[col]
        if isinstance(value, date):
            text_value = value.strftime("%m/%d/%Y")
        elif col in ["Power Factor", "Load Factor"]:
            text_value = pct(float(value) * 100)
        elif col == "4CP Charges Rate ($/KW)":
            text_value = f"${float(value):.6f}"
        elif isinstance(value, (int, float)) and col in ["Actual Demand (KW)", "Billing Demand (KW)", "4CP Charges Qty (KW)", "Usage - Actual KWH"]:
            text_value = number(float(value))
        elif isinstance(value, (int, float)):
            text_value = money(float(value))
        else:
            text_value = str(value)

        c.drawString(left, y, f"{col}: {text_value}")
        y -= line_height

    y -= 8
    c.setFont("Courier-Bold", 9)
    c.drawString(left, y, "Line Items")
    y -= 12
    c.setFont("Courier", 7.4)

    for description, qty, rate, amount in line_items:
        c.drawString(
            left,
            y,
            f"Line Item: {description} | Qty: {qty:.2f} | Rate: ${rate:.6f} | Amount: ${amount:.2f}",
        )
        y -= 9.2
        if y < 36:
            c.showPage()
            c.setFont("Courier", 7.4)
            y = height - 0.5 * inch

    y -= 6
    c.setFont("Courier-Oblique", 8)
    c.drawString(left, y, "Synthetic/sample invoice only - no confidential data")

    c.save()


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # Remove old samples so the folder always has exactly 15 PDFs.
    for old_file in OUT_DIR.glob("*.pdf"):
        old_file.unlink()

    for i in range(15):
        row, line_items, note = build_invoice(i)
        name = f"invoice_{i + 1:02d}_synthetic_energy.pdf"
        if i == 13:
            name = "invoice_14_flag_bill_total_mismatch.pdf"
        elif i == 14:
            name = "invoice_15_flag_4cp_mismatch.pdf"
        draw_invoice(OUT_DIR / name, row, line_items, note, i)

    print(f"Created 15 sample invoices in: {OUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
