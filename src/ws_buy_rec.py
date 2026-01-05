# ws_buy_rec.py
# Wholesale Buy Recommendation Engine
# Python-native Excel generator for BUY/KILL/Prioritize decisions
# Author: John Wesley Quintero / WesAI

import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Protection

# --- CONFIG ---
BRAND = "SL"  # default brand
OUTPUT_DIR = "excel_templates"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "BUY_RECOMMENDATIONS.xlsx")
MAX_ROWS = 200  # pre-fill rows for formulas

# --- Tabs and Headers ---
TABS = {
    "BUY": [
        "ASIN", "AMZ Title", "Est. Qty", "Sell Price", "Current BSR", "Profit", "ROI",
        "Sales Opp.", "Suggested Pk", "Pack Qty", "Odr Qty (Pk)", "Item Code", "Description",
        "Cost", "Order Qty (Unit)", "Order Amount", "Var Weight", "Case Pack",
        "Offers 90d", "ROI/Profit", "Visible", "Var. Opp.", "Use Qty", "BSR 30 Avg",
        "% Chng30", "BSR 90 Avg", "% Chng90", "Avail. Qty", "AMZ Boss", "AMZ OOS", "Bbox OOS",
        "Brand", "Keepa Link", "AMZ Product Page", "Check Gated in SC", "Extra1", "Extra2", "Extra3"
    ],
    "RESEARCH": [
        "ASIN", "Product ID", "Title", "Package Quantity", "Brand", "Product Group",
        "Total Offers", "Sales Rank", "Estimated Number of Sales", "Profit", "Margin",
        "ROI", "Purchase Price", "Sell Price", "Buy Box Landed", "Seller Proceeds",
        "Low New Fba Price", "Low New Mfn Price", "Referral Fee", "Variable Closing Fee",
        "Fulfillment Subtotal", "Cost Sub Total", "VAT %", "VAT $", "Inbound Shipping Estimate",
        "Package Weight", "Package Height", "Package Length", "Package Width",
        "New FBA Num Offers", "New MFN Num Offers", "Keepa", "CamelCC", "Size", "Last Run",
        "Item Number", "PRODUCT", "SKU #", "Size(1)", "Case Pk", "Salon price / each",
        "Salon price / case", "QTY", "YOUR TOTAL COST", "CASE COUNT"
    ],
    "ORDER": ["Item Code", "Description", "Unit Cost", "Order Qty", "Total Amount"],
    "KEEPA": [
        "ASIN", "Baseline", "BB Avg", "AMZ Avg", "wtd avg", "Variations", "Review Wt", "Confidence",
        "ASIN_SRC", "Title", "Buy Box: Current", "Buy Box: 30 days avg.", "Buy Box: 30 days drop %",
        "Buy Box: 90 days avg.", "Buy Box: 90 days drop %", "Buy Box: 180 days avg.",
        "Buy Box Seller", "Amazon: Current", "Amazon: 30 days avg.", "Amazon: 30 days drop %",
        "Amazon: 90 days avg.", "Amazon: 90 days drop %", "Amazon: 180 days avg.",
        "Amazon out of stock percentage: 90 days OOS %", "Sales Rank: Current", "Sales Rank: 30 days avg.",
        "Sales Rank: 30 days drop %", "Sales Rank: 90 days avg.", "Sales Rank: 90 days drop %",
        "Sales Rank: 180 days avg.", "Count of retrieved live offers: New, FBA",
        "Count of retrieved live offers: New, FBM", "New Offer Count: Current", "New Offer Count: 90 days avg.",
        "New, 3rd Party FBA: 30 days avg.", "New, 3rd Party FBA: 90 days avg.", "New, 3rd Party FBA: 180 days avg.",
        "Variation ASINs", "Parent ASIN", "Reviews: Reviews - Format Specific", "Variation Attributes",
        "Reviews: Review Count", "Model", "Number of Items", "Categories: Root",
        "FBA Fees: Referral Fee %", "Buy Box out of stock percentage: 90 days OOS %"
    ],
    "IP Qty": [
        "Image", "Name", "SKU", "Barcode", "Cost Price", "Price", "Replenishment", "To Order",
        "Stock", "On Order", "Sales", "Adjusted Sales Velocity/mo", "Lead Time", "Days of Stock",
        "Vendors", "Brand", "AVG Retail Price", "Replenishable", "In Buy Sheet?"
    ]
}

# --- HELPER FUNCTION ---
def create_tab(ws, headers):
    ws.append(headers)
    for col_num in range(1, len(headers)+1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.protection = Protection(locked=True)

# --- INIT WORKBOOK ---
wb = Workbook()
# Remove default sheet
default_sheet = wb.active
wb.remove(default_sheet)

# Create tabs
for tab_name, headers in TABS.items():
    ws = wb.create_sheet(title=tab_name)
    create_tab(ws, headers)

# --- BUY TAB FORMULA INJECTION ---
buy_ws = wb["BUY"]
for row in range(2, MAX_ROWS+2):
    # Profit = Sell Price - Cost
    buy_ws[f"F{row}"] = f"=D{row}-N{row}"
    # ROI = Profit / Cost
    buy_ws[f"G{row}"] = f"=IF(N{row}=0,0,F{row}/N{row})"
    # Days of Supply = Avail Qty / Avg Daily Sales placeholder
    buy_ws[f"L{row}"] = f"=IF(Y{row}=0,0,Z{row}/Y{row})"
    # Max BSR 30 & % Change
    buy_ws[f"X{row}"] = f"=MAX([@[BSR 30 Avg]],C{row})"  # placeholder example
    buy_ws[f"Y{row}"] = f"=IF(C{row}=0,0,(C{row}-X{row})/C{row})"
    # Max BSR 90 & % Change
    buy_ws[f"Z{row}"] = f"=MAX([@[BSR 90 Avg]],C{row})"
    buy_ws[f"AA{row}"] = f"=IF(C{row}=0,0,(C{row}-Z{row})/C{row})"
    # Var Opp
    buy_ws[f"AB{row}"] = f"=IF(D{row}=0,0,F{row}/D{row})"
    # Recommended Order Qty = desired coverage logic
    buy_ws[f"O{row}"] = f"=MAX(0,(30*Y{row})-P{row})"

# --- SAVE WORKBOOK ---
wb.save(OUTPUT_FILE)
print(f"✅ BUY recommendation workbook generated: {OUTPUT_FILE}")
