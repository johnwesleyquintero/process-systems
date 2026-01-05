# ws_buy_rec.py
# Wholesale Buy Recommendation Engine
# Python-native Excel generator for BUY/KILL/Prioritize decisions
# Author: John Wesley Quintero / WesAI
#
# --- LEGACY REVERSE ENGINEERING NOTES ---
# 1. UI/UX: 
#    - Frozen Panes: Standardized to row 4 (headers) and col B (ASIN/Title) for BUY/RESEARCH/KEEPA.
#    - AutoFilters: Enabled on all header rows to allow quick filtering by ASIN, Profit, etc.
#    - Column Widths: Heuristic-based (Title=55, ASIN=16, Numbers=12-14) to replicate legacy "pretty" look.
# 2. Logic:
#    - BUY Tab: Core decision sheet using VLOOKUPs to KEEPA and RESEARCH.
#    - KEEPA Tab: Implements weighted BB/AMZ averages. (30d=1, 90d=3, 180d=1 weights).
#    - IP Qty: Added 'In Buy Sheet?' check to cross-reference inventory planning with current review.
#    - ORDER: Dynamically pulls ASIN/Qty from BUY tab where Order Qty > 0.
# 3. Improvements:
#    - Removed sheet protection for unrestricted editing.
#    - Dynamic ranges (A:A) for VLOOKUPs to prevent breakages on large datasets.
#    - NA() and IFERROR() handling for cleaner mathematical outputs.

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Protection
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- CONFIG ---
OUTPUT_DIR = "excel_templates"
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "BUY_RECOMMENDATIONS.xlsx")
MAX_ROWS = 200  # pre-fill rows for formulas

# --- STYLES ---
HEADER_FILL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
SUMMARY_LABEL_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
SUMMARY_VALUE_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
BOLD_FONT = Font(bold=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
WRAPPED_CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN_WRAPPED = Alignment(horizontal="left", vertical="top", wrap_text=True)

# Protection Styles
LOCKED = Protection(locked=True)
UNLOCKED = Protection(locked=False)

# Conditional Formatting Styles
RED_FONT = Font(color="FF9C0006")
RED_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
GREEN_FONT = Font(color="FF006100")
GREEN_FILL = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
ORANGE_FONT = Font(color="FF9C6500")
ORANGE_FILL = PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid")

# --- Tabs and Headers ---
TABS_CONFIG = {
    "README": {
        "header_row": 1,
        "headers": ["Section", "Description / Logic"],
        "content": [
            ["AUTHOR", "John Wesley Quintero"],
            ["USER GUIDE", "This workbook is used to calculate and recommend wholesale buys based on Amazon and Keepa data."],
            ["BUY Tab", "Main decision sheet. Enter ASINs here to see order recommendations."],
            ["AZInsight_Data Tab", "Raw data from AZInsight research tool. Paste your Amazon research exports here."],
            ["KEEPA Tab", "Historical price and rank data. Paste Keepa data exports here."],
            ["IP Qty Tab", "Inventory Planning data. Used for barcode lookups and stock levels."],
            ["", ""],
            ["CORE CALCULATIONS", "Key formulas used in this template:"],
            ["Weighted BB Avg", "((30d Avg * 30wt) + (90d Avg * 90wt) + (180d Avg * 180wt)) / Total Weights"],
            ["wtd avg", "If AMZ OOS >= 65%, use BB Avg. Otherwise, use MIN(BB Avg, AMZ Avg * (1 - AMZ Less))."],
            ["Var. Opp.", "Estimated Sales * Variation Weight * (Target Days / 30). Adjusted by AMZ Boss multiplier if applicable."],
            ["ROI/Profit", "Highlights 'OK' if ROI >= Min ROI AND Profit >= Min Profit thresholds."],
            ["% Chng30/90", "(Avg BSR - Current BSR) / Avg BSR. Positive means rank is improving."],
        ]
    },
    "BUY": {
        "summary_labels": {
            "A1": "Total Items", "C1": "Min Est Qty", "H1": "Min Sell Price", "I1": "Max BSR", 
            "L1": "3PL / Unit", "O1": "Total Profit", "P1": "Overall ROI", "Q1": "Total Sales", 
            "U1": "Max Order Qty", "V1": "Days of Stock", "W1": "Supplier Commission", "X1": "Discount", 
            "Y1": "Total Units", "Z1": "Total Amount", "AA1": "AMZ Less", "AB1": "Case Pk Default", 
            "AC1": "Min ROI", "AD1": "25%", "AE1": "20%", "AI1": "Max BSR 30", "AJ1": "Max Chng 30", 
            "AK1": "Max BSR 90", "AL1": "Max Chng 90", "AM1": "Min Est Qty", "AN1": "Var Opp", 
            "AO1": "Max AMZ OOS", "AP1": "Max Bbox OOS"
        },
        "summary_values": {
            "B1": 0, "D1": 30, "H2": 9.00, "I2": 100000, "L2": 0.00, "O2": 0.00, "P2": 0.00, 
            "Q2": 0.00, "U2": 240, "V2": 60, "W2": 0.00, "X2": 0.00, "Y2": 0, "Z2": 0.00, 
            "AA2": 0.00, "AB2": 12, "AC2": 0.05, "AD2": 0.25, "AE2": 0.20, "AF2": "Min Profit", 
            "AG2": 1.50, "AH2": 2.00, "AI2": 125000, "AJ2": 0.50, "AK2": 150000, "AL2": 0.75, 
            "AM2": 300, "AN2": 0.07, "AO2": 0.35, "AP2": 0.50
        },
        "headers": [
            "ASIN", "AMZ Title", "Est. Qty", "New FBA", "New MFN", "Total Offers", "Sell Price", 
            "Current BSR", "Referral %", "S & H", "Proceeds", "Profit", "ROI", "Sales Opp.", 
            "Suggested Pk", "Pack Qty", "Odr Qty (Pk)", "Item Code", "Description", "Cost", 
            "Order Qty (Unit)", "Order Amount", "Var Weight", "Case Pack", "IP Qty", "Stock", 
            "On Order", "Velocity", "Units Sold", "Offers 90d", "ROI/Profit", "Visible", 
            "Var. Opp.", "Use Qty", "BSR 30 Avg", "% Chng30", "BSR 90 Avg", "% Chng90", 
            "Avail. Qty", "AMZ Boss", "AMZ OOS", "Bbox OOS", "Brand", "Keepa Link", 
            "AMZ Product Page", "Check Gated in SC", "Extra1", "Extra2", "Extra3"
        ],
        "header_row": 3
    },
    "AZInsight_Data": {
        "summary_labels": {
            "A1": "Total Items", "C1": "Matches Found on AMZ", "E1": "Not Found on AMZ"
        },
        "summary_values": {
            "B1": f"=COUNTA(A4:A{3 + MAX_ROWS})", 
            "D1": f"=COUNTIF(A4:A{3 + MAX_ROWS}, \"*?\")", 
            "F1": "=B1-D1"
        },
        "headers": [
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
        "header_row": 3
    },
    "ORDER": {
        "headers": ["Item Code", "Description", "Unit Cost", "Order Qty", "Total Amount"],
        "header_row": 1
    },
    "KEEPA": {
        "summary_labels": {
            "B1": "30 wt", "C1": "90 wt", "D1": "180 wt", "F1": "AMZ Less:"
        },
        "summary_values": {
            "B2": 1, "C2": 3, "D2": 1, "F2": 0.05
        },
        "headers": [
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
            "Reviews: Review Count", "Model", "Number of Items", "Categories: Root", "FBA Fees:",
            "Referral Fee %", "Buy Box out of stock percentage: 90 days OOS %"
        ],
        "header_row": 3
    },
    "IP Qty": {
        "headers": [
            "Image", "Name", "SKU", "Barcode", "Cost Price", "Price", "Replenishment", "To Order",
            "Stock", "On Order", "Sales", "Adjusted Sales Velocity/mo", "Lead Time", "Days of Stock",
            "Vendors", "Brand", "AVG Retail Price", "Replenishable", "In Buy Sheet?"
        ],
        "header_row": 1
    }
}

# --- Column Mapping Cache ---
COLUMN_MAPS = {}

def get_col(tab_name, header_name):
    """
    Returns the Excel column letter for a given header name within a specific tab.
    Uses a cache for performance and safety.
    """
    if tab_name not in TABS_CONFIG:
        raise ValueError(f"Tab '{tab_name}' not found in TABS_CONFIG.")
    
    if tab_name not in COLUMN_MAPS:
        headers = TABS_CONFIG[tab_name]["headers"]
        COLUMN_MAPS[tab_name] = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}
    
    if header_name not in COLUMN_MAPS[tab_name]:
        raise KeyError(f"Header '{header_name}' not found in tab '{tab_name}'. Available: {list(COLUMN_MAPS[tab_name].keys())}")
    
    return COLUMN_MAPS[tab_name][header_name]

# --- HELPER FUNCTIONS ---
def setup_sheet(ws, tab_config):
    """
    Sets up headers, formatting, filters, and frozen panes for a worksheet.
    """
    headers = tab_config["headers"]
    header_row = tab_config["header_row"]
    
    # 1. Write Headers
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = BOLD_FONT
        cell.alignment = WRAPPED_CENTER_ALIGN
        
        # Apply Column Widths (Reverse Engineered / Heuristic)
        col_letter = get_column_letter(col_idx)
        h_lower = header.lower()
        if ws.title == "README":
            ws.column_dimensions[col_letter].width = 30 if col_idx == 1 else 100
        elif any(kw in h_lower for kw in ["asin", "code", "sku", "barcode"]):
            ws.column_dimensions[col_letter].width = 16
        elif any(kw in h_lower for kw in ["title", "description", "name"]):
            ws.column_dimensions[col_letter].width = 55
        elif any(kw in h_lower for kw in ["link", "url", "page"]):
            ws.column_dimensions[col_letter].width = 10
        elif any(kw in h_lower for kw in ["qty", "bsr", "rank", "count", "offers", "sales", "stock", "items", "case count"]):
            ws.column_dimensions[col_letter].width = 12
        elif any(kw in h_lower for kw in ["price", "cost", "amount", "profit", "proceeds", "fee", "subtotal", "total cost", "roi", "margin"]):
            ws.column_dimensions[col_letter].width = 14
        else:
            ws.column_dimensions[col_letter].width = 15

        # Apply default column formatting based on header name
        fmt = 'General'
        if any(kw in h_lower for kw in ["price", "cost", "amount", "profit", "proceeds", "fee", "subtotal", "total cost"]):
            fmt = '"$"#,##0.00'
        elif any(kw in h_lower for kw in ["roi", "margin", "%", "chng", "weight", "oos"]):
            fmt = '0.00%'
        elif any(kw in h_lower for kw in ["qty", "bsr", "rank", "count", "offers", "sales", "stock", "items", "case count"]):
            fmt = '#,##0'
        elif h_lower == "roi/profit":
            fmt = '"OK";;"Not OK"'
            
        for row_idx in range(header_row + 1, header_row + MAX_ROWS + 1):
            ws.cell(row=row_idx, column=col_idx).number_format = fmt

    # 2. Enable AutoFilter on the header row
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A{header_row}:{last_col}{header_row + MAX_ROWS}"

    # 3. Cell Protection Strategy
    # By default, openpyxl cells are locked. We need to unlock columns meant for user input.
    if ws.title != "README":
        input_keywords = ["asin", "title", "cost", "case pack", "extra", "product id", "qty", "sku", "barcode", "replenishment", "lead time"]
        for col_idx, header in enumerate(headers, 1):
            h_lower = header.lower()
            # If it's a known input field, unlock the whole data range
            if any(kw in h_lower for kw in input_keywords) and "avg" not in h_lower and "oos" not in h_lower:
                for row_idx in range(header_row + 1, header_row + MAX_ROWS + 1):
                    ws.cell(row=row_idx, column=col_idx).protection = UNLOCKED
            else:
                # Explicitly lock others (just to be safe)
                for row_idx in range(header_row + 1, header_row + MAX_ROWS + 1):
                    ws.cell(row=row_idx, column=col_idx).protection = LOCKED

    # 4. Freeze Panes (Standardized for Wholesale Review)
    if ws.title in ["BUY", "AZInsight_Data", "KEEPA"]:
        # Freeze top rows (up to header) and first column (usually ASIN)
        ws.freeze_panes = ws.cell(row=header_row + 1, column=2)
    elif ws.title == "IP Qty":
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    # 4. Summary Labels and Values (if any)
    if "summary_labels" in tab_config:
        for cell_ref, label in tab_config["summary_labels"].items():
            cell = ws[cell_ref]
            cell.value = label
            cell.fill = SUMMARY_LABEL_FILL
            cell.font = BOLD_FONT
            cell.alignment = CENTER_ALIGN
            
    if "summary_values" in tab_config:
        for cell_ref, val in tab_config["summary_values"].items():
            cell = ws[cell_ref]
            cell.value = val
            cell.fill = SUMMARY_VALUE_FILL
            cell.alignment = CENTER_ALIGN
            # Summary values are usually locked unless they are specific inputs like weights
            if ws.title in ["BUY", "KEEPA"] and cell_ref in ["B2", "C2", "D2", "F2", "L2", "Q2", "R2", "S2", "T2"]:
                cell.protection = UNLOCKED
            else:
                cell.protection = LOCKED
                
            # Format numeric summary values
            if isinstance(val, (int, float)):
                if val < 1 and val > 0: # likely a percentage
                    cell.number_format = '0.00%'
                elif val > 1000:
                    cell.number_format = '#,##0'
                else:
                    cell.number_format = '0.00'

    # 6. Enable Sheet Protection with password
    ws.protection.sheet = True
    ws.protection.password = "wesai"
    # Allow filtering even when protected
    ws.protection.autoFilter = False # In openpyxl, setting to False actually ALLOWS it if it was already set

# --- INIT WORKBOOK ---
wb = Workbook()
# Remove default sheet
default_sheet = wb.active
wb.remove(default_sheet)

# Create tabs
for tab_name, config in TABS_CONFIG.items():
    ws = wb.create_sheet(title=tab_name)
    setup_sheet(ws, config)
    
    # Fill README content
    if tab_name == "README":
        for i, (section, desc) in enumerate(config["content"], 2):
            ws[f"A{i}"] = section
            ws[f"B{i}"] = desc
            ws[f"A{i}"].font = BOLD_FONT if section.isupper() else Font(bold=False)
            ws[f"B{i}"].alignment = LEFT_ALIGN_WRAPPED
            if section.isupper():
                ws[f"A{i}"].fill = SUMMARY_LABEL_FILL
                ws[f"B{i}"].fill = SUMMARY_LABEL_FILL

# --- FORMULA REFERENCES ---
# (Cache initialized on first get_col call)

# --- BUY TAB FORMULA INJECTION ---
buy_ws = wb["BUY"]
header_row_buy = TABS_CONFIG["BUY"]["header_row"]
start_row = header_row_buy + 1

# Map BUY headers to columns
b_asin = get_col("BUY", 'ASIN')
b_amz_title = get_col("BUY", 'AMZ Title')
b_est_qty = get_col("BUY", 'Est. Qty')
b_fba = get_col("BUY", 'New FBA')
b_mfn = get_col("BUY", 'New MFN')
b_offers = get_col("BUY", 'Total Offers')
b_sell = get_col("BUY", 'Sell Price')
b_bsr = get_col("BUY", 'Current BSR')
b_referral = get_col("BUY", 'Referral %')
b_sh = get_col("BUY", 'S & H')
b_proceeds = get_col("BUY", 'Proceeds')
b_profit = get_col("BUY", 'Profit')
b_roi = get_col("BUY", 'ROI')
b_opp = get_col("BUY", 'Sales Opp.')
b_sugg_pk = get_col("BUY", 'Suggested Pk')
b_pk_qty = get_col("BUY", 'Pack Qty')
b_odr_pk = get_col("BUY", 'Odr Qty (Pk)')
b_item_code = get_col("BUY", 'Item Code')
b_desc = get_col("BUY", 'Description')
b_cost = get_col("BUY", 'Cost')
b_units = get_col("BUY", 'Order Qty (Unit)')
b_amount = get_col("BUY", 'Order Amount')
b_var_weight = get_col("BUY", 'Var Weight')
b_case_pack = get_col("BUY", 'Case Pack')
b_ip_qty = get_col("BUY", 'IP Qty')
b_stock = get_col("BUY", 'Stock')
b_on_order = get_col("BUY", 'On Order')
b_velocity = get_col("BUY", 'Velocity')
b_units_sold = get_col("BUY", 'Units Sold')
b_offers_90d = get_col("BUY", 'Offers 90d')
b_roi_profit = get_col("BUY", 'ROI/Profit')
b_visible = get_col("BUY", 'Visible')
b_var_opp = get_col("BUY", 'Var. Opp.')
b_use_qty = get_col("BUY", 'Use Qty')
b_bsr30 = get_col("BUY", 'BSR 30 Avg')
b_chng30 = get_col("BUY", '% Chng30')
b_bsr90 = get_col("BUY", 'BSR 90 Avg')
b_chng90 = get_col("BUY", '% Chng90')
b_avail = get_col("BUY", 'Avail. Qty')
b_amz_boss = get_col("BUY", 'AMZ Boss')
b_amz_oos = get_col("BUY", 'AMZ OOS')
b_bbox_oos = get_col("BUY", 'Bbox OOS')
b_brand = get_col("BUY", 'Brand')
b_keepa_link = get_col("BUY", 'Keepa Link')
b_amz_link = get_col("BUY", 'AMZ Product Page')
b_gated = get_col("BUY", 'Check Gated in SC')

# Lookup Indices for AZInsight_Data
res_h = TABS_CONFIG["AZInsight_Data"]["headers"]
r_sell_price_idx = res_h.index("Sell Price") + 1
r_bsr_idx = res_h.index("Sales Rank") + 1
r_cost_idx = res_h.index("Purchase Price") + 1
r_opp_idx = res_h.index("Estimated Number of Sales") + 1
r_brand_idx = res_h.index("Brand") + 1
r_title_idx = res_h.index("Title") + 1
r_fba_idx = res_h.index("Low New Fba Price") + 1
r_mfn_idx = res_h.index("Low New Mfn Price") + 1
r_offers_idx = res_h.index("Total Offers") + 1
r_referral_idx = res_h.index("Referral Fee") + 1
r_proceeds_idx = res_h.index("Seller Proceeds") + 1
r_sh_idx = res_h.index("Inbound Shipping Estimate") + 1
r_case_pk_idx = res_h.index("Case Pk") + 1
r_item_num_idx = res_h.index("Item Number") + 1
r_size_idx = res_h.index("Size") + 1

# Lookup Indices for KEEPA
kee_h = TABS_CONFIG["KEEPA"]["headers"]
k_bsr30_idx = kee_h.index("Sales Rank: 30 days avg.") + 1
k_bsr90_idx = kee_h.index("Sales Rank: 90 days avg.") + 1
k_offers90_idx = kee_h.index("New Offer Count: 90 days avg.") + 1
k_amz_oos_idx = kee_h.index("Amazon out of stock percentage: 90 days OOS %") + 1
k_bbox_oos_idx = kee_h.index("Buy Box out of stock percentage: 90 days OOS %") + 1
k_baseline_idx = kee_h.index("Baseline") + 1
k_referral_idx = kee_h.index("Referral Fee %") + 1

# Lookup Indices for IP Qty
ip_h = TABS_CONFIG["IP Qty"]["headers"]
ip_barcode_idx = ip_h.index("Barcode") + 1
ip_to_order_idx = ip_h.index("To Order") + 1
ip_stock_idx = ip_h.index("Stock") + 1
ip_on_order_idx = ip_h.index("On Order") + 1
ip_sales_idx = ip_h.index("Sales") + 1
ip_velocity_idx = ip_h.index("Adjusted Sales Velocity/mo") + 1

for row in range(start_row, start_row + MAX_ROWS):
    asin_ref = f"{b_asin}{row}"
    
    # AZInsight_Data Tab Lookups (Wrapped in IFERROR)
    res_range = "AZInsight_Data!$A:$AS"
    buy_ws[f"{b_est_qty}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_opp_idx}, FALSE), 0)'
    buy_ws[f"{b_sell}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_sell_price_idx}, FALSE), 0)'
    buy_ws[f"{b_bsr}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_bsr_idx}, FALSE), 0)'
    buy_ws[f"{b_cost}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_cost_idx}, FALSE), 0)'
    
    # AMZ Title Lookup from AZInsight_Data Tab
    buy_ws[f"{b_amz_title}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_title_idx}, FALSE), "")'

    # Financial & Offer Lookups from AZInsight_Data Tab
    buy_ws[f"{b_fba}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_fba_idx}, FALSE), 0)'
    buy_ws[f"{b_mfn}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_mfn_idx}, FALSE), 0)'
    buy_ws[f"{b_offers}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_offers_idx}, FALSE), 0)'
    buy_ws[f"{b_proceeds}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_proceeds_idx}, FALSE), 0)'
    buy_ws[f"{b_sh}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_sh_idx}, FALSE), 0)'
    buy_ws[f"{b_sugg_pk}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_case_pk_idx}, FALSE), 1)'
    buy_ws[f"{b_item_code}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_item_num_idx}, FALSE), "")'
    buy_ws[f"{b_case_pack}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_case_pk_idx}, FALSE), 1)'
    buy_ws[f"{b_avail}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_size_idx}, FALSE), 0)'
    buy_ws[f"{b_desc}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_title_idx}, FALSE), "")'

    # Referral % from KEEPA (more accurate for percentage)
    buy_ws[f"{b_referral}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, KEEPA!$A:$AW, {k_referral_idx}, FALSE), 0)'

    # Pack Qty Logic
    buy_ws[f"{b_pk_qty}{row}"] = f'=IF({b_sugg_pk}{row}>0, {b_sugg_pk}{row}, 1)'

    # Profit & ROI (Wrapped in IFERROR)
    # Legacy Profit: Proceeds - Cost - (S & H)
    buy_ws[f"{b_profit}{row}"] = f"=IFERROR({b_proceeds}{row}-{b_cost}{row}-{b_sh}{row}, 0)"
    buy_ws[f"{b_roi}{row}"] = f"=IFERROR(IF({b_cost}{row}=0, 0, {b_profit}{row}/{b_cost}{row}), 0)"

    # KEEPA Tab Lookups (Wrapped in IFERROR)
    kee_range = "KEEPA!$A:$AW"
    buy_ws[f"{b_bsr30}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {kee_range}, {k_bsr30_idx}, FALSE), 0)'
    buy_ws[f"{b_bsr90}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {kee_range}, {k_bsr90_idx}, FALSE), 0)'
    buy_ws[f"{b_offers_90d}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {kee_range}, {k_offers90_idx}, FALSE), 0)'
    buy_ws[f"{b_amz_oos}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {kee_range}, {k_amz_oos_idx}, FALSE), 0)'
    buy_ws[f"{b_bbox_oos}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {kee_range}, {k_bbox_oos_idx}, FALSE), 0)'

    # IP Qty Tab Lookups (Using ASIN as lookup value per legacy logic)
    ip_range = "IP Qty!$D:$R" # Barcode is column D
    buy_ws[f"{b_ip_qty}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {ip_range}, {ip_to_order_idx - ip_barcode_idx + 1}, FALSE), 0)'
    buy_ws[f"{b_stock}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {ip_range}, {ip_stock_idx - ip_barcode_idx + 1}, FALSE), 0)'
    buy_ws[f"{b_on_order}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {ip_range}, {ip_on_order_idx - ip_barcode_idx + 1}, FALSE), 0)'
    buy_ws[f"{b_velocity}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {ip_range}, {ip_velocity_idx - ip_barcode_idx + 1}, FALSE), 0)'
    buy_ws[f"{b_units_sold}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {ip_range}, {ip_sales_idx - ip_barcode_idx + 1}, FALSE), 0)'

    # Trends (Wrapped in IFERROR)
    buy_ws[f"{b_chng30}{row}"] = f"=IFERROR(IF({b_bsr30}{row}=0, 0, ({b_bsr30}{row}-{b_bsr}{row})/{b_bsr}{row}), 0)"
    buy_ws[f"{b_chng90}{row}"] = f"=IFERROR(IF({b_bsr90}{row}=0, 0, ({b_bsr90}{row}-{b_bsr}{row})/{b_bsr}{row}), 0)"

    # Sales Opp (Legacy: Est Qty * (Days Stock / 30))
    stock_days = "$V$2"
    buy_ws[f"{b_opp}{row}"] = f"=IFERROR({b_est_qty}{row}*({stock_days}/30), 0)"

    # Visible (Legacy: If ROI/Profit > 0, 1, 0)
    buy_ws[f"{b_visible}{row}"] = f"=IF({b_roi_profit}{row}>0, 1, 0)"

    # Var. Opp. (Legacy logic)
    var_opp_formula = f"IF(ISNUMBER({b_var_weight}{row}), IF({b_amz_boss}{row}, {b_opp}{row}*{b_var_weight}{row}*$AN$2, {b_opp}{row}*{b_var_weight}{row}), NA())"
    buy_ws[f"{b_var_opp}{row}"] = f"=IFERROR({var_opp_formula}, NA())"

    # Use Qty (Legacy logic)
    max_order = "$U$2"
    buy_ws[f"{b_use_qty}{row}"] = f"=IFERROR(MIN({b_var_opp}{row}, {max_order}), 0)"

    # Order Qty (Unit)
    buy_ws[f"{b_units}{row}"] = f"={b_use_qty}{row}"
    
    # Order Amount
    buy_ws[f"{b_amount}{row}"] = f"={b_cost}{row}*{b_units}{row}"
    
    # Odr Qty (Pk)
    buy_ws[f"{b_odr_pk}{row}"] = f"=IFERROR({b_units}{row}/{b_pk_qty}{row}, 0)"

    # Brand Lookup from AZInsight_Data Tab (Wrapped in IFERROR)
    buy_ws[f"{b_brand}{row}"] = f'=IFERROR(VLOOKUP({asin_ref}, {res_range}, {r_brand_idx}, FALSE), "")'

    # Hyperlinks
    buy_ws[f"{b_keepa_link}{row}"] = f'=IF(OR({asin_ref}="", {asin_ref}=0), "", HYPERLINK("https://keepa.com/#!product/1-"&{asin_ref}, "Keepa"))'
    buy_ws[f"{b_amz_link}{row}"] = f'=IF(OR({asin_ref}="", {asin_ref}=0), "", HYPERLINK("https://www.amazon.com/dp/"&{asin_ref}, "Amazon"))'
    buy_ws[f"{b_gated}{row}"] = f'=IF(OR({asin_ref}="", {asin_ref}=0), "", HYPERLINK("https://sellercentral.amazon.com/product-search/search?q="&{asin_ref}, "Check Gated"))'

    # AMZ Boss (Legacy logic)
    # AND(VLOOKUP(...)<=$AO$2, Est. Qty >= $AM$2)
    buy_ws[f"{b_amz_boss}{row}"] = f'=IFERROR(AND(VLOOKUP({asin_ref}, KEEPA!$A:$AW, {k_baseline_idx}, FALSE)<=$AO$2, {b_opp}{row} >= $AM$2), FALSE)'

    # ROI/Profit (Legacy Complex Logic)
    # Tiers based on Summary Values
    tier_logic = f"((({b_roi}{row}>=$AD$2)*({b_profit}{row}>=$AH$2)) + (({b_roi}{row}<$AD$2)*({b_roi}{row}>=$AE$2)*({b_profit}{row}>=$AG$2)) + (({b_roi}{row}<$AE$2)*({b_roi}{row}>=$AC$2)*({b_profit}{row}>=0)))"
    bsr_logic = f"IF($AI$2=0, 1, {b_bsr30}{row}<=$AI$2) * IF($AK$2=0, 1, {b_bsr90}{row}<=$AK$2)"
    other_filters = f"AND({b_opp}{row}>=$D$1, {b_bsr}{row}<=$I$2, {b_sell}{row}>=$H$2)"
    buy_ws[f"{b_roi_profit}{row}"] = f"=IF({b_pk_qty}{row}>1, ({tier_logic}*2 * {bsr_logic} * {other_filters}), ({tier_logic} * {bsr_logic} * {other_filters}))"

# --- BUY TAB CONDITIONAL FORMATTING & VALIDATION ---
# 1. ASIN Duplicate highlight
buy_ws.conditional_formatting.add(
    f"{b_asin}{start_row}:{b_asin}{start_row + MAX_ROWS}",
    Rule(type="duplicateValues", dxf=DifferentialStyle(fill=RED_FILL, font=RED_FONT))
)

# 2. AMZ Title Keywords (Frozen, Cool Ship, Spray)
b_title_col = get_col("BUY", "AMZ Title")
for kw in ["Frozen", "Cool Ship", "Spray"]:
    # Use relative reference for CF formula as suggested
    formula = f'NOT(ISERROR(SEARCH("{kw}",{b_title_col}{start_row})))'
    buy_ws.conditional_formatting.add(
        f"{b_title_col}{start_row}:{b_title_col}{start_row + MAX_ROWS}",
        Rule(type="expression", formula=[formula], dxf=DifferentialStyle(font=RED_FONT))
    )

# 3. Stock Low highlight (< 3)
buy_ws.conditional_formatting.add(
    f"{b_stock}{start_row}:{b_stock}{start_row + MAX_ROWS}",
    Rule(type="cellIs", operator='lessThan', formula=['3'], dxf=DifferentialStyle(fill=RED_FILL, font=RED_FONT))
)

# 4. Data Validation for "Use Qty" (Apply to whole range as suggested)
dv = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
buy_ws.add_data_validation(dv)
dv.add(f"{b_amz_boss}{start_row}:{b_amz_boss}{start_row + MAX_ROWS}") # Fixed range and column

# --- KEEPA TAB FORMULA INJECTION ---
keepa_ws = wb["KEEPA"]
start_row_keepa = TABS_CONFIG["KEEPA"]["header_row"] + 1

# Weights for wtd avg
w30 = "$B$2"
w90 = "$C$2"
w180 = "$D$2"
amz_less = "$F$2"

# Map KEEPA headers to columns
k_asin = get_col("KEEPA", "ASIN")
k_baseline = get_col("KEEPA", "Baseline")
k_bb_avg = get_col("KEEPA", "BB Avg")
k_amz_avg = get_col("KEEPA", "AMZ Avg")
k_wtd_avg = get_col("KEEPA", "wtd avg")
k_vars = get_col("KEEPA", "Variations")
k_rev_wt = get_col("KEEPA", "Review Wt")
k_conf = get_col("KEEPA", "Confidence")

# Data columns from KEEPA tab
k_bb_30 = get_col("KEEPA", "Buy Box: 30 days avg.")
k_bb_90 = get_col("KEEPA", "Buy Box: 90 days avg.")
k_bb_180 = get_col("KEEPA", "Buy Box: 180 days avg.")
k_bb_curr = get_col("KEEPA", "Buy Box: Current")

k_amz_30 = get_col("KEEPA", "Amazon: 30 days avg.")
k_amz_90 = get_col("KEEPA", "Amazon: 90 days avg.")
k_amz_180 = get_col("KEEPA", "Amazon: 180 days avg.")
k_amz_curr = get_col("KEEPA", "Amazon: Current")
k_amz_oos = get_col("KEEPA", "Amazon out of stock percentage: 90 days OOS %")

k_var_asins = get_col("KEEPA", "Variation ASINs")
k_parent = get_col("KEEPA", "Parent ASIN")
k_rev_fmt = get_col("KEEPA", "Reviews: Reviews - Format Specific")

for row in range(start_row_keepa, start_row_keepa + MAX_ROWS):
    # Baseline
    keepa_ws[f"{k_baseline}{row}"] = f'=IF(ISBLANK({k_amz_oos}{row}),0,IF({k_amz_oos}{row}>=0.65, "BB", "AMZ"))'
    
    # Robust Weighted Average using SUMPRODUCT-style logic (handles blanks/zeros)
    # BB Avg
    bb_num = f"SUM(IFERROR({k_bb_30}{row}*{w30},0), IFERROR({k_bb_90}{row}*{w90},0), IFERROR({k_bb_180}{row}*{w180},0))"
    bb_den = f"SUM(IF({k_bb_30}{row}>0,{w30},0), IF({k_bb_90}{row}>0,{w90},0), IF({k_bb_180}{row}>0,{w180},0))"
    keepa_ws[f"{k_bb_avg}{row}"] = f"=IF(AND({k_bb_30}{row}<{k_bb_90}{row},{k_bb_30}{row}<{k_bb_180}{row}), MIN({k_bb_30}{row},{k_bb_curr}{row}), IFERROR({bb_num}/{bb_den}, 0))"
    
    # AMZ Avg
    amz_num = f"SUM(IFERROR({k_amz_30}{row}*{w30},0), IFERROR({k_amz_90}{row}*{w90},0), IFERROR({k_amz_180}{row}*{w180},0))"
    amz_den = f"SUM(IF({k_amz_30}{row}>0,{w30},0), IF({k_amz_90}{row}>0,{w90},0), IF({k_amz_180}{row}>0,{w180},0))"
    keepa_ws[f"{k_amz_avg}{row}"] = f"=IF(AND({k_amz_30}{row}<{k_amz_90}{row},{k_amz_30}{row}<{k_amz_180}{row}), MIN({k_amz_30}{row},{k_amz_curr}{row}), IFERROR({amz_num}/{amz_den}, 0))"
    
    # wtd avg
    keepa_ws[f"{k_wtd_avg}{row}"] = f'=IF({k_baseline}{row}="BB",{k_bb_avg}{row},MIN({k_bb_avg}{row},{k_amz_avg}{row}*(1-{amz_less})))'
    
    # Variations count
    keepa_ws[f"{k_vars}{row}"] = f'=IF(ISBLANK({k_var_asins}{row}),1,LEN({k_var_asins}{row})-LEN(SUBSTITUTE({k_var_asins}{row},",",""))+1)'
    
    # Review Weight (simplified legacy logic)
    keepa_ws[f"{k_rev_wt}{row}"] = f'=IF({k_vars}{row}=1,1,IF(OR(ISBLANK({k_rev_fmt}{row}),{k_rev_fmt}{row}=0),"Manual",{k_rev_fmt}{row}/MAX(SUMIF($AK:$AK,{k_parent}{row},$AL:$AL),1)))'

    # Confidence (Using 1/0 for downstream math as suggested)
    keepa_ws[f"{k_conf}{row}"] = f"=IF({k_vars}{row}=1, 1, 0)"

# --- AZInsight_Data TAB CONDITIONAL FORMATTING ---
res_ws = wb["AZInsight_Data"]
start_row_res = TABS_CONFIG["AZInsight_Data"]["header_row"] + 1
r_rank = get_col("AZInsight_Data", "Sales Rank")
r_sales = get_col("AZInsight_Data", "Estimated Number of Sales")
r_profit = get_col("AZInsight_Data", "Profit")
r_margin = get_col("AZInsight_Data", "Margin")
r_roi = get_col("AZInsight_Data", "ROI")

# 1. Sales Rank highlight (Orange if 100k-200k)
res_ws.conditional_formatting.add(
    f"{r_rank}{start_row_res}:{r_rank}{start_row_res + MAX_ROWS}",
    Rule(type="cellIs", operator='between', formula=['100001', '200000'], dxf=DifferentialStyle(fill=ORANGE_FILL, font=ORANGE_FONT))
)

# 2. Estimated Sales highlight (Orange if 21-59)
res_ws.conditional_formatting.add(
    f"{r_sales}{start_row_res}:{r_sales}{start_row_res + MAX_ROWS}",
    Rule(type="cellIs", operator='between', formula=['21', '59'], dxf=DifferentialStyle(fill=ORANGE_FILL, font=ORANGE_FONT))
)

# 3. Profit highlight (Orange if 0.01-3.99)
res_ws.conditional_formatting.add(
    f"{r_profit}{start_row_res}:{r_profit}{start_row_res + MAX_ROWS}",
    Rule(type="cellIs", operator='between', formula=['0.01', '3.99'], dxf=DifferentialStyle(fill=ORANGE_FILL, font=ORANGE_FONT))
)

# 4. Margin/ROI highlight (Orange if < 30%)
for col in [r_margin, r_roi]:
    res_ws.conditional_formatting.add(
        f"{col}{start_row_res}:{col}{start_row_res + MAX_ROWS}",
        Rule(type="cellIs", operator='between', formula=['0.00001', '0.29999'], dxf=DifferentialStyle(fill=ORANGE_FILL, font=ORANGE_FONT))
    )

# 5. Summary Formulas (E2, F2, G2, O2, P2)
# Profit Total
buy_ws["E2"] = f"=SUM({b_profit}{start_row}:{b_profit}{start_row + MAX_ROWS})"
# Overall ROI
buy_ws["F2"] = f"=IFERROR(E2/SUM({b_cost}{start_row}:{b_cost}{start_row + MAX_ROWS}), 0)"
# Total Sales (Estimated)
buy_ws["G2"] = f"=SUM({b_opp}{start_row}:{b_opp}{start_row + MAX_ROWS})"
# Total Units
buy_ws["O2"] = f"=SUM({b_units}{start_row}:{b_units}{start_row + MAX_ROWS})"
# Total Amount
buy_ws["P2"] = f"=SUM({b_amount}{start_row}:{b_amount}{start_row + MAX_ROWS})"

# --- IP Qty TAB FORMULA INJECTION ---
ip_ws = wb["IP Qty"]
start_row_ip = TABS_CONFIG["IP Qty"]["header_row"] + 1
ip_barcode = get_col("IP Qty", "Barcode")
ip_in_buy = get_col("IP Qty", "In Buy Sheet?")

for row in range(start_row_ip, start_row_ip + MAX_ROWS):
    # Check if Barcode is in BUY tab ASIN column
    # Formula: IF(ISERROR(MATCH(Barcode, BUY!ASIN_Col, 0)), "No", "YES")
    ip_ws[f"{ip_in_buy}{row}"] = f'=IF(ISERROR(MATCH({ip_barcode}{row}, BUY!$A:$A, 0)), "No", "YES")'

# --- ORDER TAB LOGIC (Pivot-style Summary) ---
order_ws = wb["ORDER"]

# We use Dynamic Array formulas to create a "Pivot Table" effect.
# This requires Excel 365 or 2021+. If they have older, they can still use it as a static list.

# 1. Unique Item Codes where Qty > 0
# Formula: =UNIQUE(FILTER(BuyData[Item Code], (BuyData[Order Qty (Unit)]>0)*(BuyData[Item Code]<>"")))
order_ws["A2"] = '=IFERROR(UNIQUE(FILTER(BuyData[Item Code], (BuyData[Order Qty (Unit)]>0)*(BuyData[Item Code]<>""))), "No Orders")'

# 2. Description (using SPILL reference A2#)
# Formula: =XLOOKUP(A2#, BuyData[Item Code], BuyData[Description], "")
order_ws["B2"] = '=IF(A2="No Orders", "", XLOOKUP(A2#, BuyData[Item Code], BuyData[Description], ""))'

# 3. Unit Cost (Average or first match)
# Formula: =XLOOKUP(A2#, BuyData[Item Code], BuyData[Cost], 0)
order_ws["C2"] = '=IF(A2="No Orders", "", XLOOKUP(A2#, BuyData[Item Code], BuyData[Cost], 0))'

# 4. Total Order Qty (SUMIFS)
# Formula: =SUMIFS(BuyData[Order Qty (Unit)], BuyData[Item Code], A2#)
order_ws["D2"] = '=IF(A2="No Orders", "", SUMIFS(BuyData[Order Qty (Unit)], BuyData[Item Code], A2#))'

# 5. Total Order Amount (SUMIFS)
# Formula: =SUMIFS(BuyData[Order Amount], BuyData[Item Code], A2#)
order_ws["E2"] = '=IF(A2="No Orders", "", SUMIFS(BuyData[Order Amount], BuyData[Item Code], A2#))'

# 7. Wrap BUY data in an Excel Table for structured references
buy_last_col = get_column_letter(len(TABS_CONFIG["BUY"]["headers"]))
buy_table_range = f"A{TABS_CONFIG['BUY']['header_row']}:{buy_last_col}{3 + MAX_ROWS}"
buy_table = Table(displayName="BuyData", ref=buy_table_range)

# Add a default style
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
buy_table.tableStyleInfo = style
buy_ws.add_table(buy_table)

    # --- SAVE WORKBOOK ---
try:
    wb.save(OUTPUT_FILE)
    print(f"✅ BUY recommendation workbook generated with legacy improvements: {OUTPUT_FILE}")
except PermissionError:
    print(f"❌ ERROR: Could not save to {OUTPUT_FILE}.")
    print("👉 Please CLOSE the Excel file if it's currently open and try again.")
except Exception as e:
    print(f"❌ AN UNEXPECTED ERROR OCCURRED: {e}")
