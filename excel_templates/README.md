# Excel Templates Directory

This directory contains the "Trojan Horse" suite of intelligent Excel templates. These tools are programmatically forged to automate complex Amazon analysis and guide strategic decision-making.

## 🛠️ Available Templates

### 1. [BUY_RECOMMENDATIONS.xlsx](./BUY_RECOMMENDATIONS.xlsx)
- **Purpose:** Wholesale Buy Recommendation Engine.
- **Function:** Helps make data-driven BUY/KILL/Prioritize decisions for wholesale inventory.
- **Key Features:** Calculates Profit, ROI, Days of Supply, and Recommended Order Quantity based on BSR trends and sales velocity.
- **Source Script:** `src/ws_buy_rec.py`

### 2. [Buy_Box_Dominance_Tracker_v1.1.xlsx](./Buy_Box_Dominance_Tracker_v1.1.xlsx)
- **Purpose:** Buy Box monitoring and strategy dashboard.
- **Function:** Analyzes 'Detail Page Sales & Traffic' reports to identify ASINs losing the Buy Box.
- **Key Features:** Automatically prioritizes critical ASINs and suggests corrective actions to reclaim dominance.
- **Source Script:** `src/forge_trojan_horse.py`

### 3. [Listing_Health_Sentiment_Report_v1.0.xlsx](./Listing_Health_Sentiment_Report_v1.0.xlsx)
- **Purpose:** Deep-dive ASIN analysis framework.
- **Function:** Consolidates data from Voice of the Customer (VOC), Business Reports, and Competitor Analysis.
- **Key Features:** Provides a structured `ASIN_Deep_Dive` tab for rigorous investigation and recommendation formulation.
- **Source Script:** `src/forge_sentiment_report.py`

### 4. [Restock_Recommender_v1.2.xlsx](./Restock_Recommender_v1.2.xlsx)
- **Purpose:** Supply chain and inventory command center.
- **Function:** Blends FBA Inventory and Business Reports to provide precise restocking forecasts.
- **Key Features:** Includes a "Forecast Control Panel" to adjust for holidays, promotions, and growth trends. Targets a 60-day supply.
- **Source Script:** `src/forge_restock_report.py`

### 5. [Surgical_Strike_Calculator_v1.0.xlsx](./Surgical_Strike_Calculator_v1.0.xlsx)
- **Purpose:** Mathematical "truth machine" for promotion profitability.
- **Function:** Calculates the exact financial impact of proposed discounts.
- **Key Features:** Instantly exposes potential money-losing promotions by factoring in all costs, fees, and discount depths.
- **Source Script:** `src/forge_calculator.py`

## 📜 Legacy Templates

### [Buy Rec 5.0 Legacy.xlsx](./Buy%20Rec%205.0%20Legacy.xlsx)
- **Status:** Legacy Version.
- **Purpose:** Previous iteration of the Wholesale Buy Recommendation Engine.
- **Note:** Retained for historical reference or specific legacy workflows. For current operations, use [BUY_RECOMMENDATIONS.xlsx](./BUY_RECOMMENDATIONS.xlsx).

---
*Note: These templates are generated automatically. To regenerate or update them, run the corresponding NPM scripts as defined in the root [package.json](../package.json).*
