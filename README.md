# Amazon Flatfile & Systems Assistant

### Project Overview

The "Amazon Systems Assistant" is a localized, user-friendly micro-application designed to help our team automate repetitive tasks and generate intelligent reports for Amazon operations. It uses a simple architecture combining Excel for user interaction and Python for data processing and systems generation.

The goal of this system is to improve our team's efficiency and accuracy, reduce the time spent on manual work, and elevate our focus from simple data entry to high-level strategy and analysis.

### Architecture

*   **Input Layer (Excel/CSV):** Simple, non-threatening Excel or CSV files serve as the user's main interaction point.
*   **Processing Engine (Python):** A powerful suite of Python scripts that function as the "brain" of the operation. They read inputs, perform complex calculations, blend data, and forge new files.
*   **Launchpad (NPM Scripts):** The `package.json` file contains simple command-line scripts that act as easy-to-use "buttons," allowing any team member to trigger the Python automation without needing to write code.

### Workflow Blueprint

1.  **Select a Task:** Identify the process you need to run (e.g., generate a restock report, create a pricing template).
2.  **Prepare Input (if necessary):** For some scripts, you may need to fill out a simple Excel/CSV template.
3.  **Run Command:** Open your terminal in the project root and execute the corresponding NPM script (e.g., `npm run forge:restock-recommender`).
4.  **Retrieve Output:** The generated Excel template or Amazon-ready flat file will be saved in the `excel_templates/` or `output/` directory, ready for use.

---

## The Pythonic Toolkit: Current Capabilities

This repository contains two primary types of systems: **Flatfile Generators** for direct Amazon uploads, and **Advanced Template Forges** for creating data-driven analysis tools.

### Flatfile Generation Suite

This suite of scripts generates Amazon-ready flat files for direct upload, streamlining common operational tasks.

*   **Price Update Automation**
    *   **Purpose:** Automates the generation of Amazon price update flat files, ensuring quick and accurate price adjustments across your listings.
    *   **Script:** `src/price_update.py`
    *   **Command:** `npm run run:price-update -- --brand [BRAND_NAME]`
    *   **Output:** `BRANDS/[BRAND_NAME]/output/amazon_price_update_flatfile.csv`
    *   **Template Usage:** Prepare your price updates in `excel_templates/price_update_template.csv`.

*   **FBA Restock Recommendations**
    *   **Purpose:** Generates precise FBA restock recommendations by analyzing your inventory and sales data, helping you maintain optimal stock levels and avoid stockouts.
    *   **Script:** `src/restock_recommender.py`
    *   **Command:** `npm run recommend:restock -- --brand [BRAND_NAME]`
    *   **Output:** `BRANDS/[BRAND_NAME]/recommendations/restock_recommendations.csv`
    *   **Template Usage:** Ensure your FBA Inventory and Business Report data are updated in `BRANDS/[BRAND_NAME]/reports/inventory/inventory.csv` and `BRANDS/[BRAND_NAME]/reports/sales/sales.csv` respectively.

*   **Promotional Discount Suggestions**
    *   **Purpose:** Provides data-driven suggestions for promotional discounts, helping you strategize effective sales campaigns to boost product visibility and sales.
    *   **Script:** `src/generate_promotional_suggestions.py`
    *   **Command:** `npm run generate:promotions -- --brand [BRAND_NAME]`
    *   **Output:** `BRANDS/[BRAND_NAME]/output/promotional_discount_suggestions.csv`
    *   **Template Usage:** Requires up-to-date sales data for analysis.

*   **Listing Creation Automation**
    *   **Purpose:** Automates the creation of new Amazon listings, simplifying the process of adding new products to your catalog and ensuring all required fields are correctly populated.
    *   **Script:** `src/listing_creation.py`
    *   **Command:** `npm run run:listing-creation -- --brand [BRAND_NAME]`
    *   **Output:** `BRANDS/[BRAND_NAME]/output/amazon_new_listing_flatfile.csv`
    *   **Template Usage:** Fill out the `excel_templates/new_listing_template.csv` with your new listing details.

*   **Template Validation**
    *   **Purpose:** Validates input templates (e.g., price update, new listing) against Amazon's requirements, preventing errors before flat file generation and upload.
    *   **Script:** `src/template_validator.py`
    *   **Command:** `npm run validate:template -- --template [TEMPLATE_NAME] --brand [BRAND_NAME]`
    *   **Output:** Validation results are displayed in the console, indicating any discrepancies or errors.
    *   **Template Usage:** Specify the template to validate (e.g., `price_update` or `new_listing`).

---

### Strategic Analysis Templates

This suite of scripts generates advanced, data-driven Excel templates designed to automate complex analysis and guide strategic decision-making. These templates provide structured frameworks for inventory management, profitability assessment, and competitive tracking.

*   **Buy Box Dominance Tracker**
    *   **Purpose:** Generates the `Buy_Box_Dominance_Tracker.xlsx`, an analytical dashboard that processes sales and traffic data to identify ASINs losing the Buy Box and suggests corrective actions.
    *   **Script:** `src/forge_trojan_horse.py`
    *   **Command:** `npm run forge:buybox-tracker`
    *   **Output:** `excel_templates/Buy_Box_Dominance_Tracker_v1.1.xlsx`
    *   **Template Usage:** Paste the 'Detail Page Sales & Traffic' report into the `Data_Input` tab. The `Buy_Box_Dashboard` will auto-populate, highlighting critical issues and suggesting priority levels.

*   **Restock Recommender**
    *   **Purpose:** Generates the `Restock_Recommender.xlsx`, a comprehensive supply chain management tool. It integrates data from FBA Inventory and Business Reports to calculate inventory requirements and provide a precise 60-day restock forecast.
    *   **Script:** `src/forge_restock_report.py`
    *   **Command:** `npm run forge:restock-recommender`
    *   **Output:** `excel_templates/Restock_Recommender_v1.2.xlsx`
    *   **Template Usage:** Paste the FBA Inventory and Business Report data into their respective `Data_Input` tabs. The `Restock_Dashboard` will provide a complete restocking plan.

*   **Listing Health & Sentiment Report Template**
    *   **Purpose:** Generates the `Listing_Health_Sentiment_Report.xlsx`, a structured framework for in-depth listing analysis. It consolidates qualitative and quantitative data for comprehensive investigation.
    *   **Script:** `src/forge_sentiment_report.py`
    *   **Command:** `npm run forge:sentiment-report`
    *   **Output:** `excel_templates/Listing_Health_Sentiment_Report_v1.0.xlsx`
    *   **Template Usage:** Paste data from VOC, Business Reports, and competitor analysis into the input tabs. Use the `ASIN_Deep_Dive` tab to consolidate findings and formulate recommendations.

*   **Surgical Strike Profitability Calculator**
    *   **Purpose:** Generates the `Surgical_Strike_Calculator.xlsx`, a precision tool designed to assess the profitability of proposed promotional discounts.
    *   **Script:** `src/forge_calculator.py`
    *   **Command:** `npm run forge:promo-calculator`
    *   **Output:** `excel_templates/Surgical_Strike_Calculator_v1.0.xlsx`
    *   **Template Usage:** Input the product's price, costs, fees, and a proposed discount percentage. The calculator will instantly display the final profit per unit and highlight potential losses.

---
