import csv
from datetime import datetime, timedelta
from collections import defaultdict
import argparse
import os

def parse_sales_data(sales_file_path):
    sales_velocity = defaultdict(lambda: {"total_quantity": 0, "days_sold": set()})
    try:
        with open(sales_file_path, 'r', encoding='utf-8') as f:
            # The sales file is tab-separated, so we specify the delimiter.
            reader = csv.DictReader(f, delimiter='\t')
            for row in reader:
                # Only consider orders that have been shipped.
                if row.get('order-status') != 'Shipped':
                    continue

                sku = row.get('sku')
                try:
                    quantity = int(row.get('quantity', 0))
                except (ValueError, TypeError):
                    continue # Skip if quantity is not a valid integer

                purchase_date_str = row.get('purchase-date')

                if sku and quantity > 0 and purchase_date_str:
                    try:
                        # Handle ISO 8601 format like '2025-07-10T20:58:30+00:00'
                        sale_date = datetime.fromisoformat(purchase_date_str.replace('+00:00', ''))
                        sales_velocity[sku]["total_quantity"] += quantity
                        sales_velocity[sku]["days_sold"].add(sale_date.date())
                    except (ValueError, TypeError):
                        continue
    except FileNotFoundError:
        print(f"Error: Sales file not found at {sales_file_path}")
        return None
    except Exception as e:
        print(f"Error processing sales file: {e}")
        return None
    return sales_velocity

def parse_inventory_data(inventory_file_path):
    inventory_levels = {}
    try:
        with open(inventory_file_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f) # Assuming comma-separated for inventory
            # print(f"Inventory file headers: {reader.fieldnames}") # Debug print headers
            for row in reader:
                sku = row.get('sku') # Assuming 'sku' is the column name for SKU
                # Assuming 'quantity available' or similar is the column for available inventory
                available_quantity = 0
                try:
                    available_quantity = int(row.get('available', 0))
                except ValueError:
                    continue
                
                if sku and available_quantity >= 0:
                    inventory_levels[sku] = available_quantity
    except FileNotFoundError:
        print(f"Error: Inventory file not found at {inventory_file_path}")
        return None
    except Exception as e:
        print(f"Error processing inventory file: {e}")
        return None
    return inventory_levels

def generate_restock_recommendations(sales_data, inventory_data, lead_time_days=14, safety_stock_days=7, desired_days_of_cover=30):
    recommendations = []
    if not sales_data or not inventory_data:
        return recommendations

    for sku, sales_info in sales_data.items():
        total_quantity_sold = sales_info["total_quantity"]
        days_with_sales = len(sales_info["days_sold"])

        if days_with_sales == 0:
            continue

        avg_daily_sales = total_quantity_sold / days_with_sales
        if avg_daily_sales <= 0:
            continue

        current_inventory = inventory_data.get(sku, 0)
        
        # Core calculations based on the strategic guide
        safety_stock = safety_stock_days * avg_daily_sales
        reorder_point = (lead_time_days * avg_daily_sales) + safety_stock
        days_of_supply = current_inventory / avg_daily_sales if avg_daily_sales > 0 else float('inf')

        # Trigger recommendation if inventory is below the reorder point
        if current_inventory < reorder_point:
            # Calculate how many units to order
            # Order Quantity = (Desired Days of Cover × Daily Sales Velocity) - (Current Inventory + Inbound Inventory)
            # Assuming inbound inventory is 0 for now
            inbound_inventory = 0 
            desired_inventory_level = desired_days_of_cover * avg_daily_sales
            order_quantity = desired_inventory_level - (current_inventory + inbound_inventory)
            
            recommendations.append({
                "sku": sku,
                "avg_daily_sales": round(avg_daily_sales, 2),
                "current_inventory": current_inventory,
                "days_of_supply": round(days_of_supply, 2),
                "reorder_point": round(reorder_point, 2),
                "recommended_order_quantity": max(0, int(order_quantity)),
                "recommendation": f"Stock below reorder point ({int(reorder_point)} units). Recommend ordering."
            })
    
    # Sort recommendations by the urgency (days of supply)
    recommendations.sort(key=lambda x: x["days_of_supply"])
    
    return recommendations

def save_recommendations(recommendations, output_file_path):
    if not recommendations:
        print("No restock recommendations generated.")
        return

    try:
        with open(output_file_path, 'w', newline='', encoding='utf-8') as f:
            fieldnames = [
                "sku", "avg_daily_sales", "current_inventory", "days_of_supply", 
                "reorder_point", "recommended_order_quantity", "recommendation"
            ]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(recommendations)
        print(f"Restock recommendations saved to {output_file_path}")
    except Exception as e:
        print(f"Error saving recommendations: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate FBA restock recommendations for a specific brand.")
    parser.add_argument('--brand', type=str, default='SL', help="The brand name (e.g., 'SL', 'STK').")
    parser.add_argument('--lead_time', type=int, default=21, help="Average time from ordering to inventory being available at FBA in days.")
    parser.add_argument('--safety_stock', type=int, default=10, help="Extra stock to buffer against delays or sales spikes in days of cover.")
    parser.add_argument('--desired_cover', type=int, default=45, help="How many days of sales the new order should cover.")
    args = parser.parse_args()

    # --- Configuration ---
    LEAD_TIME_DAYS = args.lead_time
    SAFETY_STOCK_DAYS = args.safety_stock
    DESIRED_DAYS_OF_COVER = args.desired_cover

    # --- File Paths ---
    sales_file_path = os.path.join('BRANDS', args.brand, 'reports', 'sales', 'sales.csv')
    inventory_file_path = os.path.join('BRANDS', args.brand, 'reports', 'inventory', 'inventory.csv')
    recommendations_output_path = os.path.join('excel_templates', 'restock_recommendations.csv')

    print(f"Analyzing sales data from: {sales_file_path}")
    sales_data = parse_sales_data(sales_file_path)
    
    print(f"Analyzing inventory data from: {inventory_file_path}")
    inventory_data = parse_inventory_data(inventory_file_path)

    if sales_data and inventory_data:
        print("Generating restock recommendations with the following settings:")
        print(f"- Lead Time: {LEAD_TIME_DAYS} days")
        print(f"- Safety Stock: {SAFETY_STOCK_DAYS} days of cover")
        print(f"- Desired Days of Cover: {DESIRED_DAYS_OF_COVER} days")
        
        restock_recommendations = generate_restock_recommendations(
            sales_data, 
            inventory_data,
            lead_time_days=LEAD_TIME_DAYS,
            safety_stock_days=SAFETY_STOCK_DAYS,
            desired_days_of_cover=DESIRED_DAYS_OF_COVER
        )
        save_recommendations(restock_recommendations, recommendations_output_path)
    else:
        print("Could not generate recommendations due to missing data.")
