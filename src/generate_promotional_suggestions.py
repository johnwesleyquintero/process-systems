import pandas as pd
from datetime import datetime, timedelta
import argparse
import os

def generate_promotional_suggestions(
    brand_name,
    discount_percentage=0.15,  # 15% discount
    age_threshold_months=6,    # Products older than 6 months
    promotion_duration_days=7  # Promotion lasts for 7 days
):
    input_file_path = os.path.join('BRANDS', brand_name, 'all-listing-report.tsv')
    output_file_path = os.path.join('excel_templates', 'promotional_discount_suggestions.csv')
    """
    Generates a CSV file with suggested promotional discounts for Amazon listings.

    Args:
        input_file_path (str): Path to the all-listing-report.tsv file.
        output_file_path (str): Path where the promotional suggestions CSV will be saved.
        discount_percentage (float): The percentage discount to apply (e.g., 0.15 for 15%).
        age_threshold_months (int): Products older than this many months will be considered for promotion.
        promotion_duration_days (int): The number of days the promotion will last.
    """
    try:
        # Read the TSV file
        df = pd.read_csv(input_file_path, sep='\t')

        # Convert 'open-date' to datetime objects, specifying the format
        df['open-date'] = pd.to_datetime(df['open-date'], format='%Y-%m-%dT%H:%M:%SZ', errors='coerce')

        # Calculate the age threshold date
        current_date = datetime.now()
        age_threshold_date = current_date - timedelta(days=age_threshold_months * 30) # Approx months

        # Filter for products older than the age threshold and with an active status
        eligible_products = df[
            (df['open-date'] < age_threshold_date) &
            (df['status'] == 'Active') &
            (pd.to_numeric(df['price'], errors='coerce').notna())
        ].copy()

        if eligible_products.empty:
            print("No eligible products found for promotion based on the criteria.")
            # Create an empty DataFrame with the expected columns if no eligible products
            output_df = pd.DataFrame(columns=['seller-sku', 'standard-price', 'sale-price', 'sale-start-date', 'sale-end-date'])
        else:
            # Convert price to numeric, handling potential non-numeric values
            eligible_products['price'] = pd.to_numeric(eligible_products['price'], errors='coerce')

            # Calculate the sale price
            eligible_products['sale-price'] = eligible_products['price'] * (1 - discount_percentage)

            # Set promotion dates
            sale_start_date = current_date.strftime('%Y-%m-%d')
            sale_end_date = (current_date + timedelta(days=promotion_duration_days)).strftime('%Y-%m-%d')

            # Create the output DataFrame
            output_df = eligible_products[[
                'seller-sku',
                'price'
            ]].rename(columns={'price': 'standard-price'})
            output_df['sale-price'] = eligible_products['sale-price'].round(2) # Round to 2 decimal places
            output_df['sale-start-date'] = sale_start_date
            output_df['sale-end-date'] = sale_end_date

        # Save to CSV
        output_df.to_csv(output_file_path, index=False)
        print(f"Promotional suggestions saved to {output_file_path}")

    except FileNotFoundError:
        print(f"Error: Input file not found at {input_file_path}")
    except KeyError as e:
        print(f"Error: Missing expected column in TSV file: {e}. Please ensure 'open-date', 'status', 'price', and 'seller-sku' columns exist.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate promotional discount suggestions for a specific brand.")
    parser.add_argument('--brand', type=str, default='SL', help="The brand name (e.g., 'SL', 'STK').")
    parser.add_argument('--discount', type=float, default=0.15, help="The percentage discount to apply (e.g., 0.15 for 15%).")
    parser.add_argument('--age', type=int, default=6, help="Products older than this many months will be considered for promotion.")
    parser.add_argument('--duration', type=int, default=7, help="The number of days the promotion will last.")
    args = parser.parse_args()

    generate_promotional_suggestions(
        brand_name=args.brand,
        discount_percentage=args.discount,
        age_threshold_months=args.age,
        promotion_duration_days=args.duration
    )
