import pandas as pd
from difflib import get_close_matches
import os
from tkinter import Tk, filedialog, simpledialog, Label, messagebox
import webbrowser

tb_file_path = ''
pwp_file_path = ''
save_dir = ''
selected_promo = ''
promo_names = []


def T_tb_file_path(filename):
    global tb_file_path
    tb_file_path = filename
    
def T_pwp_file_path(filename):
    global pwp_file_path
    pwp_file_path = filename

def T_save_dir(filename):
    global save_dir
    save_dir = filename


def find_closest_column(df, target_col):
    col_names = [str(col) for col in df.columns]
    matches = get_close_matches(target_col, col_names, n=1, cutoff=0.6)
    return matches[0] if matches else None



def open_file_directory(file_path):
    webbrowser.open(f'file:///{os.path.dirname(file_path)}')



def T_get_promo():
    global pwp_file_path
    global promo_names
    
    # Read Excel file into DataFrame
    pwp_df = pd.read_excel(pwp_file_path, sheet_name="TikTok | Campaign List", header=None)
    
    # Drop a specific row (example: dropping row index 5)
    pwp_df.drop(pwp_df.index[:5], inplace=True)
    pwp_df.columns = pwp_df.iloc[0]
    pwp_df = pwp_df[1:].reset_index(drop=True)

    promo_name_col = "Promo Name (Scheme)"
    
    # Find the closest column name or default to "Promo Name (Scheme)"
    promo_name_col = find_closest_column(pwp_df, "Promo Name (Scheme)") or "Promo Name (Scheme)"

    promo_names = pwp_df[promo_name_col].dropna().unique()

    print(promo_names)
    


def T_promo(choice):
    global selected_promo
    selected_promo = choice



def main():


    root = Tk()
    root.withdraw()


    updated_tb_file_path = os.path.join(save_dir, "Updated_" + os.path.basename(tb_file_path))

    if os.path.exists(updated_tb_file_path):
        os.chmod(updated_tb_file_path, 0o777)

    # Read TB file with header in the second row, assuming the file has headers but no data
    tb_df = pd.read_excel(tb_file_path, sheet_name=0, header=1)
    print("TB DataFrame after loading with header in the second row:")
    print(tb_df.head())

    # Verify if TB DataFrame is empty
    if tb_df.empty:
        print("TB DataFrame is empty after loading. Continuing to populate with PWP data.")

    # Read PWP file
    pwp_df = pd.read_excel(pwp_file_path, sheet_name="TikTok | Campaign List", header=None)
    pwp_df.drop(pwp_df.index[:5], inplace=True)
    pwp_df.columns = pwp_df.iloc[0]
    pwp_df = pwp_df[1:].reset_index(drop=True)
    print("PWP DataFrame after setting headers:")
    print(pwp_df.head())

    # Column names identification
    tb_id_col = 'SKU ID'
    pwp_id_col = 'SKU ID'
    tb_product_id_col = 'Product ID'
    pwp_product_id_col = 'Product Id'
    campaign_price_col = 'Recommended Campaign Price'
    promo_name_col = 'Promo Name (Scheme)'
    discounted_price_col = 'Discounted Price/ASP (VATIN)'
    sales_price_col = 'Campaign Price'

    tb_id_col = find_closest_column(tb_df, tb_id_col) or tb_id_col
    pwp_id_col = find_closest_column(pwp_df, pwp_id_col) or pwp_id_col
    tb_product_id_col = find_closest_column(tb_df, tb_product_id_col) or tb_product_id_col
    pwp_product_id_col = find_closest_column(pwp_df, pwp_product_id_col) or pwp_product_id_col
    campaign_price_col = find_closest_column(tb_df, campaign_price_col) or campaign_price_col
    promo_name_col = find_closest_column(pwp_df, promo_name_col) or promo_name_col
    discounted_price_col = find_closest_column(pwp_df, discounted_price_col) or discounted_price_col
    sales_price_col = find_closest_column(tb_df, sales_price_col) or sales_price_col

    print(f"Identified Columns:\nSKU ID: {tb_id_col}\nPWP SKU ID: {pwp_id_col}\nProduct ID: {tb_product_id_col}\nPWP Product ID: {pwp_product_id_col}\nCampaign Price: {campaign_price_col}\nPromo Name: {promo_name_col}\nDiscounted Price: {discounted_price_col}\nSales Price: {sales_price_col}")

    if promo_name_col not in pwp_df.columns:
        print(f"Error: '{promo_name_col}' column not found in PWP DataFrame.")
        return


    if not selected_promo:
        print("No promo selected. Exiting.")
        return

    filtered_pwp_df = pwp_df[pwp_df[promo_name_col] == selected_promo].reset_index(drop=True)
    print("Filtered PWP DataFrame:")
    print(filtered_pwp_df.head())

    # Normalize IDs to ensure correct matching
    filtered_pwp_df[pwp_id_col] = filtered_pwp_df[pwp_id_col].astype(str).str.strip().str.upper()

    # Ensure numeric comparison for prices
    filtered_pwp_df[discounted_price_col] = pd.to_numeric(filtered_pwp_df[discounted_price_col], errors='coerce')

    print("Normalized and Converted Data:")
    print(filtered_pwp_df[[pwp_id_col, discounted_price_col, pwp_product_id_col]].head())

    # Since TB is empty, we'll create a new DataFrame for good_for_upload_df
    good_for_upload_df = pd.DataFrame(columns=tb_df.columns)

    tb_ids_in_pwp = set(filtered_pwp_df[pwp_id_col].dropna().unique())
    print(f"IDs in PWP: {tb_ids_in_pwp}")

    for idx, row in filtered_pwp_df.iterrows():
        product_id = row[pwp_id_col]
        pwp_product_id = row[pwp_product_id_col]
        discounted_price = row[discounted_price_col]
        new_row = pd.DataFrame({
            tb_product_id_col: [pwp_product_id],
            tb_id_col: [product_id],
            campaign_price_col: [discounted_price],
            sales_price_col: [discounted_price]
        })
        good_for_upload_df = pd.concat([good_for_upload_df, new_row], ignore_index=True)

    # Ensure the IDs are saved as text to prevent scientific notation
    good_for_upload_df[tb_product_id_col] = good_for_upload_df[tb_product_id_col].astype(str)
    good_for_upload_df[tb_id_col] = good_for_upload_df[tb_id_col].astype(str)

    print("Good for upload DataFrame:")
    print(good_for_upload_df)

    try:
        with pd.ExcelWriter(updated_tb_file_path, engine='openpyxl') as writer:
            good_for_upload_df.to_excel(writer, sheet_name="Good for upload", index=False)
            # Adjust column widths for readability
            for sheetname in writer.sheets:
                worksheet = writer.sheets[sheetname]
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column].width = adjusted_width
        print(f"Processing complete. Updated file saved to {updated_tb_file_path}.")
    except PermissionError as e:
        print(f"PermissionError: {e}. Ensure the file is not open or read-only and try again.")

    messagebox.showinfo("Process Complete", f"Processing complete. Updated file saved to \n{updated_tb_file_path}.")
    open_file_directory(updated_tb_file_path)

if __name__ == "__main__":
    main()
