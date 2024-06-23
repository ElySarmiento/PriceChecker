import pandas as pd
import numpy as np
from difflib import get_close_matches
import os
from tkinter import Tk, filedialog, simpledialog, Label, messagebox
import webbrowser


tb_file_path = ''
pwp_file_path = ''
save_dir = ''
selected_promo = ''
promo_names = []


def S_tb_file_path(filename):
    global tb_file_path
    tb_file_path = filename
    
def S_pwp_file_path(filename):
    global pwp_file_path
    pwp_file_path = filename

def S_save_dir(filename):
    global save_dir
    save_dir = filename



def find_closest_column(df, target_col):
    col_names = [str(col) for col in df.columns]
    matches = get_close_matches(target_col, col_names, n=1, cutoff=0.6)
    return matches[0] if matches else None


def open_file_directory(file_path):
    webbrowser.open(f'file:///{os.path.dirname(file_path)}')

def S_get_promo():
    global pwp_file_path
    global promo_names
    
    # Read Excel file into DataFrame
    pwp_df = pd.read_excel(pwp_file_path, sheet_name="Shp | Campaign List", header=None)
    
    # Drop a specific row (example: dropping row index 5)
    pwp_df.drop(pwp_df.index[:5], inplace=True)
    pwp_df.columns = pwp_df.iloc[0]
    pwp_df = pwp_df[1:].reset_index(drop=True)

    promo_name_col = "Promo Name (Scheme)"
    
    # Find the closest column name or default to "Promo Name (Scheme)"
    promo_name_col = find_closest_column(pwp_df, "Promo Name (Scheme)") or "Promo Name (Scheme)"

    promo_names = pwp_df[promo_name_col].dropna().unique()

    print(promo_names)
    


def S_promo(choice):
    global selected_promo
    selected_promo = choice



def main():

    root = Tk()
    root.withdraw()

    updated_tb_file_path = os.path.join(save_dir, "Updated_" + os.path.basename(tb_file_path))

    if os.path.exists(updated_tb_file_path):
        os.chmod(updated_tb_file_path, 0o777)

    tb_df = pd.read_excel(tb_file_path, sheet_name=0)
    pwp_df = pd.read_excel(pwp_file_path, sheet_name="Shp | Campaign List", header=None)
    pwp_df.drop(pwp_df.index[:5], inplace=True)
    pwp_df.columns = pwp_df.iloc[0]
    pwp_df = pwp_df[1:].reset_index(drop=True)

    print("TB DataFrame columns:")
    print(tb_df.columns)
    print(tb_df.head())

    print("PWP DataFrame columns:")
    print(pwp_df.columns)
    print(pwp_df.head())

    tb_id_col = 'Product ID'
    pwp_id_col = 'Product ID'
    campaign_price_col = 'Recommended Campaign Price'
    discounted_price_col = 'Discounted Price/ASP (VATIN)'
    sales_price_col = 'Campaign Price'
    promo_name_col = "Promo Name (Scheme)"

    tb_id_col = find_closest_column(tb_df, tb_id_col) or tb_id_col
    pwp_id_col = find_closest_column(pwp_df, pwp_id_col) or pwp_id_col
    campaign_price_col = find_closest_column(tb_df, campaign_price_col) or campaign_price_col
    promo_name_col = find_closest_column(pwp_df, promo_name_col) or promo_name_col
    discounted_price_col = find_closest_column(pwp_df, discounted_price_col) or discounted_price_col
    sales_price_col = find_closest_column(tb_df, sales_price_col) or sales_price_col

    print(f"Identified Columns:\nProduct ID: {tb_id_col}\nPWP Product ID: {pwp_id_col}\nCampaign Price: {campaign_price_col}\nPromo Name: {promo_name_col}\nDiscounted Price: {discounted_price_col}\nSales Price: {sales_price_col}")



    filtered_pwp_df = pwp_df[pwp_df[promo_name_col] == selected_promo]

    print("Filtered PWP DataFrame columns:")
    print(filtered_pwp_df.columns)

    # Normalize IDs to ensure correct matching
    tb_df[tb_id_col] = tb_df[tb_id_col].astype(str).str.strip().str.upper()
    filtered_pwp_df[pwp_id_col] = filtered_pwp_df[pwp_id_col].astype(str).str.strip().str.upper()

    # Ensure numeric comparison for prices
    tb_df[campaign_price_col] = pd.to_numeric(tb_df[campaign_price_col], errors='coerce')

    filtered_pwp_df[discounted_price_col] = pd.to_numeric(filtered_pwp_df[discounted_price_col], errors='coerce')

    good_for_upload_df = tb_df.copy()
    platform_df = pd.DataFrame(columns=[tb_id_col, campaign_price_col, 'Escalation Reason'])
    brand_df = pd.DataFrame(columns=[tb_id_col, sales_price_col])

    tb_ids_in_pwp = set(filtered_pwp_df[pwp_id_col].dropna().unique())
    print(f"IDs in PWP: {tb_ids_in_pwp}")

    for idx, row in tb_df.iterrows():
        product_id = row[tb_id_col]
        if product_id in tb_ids_in_pwp:
            matching_rows = filtered_pwp_df[filtered_pwp_df[pwp_id_col] == product_id]
            if not matching_rows.empty:
                for _, matching_row in matching_rows.iterrows():
                    discounted_price = matching_row[discounted_price_col]
                    if row[campaign_price_col] >= discounted_price:
                        good_for_upload_df.at[idx, sales_price_col] = discounted_price
                        break
                else:
                    platform_df = pd.concat([platform_df, pd.DataFrame({tb_id_col: [product_id], campaign_price_col: [discounted_price], 'Escalation Reason': ['do not meet the reco price']})])
        else:
            brand_df = pd.concat([brand_df, pd.DataFrame({tb_id_col: [product_id], sales_price_col: [row[sales_price_col]]})])

    good_for_upload_df = good_for_upload_df.dropna(subset=[sales_price_col])

    # Add IDs from PWP not found in TB to Platform sheet and IDs in TB with higher prices
    pwp_ids_not_in_tb = filtered_pwp_df[~filtered_pwp_df[pwp_id_col].isin(tb_df[tb_id_col])]
    platform_df = pd.concat([platform_df, pwp_ids_not_in_tb[[pwp_id_col, discounted_price_col]].rename(columns={pwp_id_col: tb_id_col, discounted_price_col: campaign_price_col}).assign(**{'Escalation Reason': 'not eligible'})])

    print("Brand IDs not found in PWP:")
    print(brand_df[tb_id_col].values)

    print("Good for upload DataFrame:")
    print(good_for_upload_df)

    try:
        with pd.ExcelWriter(updated_tb_file_path, engine='openpyxl') as writer:
            good_for_upload_df.to_excel(writer, sheet_name="Good for upload", index=False)
            platform_df.to_excel(writer, sheet_name="Platform", index=False)
            brand_df.to_excel(writer, sheet_name="Brand", index=False)
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

    messagebox.showinfo("Process Complete", f"Processing complete. Updated file saved to \n {updated_tb_file_path}.")
    open_file_directory(updated_tb_file_path)

if __name__ == "__main__":
    main()
