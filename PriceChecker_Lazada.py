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


def L_tb_file_path(filename):
    global tb_file_path
    tb_file_path = filename
    
def L_pwp_file_path(filename):
    global pwp_file_path
    pwp_file_path = filename

def L_save_dir(filename):
    global save_dir
    save_dir = filename

def find_closest_column(df, target_col):
    col_names = [str(col) for col in df.columns]
    matches = get_close_matches(target_col, col_names, n=1, cutoff=0.6)
    return matches[0] if matches else None

def open_file_directory(file_path):
    webbrowser.open(f'file:///{os.path.dirname(file_path)}')

def L_get_promo():
    global pwp_file_path
    global promo_names
    
    # Read Excel file into DataFrame
    pwp_df = pd.read_excel(pwp_file_path, sheet_name="Lzd | Campaign List", header=None)
    
    # Drop a specific row (example: dropping row index 5)
    pwp_df.drop(pwp_df.index[:5], inplace=True)
    pwp_df.columns = pwp_df.iloc[0]
    pwp_df = pwp_df[1:].reset_index(drop=True)

    promo_name_col = 'Promo Name (Scheme)'
    
    # Find the closest column name or default to "Promo Name (Scheme)"
    promo_name_col = find_closest_column(pwp_df, "Promo Name (Scheme)") or "Promo Name (Scheme)"

    promo_names = pwp_df[promo_name_col].dropna().unique()

    print(promo_names)
    


def L_promo(choice):
    global selected_promo
    selected_promo = choice



def main():
  

    root = Tk()
    root.withdraw()


    updated_tb_file_path = os.path.join(save_dir, "Updated_" + os.path.basename(tb_file_path))

    if os.path.exists(updated_tb_file_path):
        os.chmod(updated_tb_file_path, 0o777)

    tb_df = pd.read_excel(tb_file_path, sheet_name=0)
    pwp_df = pd.read_excel(pwp_file_path, sheet_name="Lzd | Campaign List", header=None)
    pwp_df.columns = pwp_df.iloc[5]
    pwp_df = pwp_df.drop(5)

    print("TB DataFrame columns:")
    print(tb_df.columns)
    print(tb_df.head())

    print("PWP DataFrame columns:")
    print(pwp_df.columns)
    print(pwp_df.head())

    seller_sku_col = 'Seller SKU'
    campaign_price_col = 'Campaign Price'
    promo_name_col = 'Promo Name (Scheme)'
    discounted_price_col = 'Discounted Price/ASP (VATIN)'

    seller_sku_col = find_closest_column(tb_df, seller_sku_col) or seller_sku_col
    campaign_price_col = find_closest_column(tb_df, campaign_price_col) or campaign_price_col
    promo_name_col = find_closest_column(pwp_df, promo_name_col) or promo_name_col
    discounted_price_col = find_closest_column(pwp_df, discounted_price_col) or discounted_price_col

    print(f"Identified Columns:\nSeller SKU: {seller_sku_col}\nCampaign Price: {campaign_price_col}\nPromo Name: {promo_name_col}\nDiscounted Price: {discounted_price_col}")

    if promo_name_col not in pwp_df.columns:
        print(f"Error: '{promo_name_col}' column not found in PWP DataFrame.")
        return


    filtered_pwp_df = pwp_df[pwp_df[promo_name_col] == selected_promo]

    print("Filtered PWP DataFrame columns:")
    print(filtered_pwp_df.columns)

    filtered_seller_sku_col = find_closest_column(filtered_pwp_df, seller_sku_col) or seller_sku_col

    # Normalize SKUs to ensure correct matching
    tb_df[seller_sku_col] = tb_df[seller_sku_col].astype(str).str.strip().str.upper()
    filtered_pwp_df[filtered_seller_sku_col] = filtered_pwp_df[filtered_seller_sku_col].astype(str).str.strip().str.upper()

    good_for_upload_df = tb_df.copy()
    platform_df = pd.DataFrame(columns=[seller_sku_col, campaign_price_col, 'Escalation Reason'])
    brand_df = pd.DataFrame(columns=[seller_sku_col, 'Sales Price'])

    tb_skus_in_pwp = set(filtered_pwp_df[filtered_seller_sku_col].dropna().unique())
    print(f"SKUs in PWP: {tb_skus_in_pwp}")

    for idx, row in tb_df.iterrows():
        sku = row[seller_sku_col]
        if sku in tb_skus_in_pwp:
            matching_row = filtered_pwp_df[filtered_pwp_df[filtered_seller_sku_col] == sku]
            if not matching_row.empty:
                discounted_price = matching_row.iloc[0][discounted_price_col]
                if row['Sales Price'] >= discounted_price:
                    good_for_upload_df.at[idx, campaign_price_col] = discounted_price
                else:
                    platform_df = pd.concat([platform_df, pd.DataFrame({seller_sku_col: [sku], campaign_price_col: [discounted_price], 'Escalation Reason': ['do not meet the reco price']})])
        else:
            brand_df = pd.concat([brand_df, pd.DataFrame({seller_sku_col: [sku], 'Sales Price': [row['Sales Price']]})])

    good_for_upload_df = good_for_upload_df.dropna(subset=[campaign_price_col])

    # Add SKUs from PWP not found in TB to Platform sheet and SKUs in TB with higher prices
    pwp_skus_not_in_tb = filtered_pwp_df[~filtered_pwp_df[filtered_seller_sku_col].isin(tb_df[seller_sku_col])]
    platform_df = pd.concat([platform_df, pwp_skus_not_in_tb[[filtered_seller_sku_col, discounted_price_col]].rename(columns={filtered_seller_sku_col: seller_sku_col, discounted_price_col: campaign_price_col}).assign(**{'Escalation Reason': 'not eligible'})])

    print("Brand SKUs not found in PWP:")
    print(brand_df[seller_sku_col].values)

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

    messagebox.showinfo("Process Complete", f"Processing complete. Updated file saved to \n{updated_tb_file_path}.")
    open_file_directory(updated_tb_file_path)

if __name__ == "__main__":
    main()
