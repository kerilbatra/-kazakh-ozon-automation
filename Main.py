import requests
import json
import pandas as pd
from datetime import datetime, timedelta

url = "https://api-seller.ozon.ru/v2/finance/realization"

headers = {
    "Client-Id": #Enter Your Client ID Here,
    "Api-Key": #Enter Your API Key Here,
    "Content-Type": "application/json"
}

# Get the current date
today = datetime.today()

# Get the first day of the current month and subtract one day to get the last day of the previous month
first_day_of_current_month = today.replace(day=1)
last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)


# Extract previous month and year
data = {
    "month": last_day_of_previous_month.month,
    "year": last_day_of_previous_month.year
}


try:
    response = requests.post(url, headers=headers, json=data)
    
    if response.status_code == 200:
        print("Response JSON:")
        json_data = response.json()
        
        # Flatten the header (top-level metadata)
        header = json_data['result']['header']
        header_df = pd.json_normalize(header)
        
        # Flatten the rows (the main financial data)
        rows = json_data['result']['rows']
        
        # Handle the rows one by one to extract 'item' and 'delivery_commission'
        rows_flat = []
        
        for row in rows:
            # Flatten the 'item' dictionary into columns
            item = row['item']
            row_flat = {
                'rowNumber': row['rowNumber'],
                'seller_price_per_instance': row['seller_price_per_instance'],
                'commission_ratio': row['commission_ratio'],
                'item_name': item['name'],
                'item_offer_id': item['offer_id'],
                'item_barcode': item['barcode'],
                'item_sku': item['sku'],
            }
            
            # If delivery_commission exists, flatten it
            if row.get('delivery_commission'):
                delivery_commission = row['delivery_commission']
                row_flat.update({
                    'delivery_price_per_instance': delivery_commission['price_per_instance'],
                    'delivery_quantity': delivery_commission['quantity'],
                    'delivery_amount': delivery_commission['amount'],
                    'delivery_compensation': delivery_commission['compensation'],
                    'delivery_commission': delivery_commission['commission'],
                    'delivery_bonus': delivery_commission['bonus'],
                    'delivery_standard_fee': delivery_commission['standard_fee'],
                    'delivery_total': delivery_commission['total'],
                    'delivery_stars': delivery_commission['stars'],
                    'delivery_bank_coinvestment': delivery_commission['bank_coinvestment'],
                    'delivery_pick_up_point_coinvestment': delivery_commission['pick_up_point_coinvestment']
                })
            
            rows_flat.append(row_flat)

        # Convert the flattened rows to a DataFrame
        rows_df = pd.DataFrame(rows_flat)

        # Merge the header data into each row, as it's repeated across all rows
        final_df = pd.concat([header_df] * len(rows_df), ignore_index=True)
        final_df = pd.concat([final_df, rows_df], axis=1)

        # Save the DataFrame to an Excel file
        #final_df.to_excel(Enter Your Local Excel File Path Here, index=False)
        print("Data saved")
    else:
        print(f"Failed to retrieve report. Status code: {response.status_code}")
        print("Response:", response.text)
        
except requests.exceptions.RequestException as e:
    print("An error occurred:", e)



import os
import pandas as pd
from openpyxl import load_workbook
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.authentication_context import AuthenticationContext

# SharePoint link and file details
site_url = #Enter Your Site URL Here
doc_library = #Enter Your Document Library Name Here (e.g., "Shared Documents")
file_name = #Enter Your File Name Here (e.g., "Ozon.xlsx")

# Credentials
username = #Enter Your Username Here
password = #Enter Your Password Here

# Define the path to your local Excel file containing the new data
new_data_file_path = #enter Your Local Excel File Path Here

# Define the columns to check for duplicates
subset_columns = [
    'number', 'doc_date', 'start_date', 'stop_date', 'contract_date', 'contract_number', 'payer_name', 'payer_inn',
    'payer_kpp', 'receiver_name', 'receiver_inn', 'receiver_kpp', 'doc_amount', 'vat_amount', 'currency_sys_name',
    'rowNumber', 'seller_price_per_instance', 'commission_ratio', 'item_name', 'item_offer_id', 'item_barcode',
    'item_sku', 'delivery_price_per_instance', 'delivery_quantity', 'delivery_amount', 'delivery_compensation',
    'delivery_commission', 'delivery_bonus', 'delivery_standard_fee', 'delivery_total', 'delivery_stars',
    'delivery_bank_coinvestment', 'delivery_pick_up_point_coinvestment'  # Fixed closing quote
]


try:
    # Authenticate to SharePoint
    auth_ctx = AuthenticationContext(site_url)
    if not auth_ctx.acquire_token_for_user(username, password):
        print(f"Error acquiring token: {auth_ctx.get_last_error()}")
        exit(1)

    ctx = ClientContext(site_url, auth_ctx)

    # Download the existing file from SharePoint
    response = File.open_binary(ctx, f"{doc_library}/{file_name}")

    # Write the response content to a local file
    with open(file_name, "wb") as existing_file:
        existing_file.write(response.content)

    print("Existing file has been downloaded from SharePoint.")

    # Load the existing data into a DataFrame
    try:
        existing_df = pd.read_excel(file_name)
        print("Existing data loaded into DataFrame.")
    except Exception as e:
        print(f"Error loading existing data: {e}")
        exit(1)

except Exception as e:
    print(f"Error during file download for existing data: {e}")
    exit(1)

# Load the new data into a DataFrame
try:
    df = pd.read_excel(new_data_file_path)
    print("New data loaded into DataFrame.")
except Exception as e:
    print(f"Error loading new data from the file: {e}")
    exit(1)

# Combine the existing data with the new data
try:
    # Combine and deduplicate based on the subset columns
    combined_df = pd.concat([existing_df, df], ignore_index=True)
    combined_df = combined_df.drop_duplicates(subset=subset_columns)
    print("New data merged with existing data successfully.")
except Exception as e:
    print(f"Error while merging data: {e}")
    exit(1)

# Get the new rows added after combining
new_rows = combined_df[len(existing_df):]

if new_rows.empty:
    print("No new rows to add.")
else:
    # Append the new rows to the Excel sheet
    try:
        book = load_workbook(file_name)
        sheet = book['Sheet1']  # Ensure this matches your sheet name
        for row in new_rows.itertuples(index=False, name=None):
            sheet.append(row)

        # Save the modified file
        book.save(file_name)
        print("Data has been successfully written to the Excel file without duplicates.")
    except Exception as e:
        print(f"Error saving the Excel file: {e}")
        exit(1)

    # Upload the updated file back to SharePoint
    try:
        with open(file_name, 'rb') as content_file:
            file_content = content_file.read()

        File.save_binary(ctx, f"{doc_library}/{file_name}", file_content)
        print("File has been uploaded back to SharePoint.")
    except Exception as e:
        print(f"Error during file upload: {e}")

# Cleanup: Remove the local file
if os.path.exists(file_name):
    os.remove(file_name)
    print("Local file has been deleted.")