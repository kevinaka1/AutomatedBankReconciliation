import os
import re
from datetime import datetime
import pandas as pd
import requests
from PyPDF2 import PdfReader
from dotenv import load_dotenv

# -------------------- Load .env -------------------- #
load_dotenv()
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
BASE_FOLDER = os.getenv("DRIVE_ROOT")

if not ACCESS_TOKEN:
    raise ValueError("ACCESS_TOKEN not found in .env file")
if not BASE_FOLDER:
    raise ValueError("DRIVE_ROOT not found in .env file")

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# -------------------- PDF Parsing -------------------- #
def extract_transactions(pdf_path):
    reader = PdfReader(pdf_path)
    transactions = []
    beginning_balance = None

    for page in reader.pages:
        text = page.extract_text()
        lines = text.split('\n')

        for line in lines:
            # Extract beginning balance
            if beginning_balance is None:
                match = re.search(r'Beginning Balance[:\s]*\$?([\d,]+\.\d{2})', line)
                if match:
                    beginning_balance = float(match.group(1).replace(',', ''))

            # Extract transaction lines (Date, Description, Amount, Balance)
            txn_match = re.match(
                r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})', line
            )
            if txn_match:
                date, description, amount, balance = txn_match.groups()
                transactions.append({
                    'Date': date,
                    'Description': description,
                    'Amount': float(amount.replace(',', '')),
                    'Balance': float(balance.replace(',', ''))
                })

    if beginning_balance is None:
        raise ValueError("Could not find the beginning balance in the PDF.")
    if not transactions:
        raise ValueError("No transactions found in the PDF.")

    return beginning_balance, transactions

# -------------------- OneDrive Folder Utilities -------------------- #
def ensure_folder_exists(folder_path):
    """
    Ensures that the folder path exists in OneDrive.
    Creates the folder (and any missing parents) if needed.
    """
    url = f"{GRAPH_BASE_URL}/me/drive/root:/{folder_path}"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}

    # Check if folder exists
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return True  # Folder exists

    # Folder doesn't exist, create it
    parent_path = '/'.join(folder_path.strip('/').split('/')[:-1])
    folder_name = folder_path.strip('/').split('/')[-1]
    if parent_path:
        parent_url = f"{GRAPH_BASE_URL}/me/drive/root:/{parent_path}:/children"
    else:
        parent_url = f"{GRAPH_BASE_URL}/me/drive/root/children"

    data = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    create_response = requests.post(parent_url, headers={**headers, "Content-Type": "application/json"}, json=data)
    if create_response.status_code in [201, 409]:  # 409 = already exists
        return True
    else:
        print(f"Failed to create folder {folder_path}: {create_response.status_code} - {create_response.text}")
        return False

# -------------------- Microsoft Graph Upload -------------------- #
def upload_to_onedrive(file_path, statement_date):
    """
    Upload file to OneDrive under BASE_FOLDER/year/month.
    Creates folders if missing.
    """
    file_name = os.path.basename(file_path)

    dt = datetime.strptime(statement_date, "%m/%d/%Y")
    year = dt.year
    month_name = dt.strftime("%B")
    folder_path = f"{BASE_FOLDER}/{year}/{month_name}"

    # Ensure folder exists
    ensure_folder_exists(BASE_FOLDER)
    ensure_folder_exists(f"{BASE_FOLDER}/{year}")
    ensure_folder_exists(folder_path)

    # Upload file
    url = f"{GRAPH_BASE_URL}/me/drive/root:/{folder_path}/{file_name}:/content"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    with open(file_path, 'rb') as f:
        response = requests.put(url, headers=headers, data=f)

    if response.status_code in [200, 201]:
        print(f"Successfully uploaded {file_name} to {folder_path}!")
        return f"https://onedrive.live.com/edit.aspx?resid={response.json().get('id')}"
    else:
        print(f"Upload failed: {response.status_code} - {response.text}")
        return None

# -------------------- Main -------------------- #
def main():
    pdf_path = input("Enter the path to the bank statement PDF: ").strip()
    beginning_balance, transactions = extract_transactions(pdf_path)
    print(f"Beginning Balance: ${beginning_balance:,.2f}")

    # Save Excel file locally
    excel_file = pdf_path.replace('.pdf', '_transactions.xlsx')
    df = pd.DataFrame(transactions)
    df.to_excel(excel_file, index=False)
    print(f"Transactions saved locally to {excel_file}")

    # Use first transaction date to determine year/month folder
    first_date = transactions[0]['Date']
    excel_link = upload_to_onedrive(excel_file, first_date)
    if excel_link:
        print(f"Open your Excel file online here: {excel_link}")

if __name__ == "__main__":
    main()
