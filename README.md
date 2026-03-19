# Bank Reconciliation Script

This Python script automates **bank statement reconciliation**:

1. Parses a PDF bank statement to extract the **beginning balance** and **transactions**.
2. Saves the transactions to an **Excel file** locally.
3. Uploads the Excel file to your **Microsoft account (OneDrive / Excel Online)**, automatically organizing files by **year and month**.

---

## Features

- Reads PDFs and extracts transactions dynamically.
- Saves a clean `.xlsx` file using `pandas`.
- Automatically creates `/BASE_FOLDER/<year>/<month>/` folders in OneDrive if they don’t exist.
- Use a `.env` file to store your credentials.
- Provides a direct link to open the uploaded Excel file in Excel Online.

---

## Requirements

- Python 3.8+
- Packages:
  ```bash
  pip install pandas PyPDF2 requests python-dotenv openpyxl
