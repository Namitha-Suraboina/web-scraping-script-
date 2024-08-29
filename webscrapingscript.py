'''To scrape data from a Google Sheets document and extract details such as the company website, contact email, and phone number, you'll need to follow these steps:
Access the Google Sheet: Use the Google Sheets API to access and retrieve the data.
Extract the required information: Parse the data to find the company website, contact email, and phone number.
Compile the data: Store the extracted data into a new Excel file using a library like openpyxl or pandas.'''
#Prerequisites
#Install the required libraries:
#pip install gspread pandas openpyxl google-auth
#Create a Google Cloud project and enable the Google Sheets API.
#Set up authentication:
#Download the credentials JSON file from your Google Cloud project.
#Share the Google Sheet with the email address in the credentials file.

'''Steps Explained
Google Sheets Access:

The script authenticates using a service account JSON file.
It accesses the specific Google Sheet using the URL.
Data Extraction:

It retrieves all data from the first worksheet.
The data is stored in a Pandas DataFrame for easier manipulation.
Saving Data:

The script filters the DataFrame to include only the relevant columns.
The filtered data is saved to an Excel file using pandas.'''


import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# Google Sheets setup
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SERVICE_ACCOUNT_FILE = 'path/to/your/credentials.json'  # Replace with your credentials file path

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

gc = gspread.authorize(credentials)

# Access the Google Sheet
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1yFcKDsgXg0E4pBUP6Eqprzk79BFNKwW9/edit?usp=sharing"
spreadsheet = gc.open_by_url(spreadsheet_url)

# Assuming data is in the first sheet
worksheet = spreadsheet.sheet1

# Get all data in the sheet
data = worksheet.get_all_records()

# Convert to a DataFrame for easier processing
df = pd.DataFrame(data)

# Extract relevant columns (adjust according to your sheet structure)
# Replace 'Website', 'Email', 'Phone' with actual column names
required_columns = ['Company Name', 'Website', 'Email', 'Phone']
df = df[required_columns]

# Save the data to a new Excel file
output_file = 'extracted_company_details.xlsx'
df.to_excel(output_file, index=False)

print(f"Data successfully extracted and saved to {output_file}")
