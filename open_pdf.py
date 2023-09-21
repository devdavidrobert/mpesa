import os
import io
from googleapiclient.discovery import build
from google.oauth2 import service_account
import PyPDF2
import pandas as pd
import gspread

# Define your Google Drive API credentials file (JSON)
CLIENT_SERVICE_FILE = 'mpesa-script2.json'  # Replace with your credentials file

# Define the folder name where PDF files are located
FOLDER_NAME = 'Mpesa Statement'  # Replace with your folder name

# Define the Google Sheet name containing passwords
SHEET_ID = '1p0UG58Tmzhwkvk6BNzc0OedRUD_wy8ALaEoJP1ifguQ'  # Replace with your Google Sheet name

# Define the scopes for Google Drive and Google Sheets APIs
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
SHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Initialize correctPassword as None (globally)
correctPassword = None

# Create a service account credentials object for Google Drive API
drive_credentials = service_account.Credentials.from_service_account_file(
    CLIENT_SERVICE_FILE, scopes=DRIVE_SCOPES)

# Build the Google Drive API service
drive_service = build('drive', 'v3', credentials=drive_credentials)

# List PDF files in the specified folder
folder_query = f"'1vACT2I0Xwr3kJoY55fz_3zHSPjsHjyvm' in parents and mimeType='application/pdf'"
results = drive_service.files().list(q=folder_query).execute()
pdf_files = results.get('files', [])

if not pdf_files:
    print(f"No PDF files found in the '1vACT2I0Xwr3kJoY55fz_3zHSPjsHjyvm' folder in Google Drive.")
else:
    print(f"PDF files in the '1vACT2I0Xwr3kJoY55fz_3zHSPjsHjyvm' folder in Google Drive:")
    for file in pdf_files:
        print(f"File_Name: {file['name']}")
        print(f"ID: {file['id']}")

        # Get the PDF file content
        file_id = file['id']
        file_name = file['name']
        request = drive_service.files().get_media(fileId=file_id)
        file_content = request.execute()

        # Check if the PDF file is password-protected
        pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_content))

        if pdf_reader.is_encrypted:
            # Fetch passwords from the Google Sheet
            sheet_id = '1p0UG58Tmzhwkvk6BNzc0OedRUD_wy8ALaEoJP1ifguQ'

            # Read the Google Sheets document as a DataFrame
            df = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv", header=None)

            # Access the first column (assuming it contains passwords)
            passwords_column = df.iloc[:, 0]

            # Convert the passwords column to a list
            passwords = [str(password) for password in passwords_column.tolist()]

            # Initialize the Google Sheets client
            gc = gspread.service_account(filename=CLIENT_SERVICE_FILE)

            # Open the Google Sheets document
            worksheet = gc.open_by_key(SHEET_ID).worksheet('pswd')

            # Attempt to decrypt the PDF file using the fetched password
            for password in passwords:
                pdf_reader.decrypt(password.encode())
                if pdf_reader.decrypt(password.encode()):
                    # Password successfully decrypted the PDF
                    correctPassword = password  # Set correctPassword dynamically
                    break

            if 'correctPassword' in globals():
                # Password was successfully set
                print(f"Password: {correctPassword}")
                # Rest of your code here using correctPassword

                # Create a temporary PDF file
                temp_pdf_file = f"{file['name']}_temp.pdf"
                with open(temp_pdf_file, 'wb') as temp_pdf:
                    temp_pdf.write(file_content)

                # Read and print "Customer Name" and "Mobile Number" from the PDF
                pdf_text = ""
                with open(temp_pdf_file, 'rb') as pdf_file:
                    for page in pdf_reader.pages:
                        first_page = pdf_reader.pages[0]
                        pdf_text = first_page.extract_text()

                # Split the text into lines and print the first two lines
                lines = pdf_text.split('\n')

                # Search for "Customer Name" and "Mobile Number" in the extracted text
                customer_name_index = lines[4].split(': ')
                mobile_number_index = lines[5].split(': ')

                if customer_name_index != -1 and mobile_number_index != -1:
                    customer_name = customer_name_index[1]
                    mobile_number = mobile_number_index[1]

                    print(f"Customer Name: {customer_name}")
                    print(f"Mobile Number: {mobile_number}")

                    # Find the row where the password is located
                    password_row = df[df[0] == correctPassword].index[0]  # +2 to account for 0-based index and header row

                    # Update the "Customer Name" column in the corresponding row
                    worksheet.update(f'B{password_row}', customer_name)
                    worksheet.update(f'C{password_row}', mobile_number)
                    worksheet.update(f'D{password_row}', file_id)
                    worksheet.update(f'E{password_row}', file_name)
                    worksheet.update(f'F{password_row}', correctPassword)

                # Clean up the temporary PDF file (optional)
                os.remove(temp_pdf_file)
            else:
                print("No correct password found.")
