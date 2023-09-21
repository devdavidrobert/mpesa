import pandas

sheet_id = '1p0UG58Tmzhwkvk6BNzc0OedRUD_wy8ALaEoJP1ifguQ'

# Read the Google Sheets document as a DataFrame
df = pandas.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv")

# Access the first column (assuming it contains passwords)
passwords_column = df.iloc[:, 0]

# Convert the passwords column to a list
passwords = passwords_column.tolist()

print(passwords)
