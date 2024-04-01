import gspread

# Authenticate (no need for full edit permissions)
gc = gspread.service_account(filename="google_cloud_key.json")

# Open the publicly shared Google Sheet by URL
sheet_url = "https://docs.google.com/spreadsheets/d/1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s/edit"
spreadsheet = gc.open_by_url(sheet_url)

# Access a specific worksheet (sub-sheet) by name
worksheet_name = "list"
worksheet = spreadsheet.worksheet(worksheet_name)

# Read data from the worksheet
data = worksheet.get_all_values()
print(data)