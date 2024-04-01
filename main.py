import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://docs.google.com/spreadsheets/d/1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s/edit#gid=287905941"]
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
docid = "your_spreadsheet_id_here"

client = gspread.authorize(credentials)
spreadsheet = client.open_by_key(docid)

for i, worksheet in enumerate(spreadsheet.worksheets()):
    # Process each worksheet as needed
    # For example, you can access cell values using worksheet.cell(row, col).value
    pass