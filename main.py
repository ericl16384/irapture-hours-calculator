import json
import os.path
import gspread

assert os.path.isfile("./google_cloud_key.json"), "Missing Google Cloud key. Contact Eric Lewis."

# Authenticate (no need for full edit permissions)
gc = gspread.service_account(filename="google_cloud_key.json")

# Open the publicly shared Google Sheet by URL
sheet_url = "https://docs.google.com/spreadsheets/d/1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s/edit"
spreadsheet = gc.open_by_url(sheet_url)

index = spreadsheet.worksheet("list").get_all_values()



with open("index.json", "w") as f:
    f.write(json.dumps(index, indent=2))

with open("dump.json", "w") as f:
    f.write(json.dumps(spreadsheet.worksheet(index[1][1]).get_all_values(), indent=2))
# print(data)