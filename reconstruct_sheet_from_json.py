import json, csv

with open("output.json", "r") as f:
    all_sheets = json.loads(f.read())

sheet = all_sheets["ec"]

with open("reconstructed_sheet.csv", "w") as f:
    csv.writer(f, lineterminator="\n").writerows(sheet)