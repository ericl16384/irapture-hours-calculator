import requests

def getGoogleSheet(spreadsheet_id, outFile, gid=0, csv=True):
    if csv:
        mode = "export?format=csv"
    else:
        mode = "edit"
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/{mode}#gid={gid}"

    response = requests.get(url)
    assert response.status_code == 200, f"Error downloading Google Sheet: {response.status_code}"

    with open(outFile, "wb") as f:
        f.write(response.content)
    # print(f"CSV saved to {outFile}")

filepath = getGoogleSheet("1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s", "google_sheet.csv")

# sys.exit(0); ## success