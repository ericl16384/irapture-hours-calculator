import requests

csv_url = "https://docs.google.com/spreadsheets/d/1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s/edit#gid=287905941"
response = requests.get(url=csv_url)
with open("google_sheet.csv", "wb") as file:
    file.write(response.content)