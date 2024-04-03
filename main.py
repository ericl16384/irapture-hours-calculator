print("loading libraries")

import time, datetime
import os.path
import csv

import gspread

assert os.path.isfile("./google_cloud_key.json"), "Missing Google Cloud key. Contact Eric Lewis."


print("authenticating")

# Authenticate (no need for full edit permissions)
gc = gspread.service_account(filename="google_cloud_key.json")

sh = gc.open_by_key("1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s")

sheets_list = sh.worksheets()



# if os.path.isfile("download_history.json"):
#     with open("download_history.json", "r") as f:
#         history = json.loads(f.read())
# else:
#     history = {}

# for s in sheets_list:
#     if s.id not in history.keys():
#         history[s.id] = {
#             "title": s.title,
#             "last_download": 0
#         }

# def save_history():
#     with open("download_history.json", "w") as f:
#         f.write(json.dumps(history, indent=2))
# save_history()



# print("working here")
# input()


# to_download = []
# for gid in history:
#     if history[gid]["last_download"] == 0:
#         to_download.append(gid)
# print(f"downloading {len(to_download)} datasheets")

# if not os.path.isdir("downloads"):
#     os.mkdir("downloads")

# base_download_delay = 0.001
# download_delay = base_download_delay
# def download_sheet(gid):
#     sh.get_worksheet(s).get_all_values()





# def refresh_downloads():
#     delay = 1
#     for gid in history:
#         sh.get_worksheet(s).get_all_values()
# refresh_downloads()



# with open("downloads/test.txt", "w") as f:
#     f.write("Hello World!")

# input()

# print(sheets_list)
# assert False

# sh.values_batch_get()


def load_sheet_data(title):
    base_delay = 0.001
    delay = base_delay

    sheet = None
    data = None
    while data == None:
        try:
            if sheet:
                data = sheet.get_all_values()
            else:
                sheet = sh.worksheet(title)
        except:
            time.sleep(delay)
            delay *= 2
            continue
        delay = base_delay
    
    return data


print("loading index list")

index = sh.worksheet("list").get_all_values()

sheets = {}
for organization, list in index:
    if list == "list":
        continue

    print(list, "\t", organization)

    sheets[list] = load_sheet_data(list)

    # debug
    if len(sheets) == 2:
        break


def get_time_from_user(msg):
    valid = False
    while not valid:
        ans = input(msg)
        if not ans:
            return None
        for format in (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%Y-%m-%d %H",
            "%Y-%m-%d",
            "%Y-%m",
            "%Y",
        ):
            try:
                return datetime.datetime.strptime(ans, format).timestamp()
            except ValueError:
                pass
        print()
        print("please enter in this format (ENTER to autocomplete):")
        print("\tYYYY-MM-DD HH:MM:SS")


print()
print(f"{len(sheets)} datasheets loaded")
print()


# import json
# with open("dumps.json", "w") as f:
#     f.write(json.dumps(sheets, indent=2))


while True:
    print("please enter timeframe")
    start = get_time_from_user("start:\t")
    end = get_time_from_user("end:\t")

    if not start:
        start = 0
    if not end:
        end = 2**64

    # print()
    # print(f"start: {start}")
    # print(f"end: {end}")

    if start >= end:
        print("start time must be before end time")
        continue

    print("calculating...")


    people = {}

    for title, sheet in sheets.items():
        if title.startswith("."):
            continue

        for i, row in enumerate(sheet):
            if i < 6:
                continue

            # if time range is valid

            person = row[1]
            hours = row[4]

            if not person or not hours:
                continue

            hours = float(hours)

            if person not in people:
                people[person] = 0
            people[person] += hours

            # print(row)
            # input()

    print(people)

    print()
    print("COMPLETE")
    print("saved to TODO")
    print()





    # print(json.dumps(sheets[list], indent=2))

# with open("dumps.json", "w") as f:
#     f.write(json.dumps(sheets, indent=2))



# with open("index.json", "w") as f:
#     f.write(json.dumps(index, indent=2))

# with open("dump.json", "w") as f:
#     f.write(json.dumps(sh.worksheet(index[1][1]).get_all_values(), indent=2))
# print(data)