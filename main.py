print("loading libraries")

# import json
import time, datetime
import os.path

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
    if len(sheets) == 3:
        break


def get_time_from_user(msg):
    valid = False
    while not valid:
        ans = input(msg)
        try:
            dt_object = datetime.datetime.strptime(ans, "%Y-%m-%d %H:%M:%S")
            return dt_object.timestamp()
        except ValueError:
            pass
        except OSError:
            pass
        print()
        print("please enter in this format: YYYY-MM-DD HH:MM:SS")


print(f"{len(sheets)} datasheets loaded")
print()
print("please enter timeframe")
start = get_time_from_user("start:\t")
end = get_time_from_user("end:\t")

print()
print(f"start: {start}")
print(f"end: {end}")

assert start < end





    # print(json.dumps(sheets[list], indent=2))

# with open("dumps.json", "w") as f:
#     f.write(json.dumps(sheets, indent=2))



# with open("index.json", "w") as f:
#     f.write(json.dumps(index, indent=2))

# with open("dump.json", "w") as f:
#     f.write(json.dumps(sh.worksheet(index[1][1]).get_all_values(), indent=2))
# print(data)