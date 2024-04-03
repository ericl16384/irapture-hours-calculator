print("loading libraries")

from datetime import datetime
import time
import os.path
import csv

import gspread

assert os.path.isfile("./google_cloud_key.json"), "Missing Google Cloud key. Contact Eric Lewis."


PROGRAM_START_TIMESTAMP = datetime.timestamp(datetime.now())


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
    base_delay = 1
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
            print("waiting", delay, "seconds")
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

    # # debug
    # if len(sheets) == 5:
    #     break


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
                return datetime.strptime(ans, format).timestamp()
            except ValueError:
                pass
        print()
        print("please enter in this format (ENTER to autocomplete):")
        print("\tYYYY-MM-DD HH:MM:SS")


print()
print(datetime.timestamp(datetime.now() - PROGRAM_START_TIMESTAMP), "seconds elapsed")
print(len(sheets), "datasheets loaded")
print()
print("HINT: DO NOT CLOSE PROGRAM, BUT INSTEAD KEEP IT OPEN BETWEEN SEARCHES")
print()


# import json
# with open("dumps.json", "w") as f:
#     f.write(json.dumps(sheets, indent=2))


def get_time_from_sheet(date, time):
    t_str = f"{date} {time}"
    am_pm = time[-2:]
    t_str = t_str[:-3]

    t = 0
    for format in (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M"
    ):
        try:
            t += datetime.strptime(t_str, format).timestamp()
            break
        except ValueError:
            pass

    if t == 0:
        print(f"WARNING: bad time stamp detected; ignoring ({date}, {time})")

    if am_pm == "AM":
        pass
    elif am_pm == "PM":
        # add 12 hours
        t += 60*60*12
    else:
        assert False
    return t


while True:
    print("please enter timeframe")
    start = get_time_from_user("start:\t")
    end = get_time_from_user("end:\t")

    if not start:
        start = 0
    if not end:
        end = 2**32

    # print()
    # print(f"start: {start}")
    # print(f"end: {end}")

    if start >= end:
        print("start time must be before end time")
        continue

    print("calculating...")


    hours_by_sheet = {}
    totals_by_person = {}

    for title, sheet in sheets.items():
        if title.startswith("."):
            continue

        hours_by_sheet[title] = {}

        for i, row in enumerate(sheet):
            if i < 6:
                continue

            if "" in row[:5]:
                continue

            date = row[0]
            person = row[1]
            in_time = row[2]
            out_time = row[3]
            hours = row[4]

            in_stamp = get_time_from_sheet(date, in_time)
            out_stamp = get_time_from_sheet(date, out_time)

            if in_stamp < start or out_stamp > end:
                # means that it was excluded by the time selection
                continue

            # print(f"{date}\t{in_time}\t{out_time}")

            hours = float(hours)

            if person not in hours_by_sheet[title]:
                hours_by_sheet[title][person] = 0
            hours_by_sheet[title][person] += hours

            if person not in totals_by_person:
                totals_by_person[person] = 0
            totals_by_person[person] += hours

            # print(row)
            # input()


    # print(hours_by_sheet)

    export_table = []

    # Organization, list
    for row in index:
        export_table.append(row[:2])

    # header
    for person in totals_by_person:
        export_table[0].append(person)

    # hours
    for i, row in enumerate(export_table):
        if i == 0:
            continue

        name, title = row
        for person in export_table[0][2:]:

            # debug
            if title not in hours_by_sheet:
                export_table[i].append("")
                continue


            if person in hours_by_sheet[title]:
                h = hours_by_sheet[title][person]
            else:
                h = 0
            export_table[i].append(h)

    # person totals
    export_table.append(["TOTAL", ""])
    for person in export_table[0][2:]:
        export_table[-1].append(totals_by_person[person])
    
    # sheet (row) totals
    export_table[0].append("TOTAL")
    for i, row in enumerate(export_table):
        if i < 1:
            continue

        try:
            v = sum(export_table[i][2:])
        except TypeError:
            v = ""
        export_table[i].append(v)


    # info (added above table)
    info_table = [
        ["https://github.com/ericl16384/irapture-hours-calculator"],
        ["https://docs.google.com/spreadsheets/d/1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s/edit#gid=287905941"],
        [],
        ["Program run time",
         datetime.fromtimestamp(PROGRAM_START_TIMESTAMP).strftime("%m/%d/%Y"),
         datetime.fromtimestamp(PROGRAM_START_TIMESTAMP).strftime("%I:%M:%S %p"),
        ],
        [],
        ["Filter start time",
         datetime.fromtimestamp(start).strftime("%m/%d/%Y"),
         datetime.fromtimestamp(start).strftime("%I:%M:%S %p")
        ],
        ["Filter end time",
         datetime.fromtimestamp(end).strftime("%m/%d/%Y"),
         datetime.fromtimestamp(end).strftime("%I:%M:%S %p")
        ],
        [],
        []
    ]
    export_table = info_table + export_table


    savefile = "output.csv"
    with open(savefile, "w") as f:
        csv.writer(f, lineterminator="\n").writerows(export_table)

    print()
    print()
    print(f"output saved to\t{savefile}")
    print()
    print("press ENTER to continue")
    print()
    input()





    # print(json.dumps(sheets[list], indent=2))

# with open("dumps.json", "w") as f:
#     f.write(json.dumps(sheets, indent=2))



# with open("index.json", "w") as f:
#     f.write(json.dumps(index, indent=2))

# with open("dump.json", "w") as f:
#     f.write(json.dumps(sh.worksheet(index[1][1]).get_all_values(), indent=2))
# print(data)