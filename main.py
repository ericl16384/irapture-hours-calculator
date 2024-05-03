



DEBUG_MODE = False


# DEBUG
if DEBUG_MODE:
    import random
    print("DEBUG MODE")
    input("> ")



print("loading libraries")

from datetime import datetime
import time
import os.path
import csv
import json

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

# def load_sheet_data(title):
#     base_delay = 1
#     delay = base_delay

#     sheet = None
#     data = None
#     while data == None:
#         try:
#             if sheet:
#                 data = sheet.get_all_values()
#             else:
#                 sheet = sh.worksheet(title)
#         except gspread.exceptions.WorksheetNotFound:
#             print("Worksheet not found (from <list>):")
#             print(f"  Hour ID: \"{title}\"")
#             input("press ENTER to skip error")
#             return []
#         except gspread.exceptions.APIError:
#             # this seems to be when there are too many requests
#             print("waiting", delay, "seconds")
#             time.sleep(delay)
#             delay *= 2
#             continue
#         delay = base_delay
    
#     return data


# print("loading index list")
# index = sh.worksheet("list").get_all_values()

print("loading datasheets from list")
sheets = {}
# for i, line in enumerate(index):
#     hour_id, organization = line

#     # if list == "list":
#     if i < 3: # skip the refresh button and header
#         continue

#     # print(" ", list, "\t", organization)
#     print(line)
#     # time.sleep(0.01)

#     sheets[hour_id] = load_sheet_data(hour_id)
for worksheet in sh.worksheets():
    title = worksheet.title
    if title == "Index":
        continue
    if title.startswith("."):
        continue

    if DEBUG_MODE and random.random() > 0.1:
        continue

    print(" ", title)


    base_delay = 1
    delay = base_delay

    # sheet = None
    sheet = worksheet
    data = None
    while data == None:
        try:
            # if sheet:
            data = sheet.get_all_values()
            # else:
            #     sheet = sh.worksheet(title)
        # except gspread.exceptions.WorksheetNotFound:
        #     print("Worksheet not found (from <list>):")
        #     print(f"  Hour ID: \"{title}\"")
        #     input("press ENTER to skip error")
        #     return []
        except gspread.exceptions.APIError:
            # this seems to be when there are too many requests
            print("waiting", delay, "seconds")
            time.sleep(delay)
            delay *= 2
            continue
        delay = base_delay

    # return data
    sheets[title] = data


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
print(datetime.timestamp(datetime.now()) - PROGRAM_START_TIMESTAMP, "seconds elapsed")
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
    
    if am_pm == "AM":
        pass
    elif am_pm == "PM":
        # add 12 hours
        t += 60*60*12
    else:
        # assert False
        t_str = f"{date} {time}"

    for format in (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M"
    ):
        try:
            t += datetime.strptime(t_str, format).timestamp()
            break
        except ValueError:
            pass
        except OSError:
            # print(f"WARNING: OSError in strptime; ignoring ({date}) ({time})")
            return f"WARNING: OSError in strptime; ignoring ({date}) ({time})"

    if t == 0:
        # print(f"WARNING: bad time stamp detected; ignoring ({date}) ({time})")
        return f"WARNING: bad time stamp detected; ignoring ({date}) ({time})"

    return t

def coord_to_cell_id(column_x, row_y):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    return letters[column_x] + str(row_y+1)


with open("output.json", "w") as f:
    f.write(json.dumps(sheets, indent=2))


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


    # datasheet_id, cell, msg
    spreadsheet_errors = []


    hours_by_sheet = {}
    totals_by_person = {}

    for title, sheet in sheets.items():
        if title.startswith("."):
            continue

        hours_by_sheet[title] = {}

        for i, row in enumerate(sheet):
            # header
            if i < 6:
                continue

            if "".join(row) == "":
                # spreadsheet_errors.append((title,
                #     f"{coord_to_cell_id(0, i)}:{coord_to_cell_id(4, i)}",
                #     "WARNING: empty sheet row; ignoring"
                # ))
                continue
            if "".join(row[:4]) == "":
                try:
                    if float(row[4]) == 0:
                        continue
                except:
                    pass

            if len(row) < 5 or "" in row[:5]:
                spreadsheet_errors.append((title,
                    f"{coord_to_cell_id(0, i)}:{coord_to_cell_id(4, i)}",
                    "WARNING: incomplete sheet row; ignoring"
                ))
                continue

            # if "" in row[:5]:
            #     continue

            date = row[0]
            person = row[1]
            in_time = row[2]
            out_time = row[3]
            hours = row[4]

            # if not in_time:
            #     in_time = "12:00:00 PM"
            # if not out_time:
            #     out_time = "12:00:00 PM"

            in_stamp = get_time_from_sheet(date, in_time)
            out_stamp = get_time_from_sheet(date, out_time)

            # handle errors
            if isinstance(in_stamp, str):
                spreadsheet_errors.append((
                    title,
                    f"{coord_to_cell_id(0, i)}, {coord_to_cell_id(2, i)}",
                    in_stamp
                ))
                continue
            if isinstance(out_stamp, str):
                spreadsheet_errors.append((
                    title,
                    f"{coord_to_cell_id(0, i)}, {coord_to_cell_id(3, i)}",
                    out_stamp
                ))
                continue

            if in_stamp < start or out_stamp > end:
                # means that it was excluded by the time selection
                continue

            # print(f"{date}\t{in_time}\t{out_time}")

            try:
                hours = float(hours)
            except ValueError:
                # print(f"WARNING: bad hours format; ignoring ({hours})")
                spreadsheet_errors.append((
                    title, coord_to_cell_id(4, i),
                    f"WARNING: bad hours format; ignoring ({hours})"
                ))
                continue

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
    export_table.append(["client\\dev"])
    for client in sheets:
        export_table.append([client])

    # header
    for person in totals_by_person:
        export_table[0].append(person)

    # hours
    for i, row in enumerate(export_table):
        if i == 0:
            continue

        title = row[0]
        for person in export_table[0][1:]:

            if title not in hours_by_sheet:
                export_table[i].append("")
                continue


            if person in hours_by_sheet[title]:
                h = hours_by_sheet[title][person]
            else:
                h = 0
            export_table[i].append(h)

    # person totals
    export_table.append([""]*len(export_table[0]))
    export_table[-1][0] = "TOTAL"
    for i, person in enumerate(export_table[0]):
        if i == 0:
            continue
        export_table[-1][i] = totals_by_person[person]
        # assert totals_by_person[person] == sum([row[i] for row in export_table[1:]])

    # sheet (row) totals
    export_table[0].append("TOTAL")
    for i, row in enumerate(export_table):
        if i == 0:
            continue

        # try:
        # print(export_table[i][1:])
        v = sum(export_table[i][1:])
        # except TypeError:
        #     v = ""
        export_table[i].append(v)


    # error log (added above table)
    print()
    print()
    if(spreadsheet_errors):
        print(f"{len(spreadsheet_errors)} errors ignored (hours not processed and invalidated)")
        spreadsheet_errors = [["Datasheet", "Cell", "Error message"]] + spreadsheet_errors
    else:
        print("No spreadsheet errors! Great job!")
        spreadsheet_errors = [["No spreadsheet errors detected! Great!"]]
    export_table = spreadsheet_errors + [[], []] + export_table


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