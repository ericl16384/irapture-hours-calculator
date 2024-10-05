



    





print("loading libraries")

from datetime import datetime
import csv
import json
import numpy as np
import os.path
import random
import sys
import time


# DEBUG_MODE = False
# if "--debug" in sys.argv:
#     DEBUG_MODE = True
#     print("""
#                    _      _                 
#                   | |    | |                
#   ______ ______ __| | ___| |__  _   _  __ _ 
#  |______|______/ _` |/ _ \ '_ \| | | |/ _` |
#               | (_| |  __/ |_) | |_| | (_| |
#                \__,_|\___|_.__/ \__,_|\__, |
#                                        __/ |
#                                       |___/ 
# """)
    
# USE_GSPREAD = False
# if "--gspread" in sys.argv:
#     USE_GSPREAD = True



def exit_on_user_input(*msgs):
    print()
    if msgs:
        print()
        for msg in msgs: print(msg)
    print()
    print("Press ENTER to exit the program.")
    input()
    sys.exit()

def exit_prompting_installation_of_modules():
    exit_on_user_input(
        "Required libraries not installed. Please close this program, run install_python_modules.bat and try again.",
        f"({os.getcwd()}\\install_python_modules.bat)"
    )



try:
    # import gspread
    import pandas as pd
except ModuleNotFoundError:
    exit_prompting_installation_of_modules()







# # DEBUG
# if DEBUG_MODE:
#     import random
#     print("DEBUG MODE (random sample of data, not complete)")
#     print("press ENTER")
#     input("> ")



    


PROGRAM_START_TIMESTAMP = datetime.timestamp(datetime.now())
    




# assert os.path.isfile("./google_cloud_key.json"), "Missing Google Cloud key. Contact Eric Lewis."

# def manage_request_overload(lambda_web_api_request):
#     delay = 1
#     while True:
#         try:
#             return lambda_web_api_request()
#         except gspread.exceptions.APIError as err:
#             # this seems to be when there are too many requests
#             print("waiting", delay, "seconds", "due to error:")
#             print(" ", str(err))
#             time.sleep(delay)
#             delay *= 2


# print("authenticating")

# # Authenticate (no need for full edit permissions)
# gc = gspread.service_account(filename="google_cloud_key.json")

# sh = manage_request_overload(
#     lambda: gc.open_by_key("1WjSQnzFIqTKtnKVx5O19OxBul2MoiAgyy1iYgFn495s")
# )


# sheets_list = sh.worksheets()


# print("loading datasheets from list")
# sheets = {}
# for worksheet in sh.worksheets():
#     title = worksheet.title
#     if title == "Index":
#         continue
#     if title.startswith("."):
#         continue

#     if DEBUG_MODE and random.random() > 0.1:
#         print("        SKIPPED", title)
#         continue


#     print(" ", title)


#     data = manage_request_overload(
#         lambda: worksheet.get_all_values()
#     )

#     # return data
#     sheets[title] = data





print("loading worksheets from Excel file")

try:
    # sheets = pd.read_excel("Hours.xlsx")
    excel_worksheets = pd.ExcelFile("Hours.xlsx")
except ImportError:
    exit_prompting_installation_of_modules()
except FileNotFoundError:
    exit_on_user_input(
        "Missing file Hours.xlsx. Please download the Hours Google sheet in Excel format and put it in this folder:",
        os.getcwd()
    )

sheets = {}
for sheet_name in excel_worksheets.sheet_names:
    if sheet_name.startswith("."):
        print(f"skipping {sheet_name}")
        continue
    print(f"parsing  {sheet_name}")

    sheets[sheet_name] = excel_worksheets.parse(sheet_name,
        index_col=None, header=None, dtype=str, na_filter=False
    ).replace({np.nan: ""}).values.tolist()


def get_time_from_user(msg):
    valid = False
    while not valid:
        ans = input(msg)
        if not ans:
            return None
        for format in (
            # "%Y-%m-%d %H:%M:%S",
            # "%Y-%m-%d %H:%M",
            # "%Y-%m-%d %H",
            # "%Y-%m-%d",
            # "%Y-%m",
            # "%Y",
        ):
            try:
                return datetime.strptime(ans, format).timestamp()
            except ValueError:
                pass
        print()
        print("please enter in this format (ENTER to autocomplete):")
        print("\tYYYY-MM-DD HH:MM:SS")


PROGRAM_RUN_SECONDS = datetime.timestamp(datetime.now()) - PROGRAM_START_TIMESTAMP


print()
print(PROGRAM_RUN_SECONDS, "seconds elapsed")
# print(len(sheets), "datasheets loaded")
print(len(sheets), "worksheets loaded")
# print()
# print("HINT: DO NOT CLOSE PROGRAM, BUT INSTEAD KEEP IT OPEN BETWEEN SEARCHES.")
print()

# if DEBUG_MODE:
#     print("WARNING: DEBUG MODE ENABLED. RUN PROGRAM WITHOUT DEBUG MODE FOR ACCURATE RESULTS.")
#     print()


# import json
# with open("dumps.json", "w") as f:
#     f.write(json.dumps(sheets, indent=2))


def get_time_from_sheet(date, time):
    t_str = f"{date} {time}"
    # am_pm = time[-2:]
    # t_str = t_str[:-3]

    t = 0
    
    # if am_pm == "AM":
    #     pass
    # elif am_pm == "PM":
    #     # add 12 hours
    #     t += 60*60*12
    # else:
    #     # assert False
    #     t_str = f"{date} {time}"

    for format in (
        # "%m/%d/%Y %H:%M:%S %p",
        # "%m/%d/%Y %H:%M %p",

        r"%Y-%m-%d 00:00:00 %H:%M:%S",
        r"%Y-%m-%d 00:00:00 %H:%M:%S.%f",
    ):
        try:
            t += datetime.strptime(t_str, format).timestamp()
            break
        except ValueError:
            pass
        except OSError:
            # print(f"WARNING: OSError in strptime; ignoring ({date}) ({time})")
            return f"WARNING: OSError in strptime; ignoring `{t_str}`"

    if t == 0:
        # print(f"WARNING: bad time stamp detected; ignoring ({date}) ({time})")
        return f"WARNING: bad time stamp detected; ignoring `{t_str}`"

    return t

def coord_to_cell_id(column_x, row_y):
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    return letters[column_x] + str(row_y+1)


# with open("output.json", "w") as f:
#     f.write(json.dumps(sheets, indent=2))


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


    invoice_categorization = {}


    hours_by_sheet = {}
    totals_by_person = {}

    invoices_by_sheet = {}
    projects_by_sheet = {}

    time_per_month_by_sheet = {}
    remaining_hours_by_sheet = {}


    for title, sheet in sheets.items():

        hours_by_sheet[title] = {}

        if len(sheet) < 6:
            spreadsheet_errors.append((
                title,
                "",
                "Sheet has less than six rows (no header info)"
            ))
            continue

        for i, row in enumerate(sheet):

            # if isinstance(row, float):
            #     continue

            # print(title, i, row)

            # add "Time per month" from E3:E4
            if i == 2:
                if len(row) < 5:
                    remaining_hours_by_sheet[title] = "row 3 too short"
                elif row[4] == "Time per month":
                    time_per_month_by_sheet[title] = "loading..."
                else:
                    time_per_month_by_sheet[title] = "wrong format"
            if i == 3:
                if time_per_month_by_sheet[title] == "row 3 too short":
                    pass
                elif time_per_month_by_sheet[title] == "loading...":
                    time_per_month_by_sheet[title] = row[4]
                    if time_per_month_by_sheet[title] == "":
                        time_per_month_by_sheet[title] = "MISSING"
                        spreadsheet_errors.append((
                            title,
                            "E3:E4",
                            "Value for `Time per month` in header is missing"
                        ))
                elif time_per_month_by_sheet[title] == "wrong format":
                    pass
                else:
                    assert False
            # add "Remaining hours" from I3:I4
            if i == 2:
                if len(row) < 9:
                    remaining_hours_by_sheet[title] = "row 3 too short"
                elif row[8] == "Remaining hours":
                    remaining_hours_by_sheet[title] = "loading..."
                else:
                    remaining_hours_by_sheet[title] = "wrong format"
            if i == 3:
                if remaining_hours_by_sheet[title] == "row 3 too short":
                    pass
                elif remaining_hours_by_sheet[title] == "loading...":
                    remaining_hours_by_sheet[title] = row[8]
                    if remaining_hours_by_sheet[title] == "":
                        remaining_hours_by_sheet[title] = "MISSING"
                        spreadsheet_errors.append((
                            title,
                            "I3:I4",
                            "Value for `Remaining hours` in header is missing"
                        ))
                elif remaining_hours_by_sheet[title] == "wrong format":
                    pass
                else:
                    assert False

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
                spreadsheet_errors.append((
                    title,
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

            # print()
            # print(start, datetime.fromtimestamp(start).strftime("%m/%d/%Y %I:%M:%S %p"))
            # print(in_stamp, datetime.fromtimestamp(in_stamp).strftime("%m/%d/%Y %I:%M:%S %p"))
            # print(out_stamp, datetime.fromtimestamp(out_stamp).strftime("%m/%d/%Y %I:%M:%S %p"))
            # print(end, datetime.fromtimestamp(end).strftime("%m/%d/%Y %I:%M:%S %p"))

            if in_stamp < start or out_stamp > end:
                # means that it was excluded by the time selection
                continue

            # print(f"{date}\t{in_time}\t{out_time}")

            try:
                hours = float(hours)
            except ValueError:
                # print(f"WARNING: bad hours format; ignoring ({hours})")
                spreadsheet_errors.append((
                    title,
                    coord_to_cell_id(4, i),
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

            # record which clients are VP vs Hourly for Heather's billing
            invoice = row[5]
            if invoice not in invoice_categorization:
                invoice_categorization[invoice] = []
            invoice_categorization[invoice].append((title, i))

            # determine if project column is vp or hourly
            if title not in invoices_by_sheet:
                invoices_by_sheet[title] = set()
            invoices_by_sheet[title].add(row[5].lower())
            if title not in projects_by_sheet:
                projects_by_sheet[title] = set()
            projects_by_sheet[title].add(row[6].lower())


    # record which clients are VP vs Hourly for Heather's billing
    invoice_export_table = [("Invoice msg", "Sheet title", "Row")]
    categories = ["", "HOURLY"]
    categories.extend(invoice_categorization.keys())
    for category in categories:
        if category not in invoice_categorization:
            continue
        # sequence_length = 1
        for title, row in invoice_categorization[category]:
            # if invoice_export_table[-1][:2] == row:
            #     sequence_length += 1
            #     continue
            # else:
            #     sequence_length = 1
            # row = (category, title, sequence_length)
            row = (category, title, row+1)
            invoice_export_table.append(row)
        del invoice_categorization[category]
    invoice_tracker_file = "invoice_types.csv"
    with open(invoice_tracker_file, "w") as f:
        csv.writer(f, lineterminator="\n").writerows(invoice_export_table)





    # first columns

    export_table = [["client", "INV", "Project", "Time per month", "Remaining hours"]]
    # for client in sheets:
    #     export_table.append([client])

    # for i, row in enumerate(export_table):
    #     if i == 0: continue
    #     client = row[0]
    #     skip = client == "TOTAL"
    for client in sheets:
        row = [client]
        export_table.append(row)

        # INV
        if client in invoices_by_sheet:
            row.append(",".join(sorted(list(invoices_by_sheet[client]))))
        else:
            row.append("")

        # Project
        if client in projects_by_sheet:
            row.append(",".join(sorted(list(projects_by_sheet[client]))))
        else:
            row.append("")
        
        # Time per month
        if client in time_per_month_by_sheet:
            row.append(time_per_month_by_sheet[client])
        else:
            row.append("")
        
        # Remaining hours
        if client in remaining_hours_by_sheet:
            row.append(remaining_hours_by_sheet[client])
        else:
            row.append("")
        
    static_columns = len(export_table[0])

    # header
    for person in totals_by_person:
        export_table[0].append(person)

    # hours
    for i, row in enumerate(export_table):
        if i == 0:
            continue

        title = row[0]
        for person in export_table[0][static_columns:]:

            if title not in hours_by_sheet:
                export_table[i].append("")
                continue


            if person in hours_by_sheet[title]:
                h = hours_by_sheet[title][person]
            else:
                h = 0
            export_table[i].append(h)

    # person (column) totals
    export_table.append([""]*len(export_table[0]))
    export_table[-1][0] = "TOTAL"
    for i, person in enumerate(export_table[0]):
        if i < static_columns:
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
        v = sum(export_table[i][static_columns:])
        # except TypeError:
        #     v = ""
        export_table[i].append(v)

    # error log (added above table)
    print()
    print()
    if(spreadsheet_errors):
        print(f"{len(spreadsheet_errors)} errors logged and ignored (hours not processed and invalidated)")
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
        [
            "Program run time",
            datetime.fromtimestamp(PROGRAM_START_TIMESTAMP).strftime("%m/%d/%Y"),
            datetime.fromtimestamp(PROGRAM_START_TIMESTAMP).strftime("%I:%M:%S %p")
        ],
        [
            "Program load time elapsed (seconds)",
            PROGRAM_RUN_SECONDS
        ],
        [],
        [
            "Filter start time",
            datetime.fromtimestamp(start).strftime("%m/%d/%Y"),
            datetime.fromtimestamp(start).strftime("%I:%M:%S %p")
        ],
        [   "Filter end time",
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
    print(f"invoice categorization saved to\t{invoice_tracker_file}")
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