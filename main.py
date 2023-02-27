import pandas as pd

import sheetslib as sl
import sheetsformatlib as sfl

from csvhandler import get_csvs_from_input_folder


import time
# from googleapiclient.errors import HttpError

# The ID of the main spreadsheet
spreadsheet_id = '1TglTVF6BwgNR01v55VEQTXuXXkRRGkyOAIT58GCOtZs'

if __name__ == '__main__':
    creds, service = sl.authorize()

print(creds, service)

# Read CSV files in the "input/" directory
csv_files = get_csvs_from_input_folder()
print(csv_files)

dropped_columns = ["#","Country","CPC","CPS","Parent Keyword","Last Update","SERP Features", "Traffic potential"]
relative_difficulty_formula = '=IF(AND(INDIRECT("RC[-3]", FALSE)>=0,INDIRECT("RC[-3]", FALSE)<34),"Low",IF(AND(INDIRECT("RC[-3]", FALSE)>=34,INDIRECT("RC[-3]", FALSE)<67),"Medium",IF(AND(INDIRECT("RC[-3]", FALSE)>=67,INDIRECT("RC[-3]", FALSE)<=100),"High","")))'

all_keywords_df = None
for file in csv_files:
    df = pd.read_csv(file)
    df.drop(dropped_columns, axis="columns", inplace=True)
    print(df.info())

    # Clean up "NaN"
    df["Volume"].fillna(10, inplace=True)
    df.fillna('', inplace=True)

    # Add Relative Difficulty column
    df["Relative Difficulty"] = relative_difficulty_formula

    if all_keywords_df is None:
        all_keywords_df = df
    else:
        all_keywords_df = pd.concat([all_keywords_df,df])

    # Take top keyword and use it as the sheet title
    #print(df["Keyword"][0].title())
    title = df["Keyword"][0].title()
    sheet_id = sl.create_sheet_or_get_sheet_id(spreadsheet_id, title, service)

    body = [df.columns.values.tolist()] + df.values.tolist()

    #print(body)

    range = title + "!A1:E" + str(len(df)+1)

    sl.update_values(spreadsheet_id, range, "USER_ENTERED", body, service)

    sfl.align_vertical_middle_all(spreadsheet_id, sheet_id, service)
    sfl.bold_center_header_row(spreadsheet_id, sheet_id, service)
    sfl.color_header_row(spreadsheet_id, sheet_id, service, 5)
    sfl.center_columns(spreadsheet_id, sheet_id, service, 1, 5)
    sfl.auto_resize_column(spreadsheet_id, sheet_id, service, 1)
    sfl.auto_resize_column(spreadsheet_id, sheet_id, service, 5)
    sfl.relative_difficulty_conditional_formatting(spreadsheet_id, sheet_id, service, len(df))

    time.sleep(5)

print("***All Keywords Including Duplicates***")
print(all_keywords_df.info())

all_keywords_df.drop_duplicates(subset=['Keyword'], inplace=True)
all_keywords_df.sort_values(by=['Volume', 'Global volume'], ascending=[False, False], inplace=True)

print("***All Keywords Without Duplicates***")
print(all_keywords_df.info())

sheet_id = sl.create_sheet_or_get_sheet_id(spreadsheet_id, "All Keywords", service)
body = [all_keywords_df.columns.values.tolist()] + all_keywords_df.values.tolist()
range = "All Keywords" + "!A1:E" + str(len(all_keywords_df)+1)
sl.update_values(spreadsheet_id, range, "USER_ENTERED", body, service)
sfl.align_vertical_middle_all(spreadsheet_id, sheet_id, service)
sfl.bold_center_header_row(spreadsheet_id, sheet_id, service)
sfl.color_header_row(spreadsheet_id, sheet_id, service, 5)
sfl.center_columns(spreadsheet_id, sheet_id, service, 1, 5)
sfl.auto_resize_column(spreadsheet_id, sheet_id, service, 1)
sfl.auto_resize_column(spreadsheet_id, sheet_id, service, 5)
sfl.relative_difficulty_conditional_formatting(spreadsheet_id, sheet_id, service, len(all_keywords_df))