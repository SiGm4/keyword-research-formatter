import pandas as pd

import sheetslib as sl
import sheetsformatlib as sfl

from csvhandler import get_csvs_from_input_folder

from googleapiclient.errors import HttpError

# The ID of the main spreadsheet
spreadsheet_id = '1vwazr-XfwBhWsA7ArqoQNpOP-hg0BGUluYlMGwFvEwc'

if __name__ == '__main__':
    creds, service = sl.authorize()

print(creds, service)




# Read CSV files in the "input/" directory
csv_files = get_csvs_from_input_folder()
print(csv_files)

dropped_columns = ["#","Country","CPC","CPS","Parent Keyword","Last Update","SERP Features", "Traffic potential"]
relative_difficulty_formula = '=IF(AND(INDIRECT("RC[-3]", FALSE)>=0,INDIRECT("RC[-3]", FALSE)<34),"Low",IF(AND(INDIRECT("RC[-3]", FALSE)>=34,INDIRECT("RC[-3]", FALSE)<67),"Medium",IF(AND(INDIRECT("RC[-3]", FALSE)>=67,INDIRECT("RC[-3]", FALSE)<=100),"High","")))'

for file in csv_files:
    df = pd.read_csv(file)
    df.drop(dropped_columns, axis="columns", inplace=True)
    print(df.info())

    # Clean up "NaN"
    df.fillna('', inplace=True)

    # Add Relative Difficulty column
    df["Relative Difficulty"] = relative_difficulty_formula

    # Take top keyword and use it as the sheet title
    #print(df["Keyword"][0].title())
    title = df["Keyword"][0].title()
    sheet_id = sl.create_sheet_or_get_sheet_id(spreadsheet_id, title, service)

    body = [df.columns.values.tolist()] + df.values.tolist()

    range = title + "!A1:E" + str(len(df)+1)

    sl.update_values(spreadsheet_id, range, "USER_ENTERED", body, service)

    sfl.align_vertical_middle_all(spreadsheet_id, sheet_id, service)
    sfl.bold_center_header_row(spreadsheet_id, sheet_id, service)
    sfl.color_header_row(spreadsheet_id, sheet_id, service, 5)
    sfl.center_columns(spreadsheet_id, sheet_id, service, 1, 5)
    sfl.auto_resize_column(spreadsheet_id, sheet_id, service, 1)
    sfl.auto_resize_column(spreadsheet_id, sheet_id, service, 5)
    sfl.relative_difficulty_conditional_formatting(spreadsheet_id, sheet_id, service, len(df))

