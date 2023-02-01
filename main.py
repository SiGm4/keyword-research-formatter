import numpy as np
import pandas as pd

import sheetslib as sl

from googleapiclient.errors import HttpError

droppedColumns = ["#","Country","CPC","CPS","Parent Keyword","Last Update","SERP Features", "Traffic potential"]

df = pd.read_csv('test.csv')

df.drop(droppedColumns, axis="columns", inplace=True)

print(df.info())

# The ID of the main spreadsheet
spreadsheet_id = '1vwazr-XfwBhWsA7ArqoQNpOP-hg0BGUluYlMGwFvEwc'
SAMPLE_RANGE_NAME = 'A1:B3'

if __name__ == '__main__':
    creds, service = sl.authorize()

print(creds, service)

sheet_id = sl.create_sheet_or_get_sheet_id(spreadsheet_id, "Meta Description Issues", service)