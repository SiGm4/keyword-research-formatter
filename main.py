import numpy as np
import pandas as pd

droppedColumns = ["#","Country","CPC","CPS","Parent Keyword","Last Update","SERP Features", "Traffic potential"]

df = pd.read_csv('test.csv')

df.drop(droppedColumns, axis="columns", inplace=True)

print(df.info())
