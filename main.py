import glob
import pandas as pd

filepaths = glob.glob('files/*.xlsx')

for filepath in filepaths:
    data = pd.read_excel(filepath, sheet_name='Sheet1')
    print(data)
    
