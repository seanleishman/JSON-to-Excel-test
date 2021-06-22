import pandas as pd
import json

with open('test.json') as json_file:  # open test json file
    json_data = json.load(json_file)

df = pd.DataFrame(json_data)  # converts JSON file to data frame

severity_filter = df.loc[df['severity'] == 'ok']  # apply filter where 'severity' key is = 'ok'

severity_filter.to_excel('Severity filter spreadsheet.xlsx', index=False)  # converts data frame to excel file
