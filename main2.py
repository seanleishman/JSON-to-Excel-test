from pprint import pprint
import os
import json
import win32com.client as win32

json_data = json.loads(open('test.json').read())

with open('test.json') as JsonFile:
    jsondata = json.load(JsonFile)

    for x in jsondata:
        header = x.keys()
        #pprint(header)

rows = []

try:
    for record in jsondata:
        result_type = record['result_type']
        external_title = record['external_title']
        title = record['title']
        external_text = record['external_text']
        text = record['text']
        #severity = record['severity']
        authors = ','.join(record['authors'])
        name = record['name']
        labels = ','.join(record['labels'])
        ic_type = record['ic_type']
        module_date_modified = record['module_date_modified']
        alert_id = record['alert_id']
        issue_id = record['issue_id']
        problem_id = record['problem_id']

        rows.append([result_type, external_title, title, external_text, text, #severity,
                     authors, name, labels, ic_type, module_date_modified, alert_id, issue_id, problem_id])

        pprint(rows)

except KeyError:
    pass
