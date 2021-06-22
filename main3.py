from pprint import pprint
import os
import json
import win32com.client as win32

json_data = json.loads(open('test.json').read())

rows = []

try:
    for record in json_data:
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

        ExcelApp = win32.Dispatch('Excel.Application')
        ExcelApp.Visible = True

        wb = ExcelApp.Workbooks.Add()  # creates excel workbook
        ws = wb.Worksheets(1)

        header_labels = (
            'result_type', 'external_title', 'title', 'external_text', 'text', 'severity', 'authors', 'name', 'labels',
            'ic_type', 'module_date_modified', 'alert_id', 'issue_id', 'problem_id')

        for index, val in enumerate(header_labels):  # inserts headers
            ws.Cells(1, index + 1).Value = val

        row_tracker = 2
        column_size = len(header_labels)

        for row in rows:
            ws.Range(
                ws.Cells(row_tracker, 1),
                ws.Cells(row_tracker, column_size)
            ).value = row
            row_tracker += 1

        wb.SaveAs(os.path.join(os.getcwd(), 'Severity OK spreadsheet.xlsx'), 51)
        wb.Close()
        ExcelApp.Quit()
        ExcelApp = None


except KeyError:
    pass
