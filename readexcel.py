from openpyxl import load_workbook


def get_values_from_excel(excel_file):
    wb = load_workbook(filename = excel_file)
    sheet_ranges = wb['Daily Update']
    id = str(sheet_ranges['I11'].value)
    password = str(sheet_ranges['I12'].value)
    hours_column = 'A'
    project_column = 'B'
    task_column = 'C'
    notes_column = 'D'
    current_row = 2
    hour_time = str(sheet_ranges[hours_column+str(current_row)].value)
    while hour_time != 'None':
        hour_time = str(sheet_ranges[hours_column+str(current_row)].value)
        project = str(sheet_ranges[project_column+str(current_row)].value)
        task = str(sheet_ranges[task_column+str(current_row)].value)
        notes = str(sheet_ranges[notes_column+str(current_row)].value)
        if hour_time == 'None':
            break
        else:
            current_row += 1
            yield id, password, hour_time, project, task, notes