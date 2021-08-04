import shutil
import datetime
import os
# from openpyxl import load_workbook
import xlwings as xw

# determine filename based on today's date
today = datetime.date.today()
last_monday = today - datetime.timedelta(days=today.weekday())
file_name = f'Ward, Gratten 2021_Timesheet_{last_monday}.xls'
path = os.getcwd()
files = os.listdir(path)

def new_week(file_name):
    if file_name not in files:
        shutil.copy("Ward, Gratten 2021_Timesheet.xls", file_name)
        feedback = '\nNew file created.\n'
    else:
        feedback = "\nFile already exists.\n"
    return feedback

def add_task(file_name):

    # initiate workbook
    excel_app = xw.App(visible=False)
    wb = excel_app.books.open(file_name)
    ws = wb.sheets[0]

    # find next row to populate
    cell_range = ws.range('A21', 'A36')
    for cell in cell_range:
        if cell[0].value is None:
            empty = cell[0].row
            break
    else:
        print('Document full.')

    # collect user input
    project = input("Enter project: ")
    description = input("Enter description: ")
    seq = input("Enter sequence: ")
    act_code = input("Enter activity code: ")
    hours = input("Enter hours: ")

    # populate data
    ws.range(f'A{empty}').value = project
    ws.range(f'B{empty}').value = description
    ws.range(f'C{empty}').value = seq
    ws.range(f'D{empty}').value = act_code
    ws.range(f'E{empty}').value = hours

    # save and close
    wb.save()
    wb.close()
    excel_app.quit()

    # determine which day to populate hours

    # notify user
    feedback = '\nTask recorded.\n'
    return feedback

selection = ''
while selection != 'E':
    selection = input('W - new week \n'
                      'T - add task\n'
                      'E - exit\n\n'
                      'Make a selection...')

    if selection == 'W':
        print(new_week(file_name))
    elif selection == 'T':
        print(add_task(file_name))