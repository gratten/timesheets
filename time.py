import shutil
import datetime
import os
# from openpyxl import load_workbook # couldn't get this to work with .xls or vba
import xlwings as xw
import pandas as pd

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

    # sequence logic
    if project[:2] == 'AE':
        seq = '0'
    elif project[:1] == 'J':
        seq = '2'
    else:
        seq = input('Enter sequence: ')

    # activity code logic
    if project[:1] == 'J':
        act_code = '0010'
    elif project[:2] == 'AE':
        print('\n'
              '0010     > CUSTOMER SERVICE\n'
              '0015     > CARTON DESIGN\n'
              '0020     > TECH SALES REPORT\n'
              '0030     > PARTS SUPPORT\n'
              '0040     > HOLIDAY'
              '0042     > VACATION'
              '0045     > SICK DAY\n'
              '0047     > FAMILY ILLNESS\n'
              '0048     > BEREAVEMENT\n'
              '0049     > DR APPOINTMENT\n'
              '0050     > OTHER\n'
              '0055     > DETAIL DRAWINGS'
              '\n')
        act_code = input('Enter activity code: ')
    else:
        act_code = input('Enter activity code: ')

    hours = input("Enter hours: ")

    # populate data
    ws.range(f'A{empty}').value = project
    ws.range(f'B{empty}').value = description
    ws.range(f'C{empty}').value = seq
    ws.range(f'D{empty}').value = act_code

    # determine which day of the week it is determine where to populate hours
    weekday = datetime.datetime.today().weekday()

    if weekday == 0:
        ws.range(f'E{empty}').value = hours
    elif weekday == 1:
        ws.range(f'F{empty}').value = hours
    elif weekday == 2:
        ws.range(f'G{empty}').value = hours
    elif weekday == 3:
        ws.range(f'H{empty}').value = hours
    elif weekday == 4:
        ws.range(f'I{empty}').value = hours
    elif weekday == 5:
        ws.range(f'J{empty}').value = hours
    else:
        ws.range(f'K{empty}').value = hours

    # save and close
    wb.save()
    wb.close()
    excel_app.quit()

    # notify user
    feedback = '\nTask recorded.\n'
    return feedback

def report(file_name):

    # read in data to dataframe
    df = pd.read_excel(file_name,
                       skiprows=19,
                       usecols=('A:K'))

    # manipulate data frame
    df.columns = ['order', 'desc', 'sequence', 'activity', 1, 2, 3, 4, 5, 6, 7]
    df = df.iloc[:df['order'].isnull().values.argmax()]
    df['labor'] = df.iloc[:, 4:11].sum(axis=1)
    df = df.loc[:, df.columns.intersection(['order', 'sequence', 'activity', 'labor'])]
    df = df.groupby(['order', 'sequence', 'activity']).sum().reset_index()

    # write results to new spreadsheet and open
    df.to_excel(f'report_{last_monday}.xlsx')
    wb = xw.Book(f'report_{last_monday}.xlsx')
    last_row = wb.sheets[0].range('E' + str(wb.sheets[0].cells.last_cell.row)).end('up').row + 1
    xw.Range(f'E{last_row}').formula = f'=SUM(E2:E{last_row-1})'

    # notify user
    feedback = '\nReport complete.\n'
    return feedback

def open(file_name):
    xw.Book(file_name)
    feedback = '\nOpening timesheet...\n'
    return feedback

selection = ''
while selection != 'E':
    selection = input('W - new week \n'
                      'T - new task\n'
                      'R - report\n'
                      'O - open timesheet\n'
                      'E - exit\n\n'
                      'Make a selection...')

    if selection == 'W':
        print(new_week(file_name))
    elif selection == 'T':
        print(add_task(file_name))
    elif selection == 'R':
        print(report(file_name))
    elif selection == 'O':
        print(open(file_name))