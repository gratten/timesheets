import shutil
import datetime
import os

today = datetime.date.today()
last_monday = today - datetime.timedelta(days=today.weekday())
file_name = f'Ward, Gratten 2021_Timesheet_{last_monday}.xls'
path = os.getcwd()
files = os.listdir(path)

if file_name not in files:
    print('Creating new file...')
    shutil.copy("Ward, Gratten 2021_Timesheet.xls", file_name)
else:
    print("File already exists...")