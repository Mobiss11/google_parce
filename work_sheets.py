import socket
from datetime import datetime
import gspread

from credentials import credentials, table, settingsSheet, statusSheet, logsSheet
from consts import STATUS_FREE, STATUS_WORK, FORMAT_TIME3

connect_json = gspread.service_account_from_dict(credentials)

sheet_google = connect_json.open_by_url(table)
status_sheet = sheet_google.worksheet(statusSheet)
logs_sheet = sheet_google.worksheet(logsSheet)

get_names = status_sheet.col_values(1)
nums_robots = len(get_names)
nums_rows = nums_robots + 1


names_robots = []
tables_robot = []
rows = []
emails = []
phones = []

for num_row in range(2, nums_rows):
    get_status = status_sheet.row_values(num_row)
    status = get_status[1]

    if status == STATUS_FREE:
        name_robot = get_status[0]
        name_table = get_status[2]
        email = get_status[9]
        telephone = get_status[10]

        emails.append(email)
        phones.append(telephone)
        tables_robot.append(name_table)
        names_robots.append(name_robot)
        rows.append(num_row)

list_active = tables_robot[0]
row = rows[0]
row_for_stop = row
email_manager = emails[0]
phone_manager = phones[0]

time_now = datetime.now()
name_pk = socket.gethostname()
ip = socket.gethostbyname(socket.gethostname())
current_time = time_now.strftime(FORMAT_TIME3)

status_sheet.update(f'I{str(row_for_stop)}', f'{name_pk}')
status_sheet.update(f'E{str(row_for_stop)}', f'{ip}')
status_sheet.update(f'F{str(row_for_stop)}', f'{current_time}')
status_sheet.update(f'B{str(row)}', f'{STATUS_WORK}')

robots_list = sheet_google.worksheet(list_active)
settings_list_sheet = sheet_google.worksheet(settingsSheet)

names_people = robots_list.col_values(1)
attempts = robots_list.col_values(16)

num_names_people = len(names_people)
num_rows_list = num_names_people + 1

name_robot_active = names_robots[0]
