import time
import datetime
import psutil

from consts import STATUS_FREE, COMMAND, NAME_FILE, FORMAT_TIME2
from work_sheets import row, status_sheet, row_for_stop
from chrome import driver

dt_started = datetime.datetime.utcnow()


def stop_program():
    while True:
        status = status_sheet.acell(f'D{row_for_stop}').value
        if str(status) == COMMAND:
            driver.quit()

            dt_ended = datetime.datetime.utcnow()
            time_end = (dt_ended - dt_started).total_seconds()

            time_finaly = time.localtime()
            current_time = time.strftime(FORMAT_TIME2, time_finaly)

            status_sheet.update(f'B{str(row)}', f'{STATUS_FREE}')
            status_sheet.update(f'H{str(row)}', f'{time_end}')
            status_sheet.update(f'G{str(row)}', f'{current_time}')

            for proc in psutil.process_iter():
                if proc.name() == NAME_FILE:
                    proc.kill()

        time.sleep(10)
