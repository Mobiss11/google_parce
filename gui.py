import sys
import time

from consts import STATUS_FREE, COMMAND
from work_sheets import row, status_sheet, row_for_stop
from chrome import driver


def stop_program():
    while True:
        status = status_sheet.acell(f'D{row_for_stop}').value
        if status == COMMAND:
            status_sheet.update(f'B{str(row)}', f'{STATUS_FREE}')
            driver.quit()
            sys.exit(0)
        time.sleep(10)
