import time
import datetime
import psutil
import tkinter as tk
import sys

from consts import *
from work_sheets import row, status_sheet, names_people, attempts
from chrome import driver

dt_started = datetime.datetime.utcnow()


def display():

    root = tk.Tk()
    root.title(TITLE)
    root.attributes(ATTRIBUTE, True)
    root.geometry(ROOT_CONSTANT)

    all_names = len(names_people) + 1
    all_attempts = len(attempts) + 1

    for num_name in range(1, all_names):
        tk.Label(root,
                 text=names_people[0],
                 font=FONT).grid(row=num_name,
                                 column=1,
                                 padx=0,
                                 pady=0,
                                 sticky='w')

        names_people.remove(names_people[0])

    for num_attempt in range(1, all_attempts):
        tk.Label(root,
                 text=attempts[0],
                 font=FONT).grid(row=num_attempt,
                                 column=2,
                                 padx=0,
                                 pady=0,
                                 sticky='w')

        attempts.remove(attempts[0])

    def stop():

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

        sys.exit()

    tk.Button(root,
              text=BUTTON_TEXT_STOP,
              command=stop,
              width=11,
              fg=FONT_COLOR,
              bg=BG_COLOR_STOP,
              height=1).grid(row=1,
                             column=4,
                             padx=4,
                             pady=4)

    root.mainloop()