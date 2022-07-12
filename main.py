from datetime import datetime
import time
import winsound
from threading import Thread
import random
import socket

from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from consts import *
from work_sheets import robots_list, settings_list_sheet, num_rows_list, logs_sheet, name_robot_active, status_sheet\
    , row_for_stop, email_manager, phone_manager
from gui import stop_program
from chrome import driver


def main_bot():

    while True:
        try:
            for row in range(2, num_rows_list):

                values_list = robots_list.row_values(row)
                settings_list = settings_list_sheet.col_values(2)

                name = values_list[0]
                famila = values_list[1]
                usluga = values_list[2]
                region = values_list[5]
                pasport = values_list[6]
                nie = values_list[7]
                date = values_list[8]
                country = values_list[9]
                telefone = phone_manager
                email = email_manager
                number_of_tries = values_list[15]

                time_z = int(settings_list[1])
                time_min = int(settings_list[3])
                time_max = int(settings_list[4])

                name_pk = socket.gethostname()
                ip = socket.gethostbyname(socket.gethostname())

                if usluga == AsignacionNIE:

                    attempt = 1
                    while attempt <= int(number_of_tries):

                        time_now = datetime.now()
                        current_time = time_now.strftime(FORMAT_TIME3)

                        status_sheet.update(f'B{str(row)}', f'{STATUS_WORK} {name} {famila}')
                        status_sheet.update(f'I{str(row)}', f'{name_pk}')
                        status_sheet.update(f'E{str(row)}', f'{ip}')
                        status_sheet.update(f'F{str(row)}', f'{current_time}')

                        status_sheet.update(f'B{str(row)}', f'{STATUS_WORK} {name} {famila}')

                        logs_col = logs_sheet.col_values(1)
                        last_element = logs_col[-1]
                        index_last_element = logs_col.index(last_element)
                        number_row = index_last_element + 2

                        logs_sheet.update(f'A{str(number_row)}', f'{time_now}')

                        driver.get(MAIN_URL)
                        driver.implicitly_wait(80)

                        select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                        select_region = Select(select_input_region)
                        select_region.select_by_visible_text(region)
                        time.sleep(2)

                        driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                        driver.implicitly_wait(10)

                        select_input_usluga = driver.find_element(By.ID, ID_INPUT_USLUGA)
                        select_usluga = Select(select_input_usluga)
                        select_usluga.select_by_visible_text(str(usluga))

                        sec_first = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_first}')
                        time.sleep(sec_first)

                        driver.find_element(By.ID, ID_BUTTON_ACEPTAR).click()
                        driver.implicitly_wait(80)

                        sec_second = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_second}')
                        time.sleep(sec_second)

                        driver.find_element(By.ID, ID_BUTTON_ENTRAR).submit()
                        driver.implicitly_wait(80)

                        sec_third = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WRITE} {name} {famila}')

                        driver.find_element(By.ID, ID_INPUT_PASPORT_NIE).send_keys(pasport)

                        driver.find_element(By.ID, ID_INPUT_NOMBRE_NIE).send_keys(f'{name} {famila}')

                        driver.find_element(By.ID, ID_INPUT_DATE_NIE).send_keys(date)

                        select_element3 = driver.find_element(By.ID, ID_INPUT_COUNTRY_NIE)
                        select_object3 = Select(select_element3)
                        select_object3.select_by_visible_text(country)

                        time.sleep(sec_third)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        sec_fourth = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_fourth}')
                        time.sleep(sec_fourth)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                        # выбор офиса
                        try:
                            driver.find_element(By.XPATH, LABEL_OFFICE)

                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                            select_element4 = driver.find_element(By.XPATH, ID_INPUT_OFFICE)
                            select_object4 = Select(select_element4)
                            try:
                                select_object4.select_by_index(1)
                            except:
                                select_object4.select_by_index(0)

                            driver.find_element(By.XPATH, ID_BUTTON_OFFICE).click()

                            driver.find_element(By.XPATH, ID_INPUT_TELEFONE).send_keys(telefone)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL).send_keys(email)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL_DOUBLE).send_keys(email)

                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_3}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            sec_fivth = 900
                            status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING_USER} {sec_fivth}')

                            time.sleep(sec_fivth)

                        except:
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_2}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            status_sheet.update(f'B{str(row)}', f'{STATUS_NO_OFFICE} {time_z}')

                            time.sleep(time_z)

                        attempt += 1
                elif usluga == TomaHuellas:

                    attempt = 1
                    while attempt <= int(number_of_tries):

                        time_now = datetime.now()
                        current_time = time_now.strftime(FORMAT_TIME3)

                        status_sheet.update(f'B{str(row)}', f'{STATUS_WORK} {name} {famila}')
                        status_sheet.update(f'I{str(row)}', f'{name_pk}')
                        status_sheet.update(f'E{str(row)}', f'{ip}')
                        status_sheet.update(f'F{str(row)}', f'{current_time}')

                        logs_col = logs_sheet.col_values(1)
                        last_element = logs_col[-1]
                        index_last_element = logs_col.index(last_element)
                        number_row = index_last_element + 2

                        logs_sheet.update(f'A{str(number_row)}', f'{time_now}')

                        driver.get(MAIN_URL)
                        driver.implicitly_wait(80)

                        select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                        select_region = Select(select_input_region)
                        select_region.select_by_visible_text(region)
                        time.sleep(2)

                        driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                        driver.implicitly_wait(10)

                        select_input_usluga = driver.find_element(By.ID, ID_INPUT_USLUGA)
                        select_usluga = Select(select_input_usluga)
                        select_usluga.select_by_visible_text(str(usluga))

                        sec_first = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_first}')
                        time.sleep(sec_first)

                        driver.find_element(By.ID, ID_BUTTON_ACEPTAR).click()
                        driver.implicitly_wait(80)

                        sec_second = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_second}')
                        time.sleep(sec_second)

                        driver.find_element(By.ID, ID_BUTTON_ENTRAR).submit()
                        driver.implicitly_wait(80)

                        sec_third = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WRITE} {name} {famila}')

                        driver.find_element(By.ID, ID_INPUT_NIE_TOMA).send_keys(nie)
                        driver.implicitly_wait(80)

                        driver.find_element(By.ID, ID_INPUT_NOMBRE_FAMILAR_TOMA).send_keys(f'{name} {famila}')
                        driver.implicitly_wait(80)

                        select_input_country = driver.find_element(By.ID, ID_INPUT_COUNTRY_TOMA)
                        select_country = Select(select_input_country)
                        select_country.select_by_visible_text(country)

                        time.sleep(sec_third)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        sec_fourth = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_fourth}')
                        time.sleep(sec_fourth)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                        # выбор офиса
                        try:
                            driver.find_element(By.XPATH, LABEL_OFFICE)

                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                            select_element4 = driver.find_element(By.XPATH, ID_INPUT_OFFICE)
                            select_object4 = Select(select_element4)
                            try:
                                select_object4.select_by_index(1)
                            except:
                                select_object4.select_by_index(0)

                            driver.find_element(By.XPATH, ID_BUTTON_OFFICE).click()

                            driver.find_element(By.XPATH, ID_INPUT_TELEFONE).send_keys(telefone)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL).send_keys(email)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL_DOUBLE).send_keys(email)

                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_3}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            sec_fivth = 900
                            status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING_USER} {sec_fivth}')

                            time.sleep(sec_fivth)

                        except:
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_2}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            status_sheet.update(f'B{str(row)}', f'{STATUS_NO_OFFICE} {time_z}')

                            time.sleep(time_z)

                        attempt += 1
                elif usluga == CertificadoUE:

                    attempt = 1
                    while attempt <= int(number_of_tries):

                        time_now = datetime.now()
                        current_time = time_now.strftime(FORMAT_TIME3)

                        status_sheet.update(f'B{str(row)}', f'{STATUS_WORK} {name} {famila}')
                        status_sheet.update(f'I{str(row)}', f'{name_pk}')
                        status_sheet.update(f'E{str(row)}', f'{ip}')
                        status_sheet.update(f'F{str(row)}', f'{current_time}')

                        status_sheet.update(f'B{str(row)}', f'{STATUS_WORK} {name} {famila}')

                        logs_col = logs_sheet.col_values(1)
                        last_element = logs_col[-1]
                        index_last_element = logs_col.index(last_element)
                        number_row = index_last_element + 2

                        logs_sheet.update(f'A{str(number_row)}', f'{time_now}')

                        driver.get(MAIN_URL)
                        driver.implicitly_wait(80)

                        select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                        select_region = Select(select_input_region)
                        select_region.select_by_visible_text(region)
                        time.sleep(2)

                        driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                        driver.implicitly_wait(10)

                        select_input_usluga = driver.find_element(By.ID, ID_INPUT_USLUGA)
                        select_usluga = Select(select_input_usluga)
                        select_usluga.select_by_visible_text(str(usluga))

                        sec_first = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_first}')
                        time.sleep(sec_first)

                        driver.find_element(By.ID, ID_BUTTON_ACEPTAR).click()
                        driver.implicitly_wait(80)

                        sec_second = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_second}')
                        time.sleep(sec_second)

                        driver.find_element(By.ID, ID_BUTTON_ENTRAR).submit()
                        driver.implicitly_wait(80)

                        sec_third = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WRITE} {name} {famila}')

                        driver.find_element(By.ID, ID_INPUT_NIE_CERT).send_keys(nie)

                        driver.find_element(By.ID, ID_INPUT_NOMBRE_CERT).send_keys(f'{name} {famila}')

                        time.sleep(sec_third)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        sec_fourth = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_fourth}')
                        time.sleep(sec_fourth)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                        # выбор офиса
                        try:
                            driver.find_element(By.XPATH, LABEL_OFFICE)

                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                            select_element4 = driver.find_element(By.XPATH, ID_INPUT_OFFICE)
                            select_object4 = Select(select_element4)
                            try:
                                select_object4.select_by_index(1)
                            except:
                                select_object4.select_by_index(0)

                            driver.find_element(By.XPATH, ID_BUTTON_OFFICE).click()

                            driver.find_element(By.XPATH, ID_INPUT_TELEFONE).send_keys(telefone)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL).send_keys(email)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL_DOUBLE).send_keys(email)

                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_3}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            sec_fivth = 900
                            status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING_USER} {sec_fivth}')

                            time.sleep(sec_fivth)

                        except:
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_2}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            status_sheet.update(f'B{str(row)}', f'{STATUS_NO_OFFICE} {time_z}')

                            time.sleep(time_z)

                        attempt += 1
                elif usluga == TarjetaUkranea:

                    attempt = 1
                    while attempt <= int(number_of_tries):

                        time_now = datetime.now()
                        current_time = time_now.strftime(FORMAT_TIME3)

                        status_sheet.update(f'B{str(row)}', f'{STATUS_WORK} {name} {famila}')
                        status_sheet.update(f'I{str(row)}', f'{name_pk}')
                        status_sheet.update(f'E{str(row)}', f'{ip}')
                        status_sheet.update(f'F{str(row)}', f'{current_time}')

                        status_sheet.update(f'B{str(row)}', f'{STATUS_WORK} {name} {famila}')

                        logs_col = logs_sheet.col_values(1)
                        last_element = logs_col[-1]
                        index_last_element = logs_col.index(last_element)
                        number_row = index_last_element + 2

                        logs_sheet.update(f'A{str(number_row)}', f'{time_now}')

                        driver.get(MAIN_URL)
                        driver.implicitly_wait(80)

                        select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                        select_region = Select(select_input_region)
                        select_region.select_by_visible_text(region)
                        time.sleep(2)

                        driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                        driver.implicitly_wait(10)

                        select_input_usluga = driver.find_element(By.ID, ID_INPUT_USLUGA)
                        select_usluga = Select(select_input_usluga)
                        select_usluga.select_by_visible_text(str(usluga))

                        sec_first = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_first}')
                        time.sleep(sec_first)

                        driver.find_element(By.ID, ID_BUTTON_ACEPTAR).click()
                        driver.implicitly_wait(80)

                        sec_second = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_second}')
                        time.sleep(sec_second)

                        driver.find_element(By.ID, ID_BUTTON_ENTRAR).submit()
                        driver.implicitly_wait(80)

                        sec_third = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WRITE} {name} {famila}')

                        driver.find_element(By.ID, ID_INPUT_NIE_CERT).send_keys(nie)

                        driver.find_element(By.ID, ID_INPUT_NOMBRE_CERT).send_keys(f'{name} {famila}')

                        time.sleep(sec_third)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        sec_fourth = random.randint(time_min, time_max)
                        status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING} {sec_fourth}')
                        time.sleep(sec_fourth)

                        driver.find_element(By.ID, ID_BUTTON_ENVIAR).click()
                        driver.implicitly_wait(5)

                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                        # выбор офиса
                        try:
                            driver.find_element(By.XPATH, LABEL_OFFICE)

                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                            select_element4 = driver.find_element(By.XPATH, ID_INPUT_OFFICE)
                            select_object4 = Select(select_element4)
                            try:
                                select_object4.select_by_index(1)
                            except:
                                select_object4.select_by_index(0)

                            driver.find_element(By.XPATH, ID_BUTTON_OFFICE).click()

                            driver.find_element(By.XPATH, ID_INPUT_TELEFONE).send_keys(telefone)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL).send_keys(email)

                            driver.find_element(By.XPATH, ID_INPUT_EMAIL_DOUBLE).send_keys(email)

                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_3}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            sec_fivth = 900
                            status_sheet.update(f'B{str(row)}', f'{STATUS_WAITING_USER} {sec_fivth}')

                            time.sleep(sec_fivth)

                        except:
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime(FORMAT_TIME)

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_2}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            status_sheet.update(f'B{str(row)}', f'{STATUS_NO_OFFICE} {time_z}')

                            time.sleep(time_z)

                        attempt += 1

        except Exception as problem:

            logs_col = logs_sheet.col_values(1)
            last_element = logs_col[-1]
            index_last_element = logs_col.index(last_element)
            number_row = index_last_element + 2

            time_now = datetime.now()

            logs_sheet.update(f'A{str(number_row)}', f'{time_now}')

            time_finaly = datetime.now()
            current_time = time_finaly.strftime(FORMAT_TIME)

            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
            logs_sheet.update(f'C{str(number_row)}', f'-')
            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
            logs_sheet.update(f'E{str(number_row)}', f'-')
            logs_sheet.update(f'F{str(number_row)}', f'-')
            logs_sheet.update(f'G{str(number_row)}', f'-')
            logs_sheet.update(f'H{str(number_row)}', f'{problem}')

            sec_stop = 300
            status_sheet.update(f'B{str(row_for_stop)}', f'{STATUS_TOO_MANY_REQUEST} {sec_stop}')

            time.sleep(sec_stop)


if __name__ == '__main__':
    Thread(target=stop_program).start()
    Thread(target=main_bot).start()






