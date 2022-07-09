from datetime import datetime
import time
import winsound
from threading import Thread
import random

from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from consts import *
from work_sheets import robots_list, settings_list_sheet, num_rows_list, logs_sheet, name_robot_active
from gui import display
from chrome import driver


def main_bot():

    while True:

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
            telefone = values_list[13]
            email = values_list[14]
            numberOfTries = values_list[15]

            time_z = int(settings_list[1])
            time_min = int(settings_list[3])
            time_max = int(settings_list[4])

            if usluga == AsignacionNIE:

                driver.get(MAIN_URL)
                driver.implicitly_wait(80)

                select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                select_region = Select(select_input_region)
                select_region.select_by_visible_text(region)
                time.sleep(5)

                driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                driver.implicitly_wait(10)

                time.sleep(8)

                select_element2 = driver.find_element(By.XPATH,
                                                      '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form[1]/div[3]/div[1]/div[2]/div/fieldset/div[2]/select')

                select_object2 = Select(select_element2)
                select_object2.select_by_visible_text(str(usluga))
                time.sleep(5)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form[1]/div[4]/input[1]').click()
                time.sleep(5)
                driver.implicitly_wait(80)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div/div[3]/input[1]').submit()
                driver.implicitly_wait(80)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div/div/div[1]/div[2]/div/div/div[2]/input').send_keys(
                    pasport)
                time.sleep(5)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div/div/div[1]/div[3]/div/div/div/div/input[1]') \
                    .send_keys(f'{name} {famila}')
                time.sleep(5)
                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div/div/div[1]/div[4]/div/div/div/div/input[1]').send_keys(
                    date)
                time.sleep(5)

                select_element3 = driver.find_element(By.XPATH,
                                                      '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div/div/div[1]/div[5]/div/div/div/div/span/select')
                select_object3 = Select(select_element3)
                select_object3.select_by_visible_text(country)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()
                driver.implicitly_wait(5)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                # выбор офиса
                try:
                    driver.find_element(By.XPATH,
                                        '/html/ody/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/label')
                    select_element4 = driver.find_element(By.XPATH,
                                                          '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/select')
                    select_object4 = Select(select_element4)
                    try:
                        select_object4.select_by_index(1)
                    except:
                        select_object4.select_by_index(0)

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[2]/input[1]').click()

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[1]/input[2]').send_keys(
                        telefone)
                    time.sleep(2)

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[1]/input[2]').send_keys(
                        email)
                    time.sleep(2)

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[2]/input').send_keys(
                        email)
                    time.sleep(2)
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div[2]/input[1]').click()

                    try:
                        driver.find_element(By.ID, 'cita_1')
                        winsound.Beep(FREQUENCY, DURATION)
                        time.sleep(1)
                        winsound.Beep(FREQUENCY, DURATION)
                        time.sleep(900)
                    except:
                        print("ситов не найдено")
                except:
                    print('офисов не найдено')
                    time.sleep(time_z)
            elif usluga == TomaHuellas:

                attempt = 1
                while attempt <= int(numberOfTries):

                    logs_col = logs_sheet.col_values(1)
                    last_element = logs_col[-1]
                    index_last_element = logs_col.index(last_element)
                    number_row = index_last_element + 2

                    time_now = datetime.now()
                    current_time = time_now.strftime("%d-%m-%Y %H:%M")

                    logs_sheet.update(f'A{str(number_row)}', f'{current_time}')

                    driver.get(MAIN_URL)
                    driver.implicitly_wait(80)

                    select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                    select_region = Select(select_input_region)
                    select_region.select_by_visible_text(region)
                    time.sleep(2)

                    driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                    driver.implicitly_wait(10)

                    select_element2 = driver.find_element(By.ID, 'tramiteGrupo[1]')
                    select_object2 = Select(select_element2)
                    select_object2.select_by_visible_text(str(usluga))
                    time.sleep(random.randint(time_min, time_max))

                    driver.find_element(By.ID, 'btnAceptar').click()
                    driver.implicitly_wait(80)
                    time.sleep(random.randint(time_min, time_max))

                    driver.find_element(By.ID, 'btnEntrar').submit()
                    driver.implicitly_wait(80)

                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[2]/div/div/div[2]/input').send_keys(
                        nie)
                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[3]/div/div/div/div/input[1]') \
                        .send_keys(f'{name} {famila}')
                    select_element3 = driver.find_element(By.XPATH,
                                                          '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[4]/div/div/div/div/span/select')
                    select_object3 = Select(select_element3)
                    select_object3.select_by_visible_text(country)

                    time.sleep(random.randint(time_min, time_max))

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()
                    driver.implicitly_wait(5)
                    time.sleep(random.randint(time_min, time_max))
                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()

                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                    # выбор офиса
                    try:
                        driver.find_element(By.XPATH,
                                            '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/label')
                        select_element4 = driver.find_element(By.XPATH,
                                                              '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/select')
                        select_object4 = Select(select_element4)
                        try:
                            select_object4.select_by_index(1)
                        except:
                            select_object4.select_by_index(0)

                        driver.find_element(By.XPATH,
                                            '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[2]/input[1]').click()

                        winsound.Beep(FREQUENCY, DURATION)
                        time.sleep(1)

                        driver.implicitly_wait(80)

                        driver.find_element(By.XPATH,
                                            '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[1]/input[2]').send_keys(
                            telefone)

                        driver.find_element(By.XPATH,
                                            '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[1]/input[2]').send_keys(
                            email)

                        driver.find_element(By.XPATH,
                                            '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[2]/input').send_keys(
                            email)

                        driver.find_element(By.ID, 'btnSiguiente').click()
                        driver.implicitly_wait(80)

                        try:
                            driver.find_element(By.ID, 'cita_1')

                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime("%d-%m-%Y %H:%M")

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_3}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            winsound.Beep(FREQUENCY, DURATION)
                            time.sleep(1)
                            time.sleep(900)
                        except:
                            time_finaly = datetime.now()
                            current_time = time_finaly.strftime("%d-%m-%Y %H:%M")

                            logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                            logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_1}')
                            logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                            logs_sheet.update(f'E{str(number_row)}', f'{name}')
                            logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                            logs_sheet.update(f'G{str(number_row)}', f'{usluga}')


                    except:
                        time_finaly = datetime.now()
                        current_time = time_finaly.strftime("%d-%m-%Y %H:%M")

                        logs_sheet.update(f'B{str(number_row)}', f'{current_time}')
                        logs_sheet.update(f'C{str(number_row)}', f'{TEXT_LOGS_2}')
                        logs_sheet.update(f'D{str(number_row)}', f'{name_robot_active}')
                        logs_sheet.update(f'E{str(number_row)}', f'{name}')
                        logs_sheet.update(f'F{str(number_row)}', f'{famila}')
                        logs_sheet.update(f'G{str(number_row)}', f'{usluga}')

                        time.sleep(time_z)

                    attempt += 1
            elif usluga == CertificadoUE:

                driver.get(MAIN_URL)
                driver.implicitly_wait(80)

                select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                select_region = Select(select_input_region)
                select_region.select_by_visible_text(region)
                time.sleep(5)

                driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                driver.implicitly_wait(10)

                time.sleep(8)

                select_element2 = driver.find_element(By.XPATH,
                                                      '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form[1]/div[3]/div[1]/div[2]/div/fieldset/div[2]/select')
                select_object2 = Select(select_element2)
                select_object2.select_by_visible_text(str(usluga))
                time.sleep(3)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form[1]/div[4]/input[1]').click()
                driver.implicitly_wait(80)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div[3]/input[1]').click()
                driver.implicitly_wait(80)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[2]/div/div/div[2]/input').send_keys(
                    nie)
                time.sleep(2)
                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[3]/div/div/div/div/input[1]') \
                    .send_keys(f'{name} {famila}')
                time.sleep(2)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()
                driver.implicitly_wait(5)
                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                # выбор офиса
                try:
                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/label')
                    select_element4 = driver.find_element(By.XPATH,
                                                          '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/select')
                    select_object4 = Select(select_element4)
                    try:
                        select_object4.select_by_index(1)
                    except:
                        select_object4.select_by_index(0)

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[2]/input[1]').click()

                    winsound.Beep(FREQUENCY, DURATION)
                    time.sleep(1)
                    winsound.Beep(FREQUENCY, DURATION)
                    time.sleep(900)

                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[1]/input[2]').send_keys(
                #         telefone)
                #     time.sleep(2)
                #
                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[1]/input[2]').send_keys(
                #         email)
                #     time.sleep(2)
                #
                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[2]/input').send_keys(
                #         email)
                #     time.sleep(2)
                #
                #     while i < 3:
                #         # Scroll down to bottom
                #         driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                #
                #         # Wait to load page
                #         time.sleep(SCROLL_PAUSE_TIME)
                #
                #         # Calculate new scroll height and compare with last scroll height
                #         new_height = driver.execute_script("return document.body.scrollHeight")
                #         if new_height == last_height:
                #             break
                #         last_height = new_height
                #
                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div[2]/input[1]').click()
                #
                #     try:
                #         driver.find_element(By.ID, 'cita_1')
                #         winsound.Beep(frequency, duration)
                #         time.sleep(1)
                #         winsound.Beep(frequency, duration)
                #         time.sleep(900)
                #     except:
                #         print("ситов не найдено")

                except:
                    print('офисов не найдено')
                    time.sleep(time_z)
            elif usluga == TarjetaUkranea:

                driver.get(MAIN_URL)
                driver.implicitly_wait(80)

                select_input_region = driver.find_element(By.ID, MAIN_SELECT)
                select_region = Select(select_input_region)
                select_region.select_by_visible_text(region)
                time.sleep(5)

                driver.find_element(By.ID, MAIN_PATH_BUTTON).click()
                driver.implicitly_wait(10)

                time.sleep(8)

                select_element2 = driver.find_element(By.XPATH,
                                                      '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form[1]/div[3]/div[1]/div[2]/div/fieldset/div[2]/select')
                select_object2 = Select(select_element2)
                select_object2.select_by_visible_text(str(usluga))
                time.sleep(3)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form[1]/div[4]/input[1]').click()
                driver.implicitly_wait(80)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div/div[2]/input[1]').click()
                driver.implicitly_wait(80)

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[2]/div/div/div[2]/input').send_keys(
                    nie)
                time.sleep(2)
                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[3]/div/div/div/div/input[1]') \
                    .send_keys(f'{name} {famila}')
                time.sleep(2)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()
                driver.implicitly_wait(5)
                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()

                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                # выбор офиса
                try:
                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/label')
                    select_element4 = driver.find_element(By.XPATH,
                                                          '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/fieldset/div/select')
                    select_object4 = Select(select_element4)
                    try:
                        select_object4.select_by_index(1)
                    except:
                        select_object4.select_by_index(0)

                    driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[2]/input[1]').click()

                    winsound.Beep(FREQUENCY, DURATION)
                    time.sleep(1)
                    winsound.Beep(FREQUENCY, DURATION)
                    time.sleep(900)

                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[1]/input[2]').send_keys(
                #         telefone)
                #     time.sleep(2)
                #
                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[1]/input[2]').send_keys(
                #         email)
                #     time.sleep(2)
                #
                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[2]/input').send_keys(
                #         email)
                #     time.sleep(2)
                #
                #     while i < 3:
                #         # Scroll down to bottom
                #         driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                #
                #         # Wait to load page
                #         time.sleep(SCROLL_PAUSE_TIME)
                #
                #         # Calculate new scroll height and compare with last scroll height
                #         new_height = driver.execute_script("return document.body.scrollHeight")
                #         if new_height == last_height:
                #             break
                #         last_height = new_height
                #
                #     driver.find_element(By.XPATH,
                #                         '/html/body/div[1]/div[2]/main/div/div/section/div[2]/form/div[2]/input[1]').click()
                #
                #     try:
                #         driver.find_element(By.ID, 'cita_1')
                #         winsound.Beep(frequency, duration)
                #         time.sleep(1)
                #         winsound.Beep(frequency, duration)
                #         time.sleep(900)
                #     except:
                #         print("ситов не найдено")

                except:
                    print('офисов не найдено')
                    time.sleep(time_z)


if __name__ == '__main__':
    Thread(target = display).start()
    Thread(target = main_bot).start()






