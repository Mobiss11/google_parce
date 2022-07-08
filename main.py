import time
import winsound

import gspread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from credentials import credentials, table, sheet
from consts import *

connect_json = gspread.service_account_from_dict(credentials)

sheet_google = connect_json.open_by_url(table)
list_sheet = sheet_google.worksheet(sheet)

options = webdriver.ChromeOptions()
#options.add_argument(OPTION_1)
options.add_argument(OPTION_2)
options.add_argument(OPTION_3)
options.add_argument(OPTION_4)

driver = webdriver.Chrome(
    executable_path=NAME_DRIVER,
    options=options
)

names = values_list = list_sheet.col_values(2)
num_names = len(names)
num_names2 = num_names + 1

while True:

    for row in range(2, num_names2):

        values_list = list_sheet.row_values(row)

        name = values_list[1]
        famila = values_list[2]
        usluga = values_list[3]
        region = values_list[6]
        pasport = values_list[7]
        nie = values_list[8]
        date = values_list[9]
        country = values_list[10]
        telefone = values_list[14]
        email = values_list[15]
        time_z = int(list_sheet.acell(f'Q2').value)

        if usluga == MAIN_USL_1:

            driver.get(MAIN_URL)
            driver.implicitly_wait(80)

            select_element1 = driver.find_element(By.ID,'form')
            select_object1 = Select(select_element1)
            select_object1.select_by_visible_text(region)

            driver.find_element(By.ID, 'btnAceptar').click()
            driver.implicitly_wait(80)

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

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
        elif usluga == MAIN_USL_2:

            driver.get(MAIN_URL)
            driver.implicitly_wait(80)
            time.sleep(5)

            select_element1 = driver.find_element(By.ID, 'form')
            select_object1 = Select(select_element1)
            select_object1.select_by_visible_text(region)

            time.sleep(12)

            driver.find_element(By.ID, 'btnAceptar').click()
            driver.implicitly_wait(10)

            time.sleep(6)

            #driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            select_element2 = driver.find_element(By.ID, 'tramiteGrupo[1]')
            select_object2 = Select(select_element2)
            select_object2.select_by_index(9)
            time.sleep(5)

            driver.find_element(By.ID, 'btnAceptar').click()
            driver.implicitly_wait(80)
            time.sleep(12)

            #driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            driver.find_element(By.ID, 'btnEntrar').submit()
            driver.implicitly_wait(80)

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[2]/div/div/div[2]/input').send_keys(
                nie)
            time.sleep(6)
            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[3]/div/div/div/div/input[1]') \
                .send_keys(f'{name} {famila}')
            time.sleep(7)
            select_element3 = driver.find_element(By.XPATH,
                                                  '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[1]/div[4]/div/div/div/div/span/select')
            select_object3 = Select(select_element3)
            select_object3.select_by_visible_text(country)

            time.sleep(10)

            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div/main/div/div/section/div[2]/form/div/div/div[2]/input[1]').click()
            driver.implicitly_wait(5)
            time.sleep(4)
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
                time.sleep(3)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[1]/input[2]').send_keys(
                    email)
                time.sleep(4)

                driver.find_element(By.XPATH,
                                    '/html/body/div[1]/div/main/div/div/section/div[2]/form/div[1]/div/fieldset[2]/div[2]/div/div[2]/input').send_keys(
                    email)
                time.sleep(5)

                driver.find_element(By.ID,'btnSiguiente').click()

                try:
                    driver.find_element(By.ID, 'cita_1')
                    winsound.Beep(FREQUENCY, DURATION)
                    time.sleep(1)
                    winsound.Beep(FREQUENCY, DURATION)
                    time.sleep(1)
                    winsound.Beep(FREQUENCY, DURATION)
                    time.sleep(1)
                    time.sleep(900)
                except:
                    print("ситов не найдено")

            except:
                print('офисов не найдено')
                time.sleep(time_z)
        elif usluga == MAIN_USL_3:

            driver.get(MAIN_URL)
            driver.implicitly_wait(80)

            select_element1 = driver.find_element(By.ID, 'form')
            select_object1 = Select(select_element1)
            select_object1.select_by_visible_text(region)

            driver.find_element(By.ID, 'btnAceptar').click()
            driver.implicitly_wait(80)

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

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
        elif usluga == MAIN_USL_4:

            driver.get(MAIN_URL)
            driver.implicitly_wait(80)

            select_element1 = driver.find_element(By.ID, 'form')
            select_object1 = Select(select_element1)
            select_object1.select_by_visible_text(region)

            driver.find_element(By.ID, 'btnAceptar').click()
            driver.implicitly_wait(80)

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

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






