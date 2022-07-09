from selenium import webdriver

from consts import OPTION_1, OPTION_2, OPTION_3, NAME_DRIVER

options = webdriver.ChromeOptions()
options.add_argument(OPTION_1)
options.add_argument(OPTION_2)
options.add_argument(OPTION_3)

driver = webdriver.Chrome(
    executable_path=NAME_DRIVER,
    options=options
)
