from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import os
import errno

driver = ''
try:
    # get webdriver for chrome chromedriver.exe path - this would have to change for everyone
    # driver = webdriver.Chrome(' C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBotRoot\\DebarmentCheckAIBot\\chromedriver_win32_v78\\chromedriver.exe')

    chrome_options = Options()
    chrome_options.add_argument("--headless")

    driver = webdriver.Chrome(chrome_options=chrome_options,
                              executable_path='C:\\Users\\yangb\\PycharmProjects\\DebarmentCheckAIBotRoot\\DebarmentCheckAIBot\\chromedriver_win32_v78\\chromedriver.exe')
    driver.get("https://duckduckgo.com/")
# chrome_options.binary_location = ''
# driver = webdriver.PhantomJS()

# driver.set_window_size(1120, 550)

# driver.maximize_window()  # maxout the window size
except Exception as err:
    print(err)

# driver.close()
# driver.quit()
