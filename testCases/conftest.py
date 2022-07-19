import os

import pytest
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import ChromeOptions, EdgeOptions
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium import webdriver


# @pytest.fixture(params=["chrome", "edge"])
# def init_driver(request):
#     if request.param == "chrome":
#         options = ChromeOptions()
#         options.add_experimental_option("detach", True)
#         options.add_experimental_option('excludeSwitches', ['enable-logging'])
#         web_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
#         web_driver.maximize_window()
#     if request.param == "edge":
#         options = EdgeOptions()
#         options.add_experimental_option("detach", True)
#         web_driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
#         web_driver.maximize_window()
#
#     request.cls.driver = web_driver
#     web_driver.implicitly_wait(10)
#
#     yield
#     web_driver.close()

@pytest.fixture()
def init_driver(request):
    options = ChromeOptions()
    prefs = {'profile.default_content_setting_values.automatic_downloads': 1}
    prefs["profile.default_content_settings.popups"] = 0
    # getcwd should always return the root directory of the framework
    prefs["download.default_directory"] = f"{os.getcwd()}\\ActualOutputs"
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("detach", True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    web_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    request.cls.driver = web_driver
    web_driver.maximize_window()
    web_driver.implicitly_wait(10)

    yield
    web_driver.quit()
