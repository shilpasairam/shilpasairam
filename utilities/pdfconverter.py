import sys
import json
import base64

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import staleness_of
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

def send_devtools(driver, cmd, params={}):
    resource = "/session/%s/chromium/send_command_and_get_result" % driver.session_id
    url = driver.command_executor._url + resource
    body = json.dumps({"cmd": cmd, "params": params})
    response = driver.command_executor._request("POST", url, body)

    if not response:
        raise Exception(response.get("value"))

    return response.get("value")


def get_pdf_from_html(
    path, print_options = {}, install_driver=True, exec_path=None
):
    webdriver_options = Options()
    webdriver_prefs = {}
    driver = None

    webdriver_options.add_argument("--headless")
    webdriver_options.add_argument("--disable-gpu")
    webdriver_options.add_argument("--no-sandbox")
    webdriver_options.add_argument("--disable-dev-shm-usage")
    webdriver_options.experimental_options["prefs"] = webdriver_prefs

    webdriver_prefs["profile.default_content_settings"] = {"images": 2}

    if install_driver:
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=webdriver_options
        )
    else:
        driver = webdriver.Chrome(executable_path=exec_path, options=webdriver_options)

    driver.get(path)

    calculated_print_options = {
        "landscape": False,
        "displayHeaderFooter": False,
        "printBackground": True,
        "preferCSSPageSize": True,
    }
    calculated_print_options.update(print_options)
    result = send_devtools(
        driver, "Page.printToPDF", calculated_print_options)
    driver.quit()
    return base64.b64decode(result["data"])
