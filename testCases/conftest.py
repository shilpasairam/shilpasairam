import configparser
import os
import shutil
import zipfile

from py.xml import html
import pytest
import platform
import socket
import psutil
import datetime
import glob
from pyhtml2pdf import converter
from pathlib import Path
from utilities.readProperties import ReadConfig
from selenium.webdriver.chrome.service import Service
from selenium.webdriver import ChromeOptions, EdgeOptions
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium import webdriver


def pytest_addoption(parser):
    config = configparser.ConfigParser()
    config.read(os.getcwd()+"\\Configurations\\config.ini")
    parser.addoption("--env", action="store", default=config.get('commonInfo', 'environment'))


@pytest.fixture()
def env(request):
    return request.config.getoption("--env")


@pytest.fixture()
def caseid(request):
    return request.config.getoption("-m")

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
    options.add_argument("--headless")
    options.add_argument("--window-size=1920,1080")
    options.add_argument('--disable-gpu')
    options.add_argument("--log-level=3")  # fatal
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("detach", True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    web_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    params = {'behavior': 'allow', 'downloadPath': f"{os.getcwd()}\\ActualOutputs"}
    web_driver.execute_cdp_cmd('Page.setDownloadBehavior', params)

    request.cls.driver = web_driver
    # web_driver.maximize_window()
    web_driver.implicitly_wait(10)

    # Yield will act as Teardown method which automatically quit driver once the test is completed
    yield
    web_driver.quit()


# ############ pytest HTML Report ##############
@pytest.hookimpl(tryfirst=True)
def pytest_sessionstart(session):
    if os.path.exists(f'Logs'):
        # Clearing the logs before test runs
        open(".\\Logs\\testlog.log", "w").close()

    # Removing the screenshots and results reports before the test runs
    if os.path.exists(f'Reports'):
        for root, dirs, files in os.walk(f'Reports'):
            for file in files:
                os.remove(os.path.join(root, file))
        for root, dirs, files in os.walk(f'Reports/screenshots'):
            for file in files:
                os.remove(os.path.join(root, file))

    # Removing the downloaded report files before the test runs
    if os.path.exists(f'ActualOutputs'):
        for root, dirs, files in os.walk(f'ActualOutputs'):
            for file in files:
                os.remove(os.path.join(root, file))    


# @pytest.hookimpl(tryfirst=True)
def pytest_configure(config):
    # config._metadata['Test Suite'] = ReadConfig.getTestdata("liveref_data").replace(".xlsx","")
    config._metadata['Machine Configuration'] = f'{socket.getfqdn()}, {platform.processor()}, ' \
                                                f'{str(round(psutil.virtual_memory().total / (1024.0 **3)))+" GB"}'
    # config._metadata['Application URL'] = ReadConfig.getApplicationURL()
    config._metadata['Environment'] = config.getoption("--env")
    config._metadata['Tester'] = ReadConfig.getUserName()
    config._metadata['OS'] = platform.platform()
    config._metadata['Browser'] = 'Chrome'

    # set custom options only if none are provided from command line
    # create report target dir
    reports_dir = Path('Reports')
    reports_dir.mkdir(parents=True, exist_ok=True)
    # custom report file
    report = reports_dir / f"{config.getoption('-m')}_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.html"
    # adjust plugin options
    config.option.htmlpath = report
    config.option.self_contained_html = True


def pytest_metadata(metadata):
    metadata.pop("JAVA_HOME", None)
    metadata.pop("Plugins", None)
    metadata.pop("Packages", None)
    metadata.pop("Platform", None)
    metadata.pop("Python", None)


def pytest_html_report_title(report):
    """ modifying the title  of html report"""
    report.title = "Automation Report"


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    report = outcome.get_result()
    setattr(report, "duration_formatter", "%M:%S")
    report._tcid = getattr(item, '_tcid', '')
    report._title = getattr(item, '_title', '')


@pytest.mark.optionalhook
def pytest_html_results_table_header(cells):
    del cells[1]
    cells.insert(1, html.th('TC ID'))
    cells.insert(2, html.th('TC Title'))
    cells.pop()


@pytest.mark.optionalhook
def pytest_html_results_table_row(report, cells):
    del cells[1]
    cells.insert(1, html.td(getattr(report, '_tcid', '')))
    cells.insert(2, html.td(getattr(report, '_title', '')))
    cells.pop()

# # This deletes the log window in the report
# def pytest_html_results_table_html(data):
#         del data[-1]


@pytest.hookimpl(trylast=True)
def pytest_sessionfinish(session, exitstatus):

    # Converting the HTML report file to PDF format
    filename = Path(glob.glob('Reports//*.html')[0]).stem
    path = os.path.abspath(f'Reports//{filename}.html')
    converter.convert(f'file:///{path}?collapsed=Skipped', f"{filename}.pdf")

    dir_list = ['ActualOutputs', 'Logs', 'Reports']
    zip_name = f"{session.config.getoption('-m')}_results.zip"

    zip_file = zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED)
    for dir in dir_list:
        for dirpath, dirnames, filenames in os.walk(dir):
            for filename in filenames:
                zip_file.write(
                    os.path.join(dirpath, filename),
                    os.path.relpath(os.path.join(dirpath, filename), os.path.join(dir_list[0], '..')))

    zip_file.close()
