import datetime
import time

from pytest_html import extras
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec


############## adding all screenshots to html report ##############
def add_screenshot(driver, extra, filename, message, log='logs', savepath='screenshots'):
    """
    input -
    savepath - path where screenshot is to be saved
    filename - name of the screenshot
    message - message to be displayed in the report next to the screenshot
    """
    timenow = datetime.datetime.now()
    filepath = f'{savepath}\\{filename}_{timenow.strftime("%Y%m%d%H%M%S")}.png'
    # save screenshot
    time.sleep(2)
    driver.save_screenshot(f'.\\Reports\\{filepath}')
    if log == 'logs':
        # Add logs to HTML report
        html = f'''
                <table style="width: 100%">
                    <tr>
                        <td style="width: 10%">{timenow.strftime("%Y-%m-%d %H:%M:%S")}</td>
                        <td style="width: 40%;">{message}</td>
                    </tr>
                </table>
                '''
    else:
        # Add screenshot to HTML report
        html = f'''
            <table style="width: 100%">
                <tr>
                    <td style="width: 10%">{timenow.strftime("%Y-%m-%d %H:%M:%S")}</td>
                    <td style="width: 40%;">{message}</td>
                    <td style="width: 50%;"><img src= {filepath} alt="screenshot" style="width:304px;height:228px;" onclick="window.open(this.src)"></td>
                </tr>
            </table>
            '''
    extra.append(extras.html(html))
