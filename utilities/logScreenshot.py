import os

from utilities.customLogger import LogGen
from utilities.reportScreenshot import add_screenshot


class cLogScreenshot:

    def __init__(self, driver, extra):
        self.driver = driver
        self.extra = extra
        # instantiate the logger class
        self.logger = LogGen.loggen()

    def fLogScreenshot(self, message, pass_=True, log=True, screenshot=True):
        if screenshot and pass_ and log:
            self.logger.info(f'PASS: {message}')
            add_screenshot(self.driver, self.extra, filename='', message=f'PASS: {message}', log='screenshot')
        elif log and pass_:
            self.logger.info(f'PASS: {message}')
            add_screenshot(self.driver, self.extra, filename='', message=f'PASS: {message}', log='logs')
        elif log:
            self.logger.error(f'FAIL: {message}')
            add_screenshot(self.driver, self.extra, filename='', message=f'FAIL: {message}', log='logs')
        elif screenshot and pass_:
            add_screenshot(self.driver, self.extra, filename='', message=f'PASS: {message}', log='screenshot')
        elif screenshot and log:
            add_screenshot(self.driver, self.extra, filename='', message=f'PASS: {message}', log='screenshot')
        elif screenshot:
            add_screenshot(self.driver, self.extra, filename='', message=f'FAIL: {message}', log='screenshot')
        else:
            pass

        # if log and pass_:
        #     self.logger.info(f'PASS: {message}')
        # elif log:
        #     self.logger.error(f'FAIL: {message}')
        # else:
        #     pass
        #
        # if screenshot and pass_:
        #     add_screenshot(self.driver, self.extra, filename='', message=f'PASS: {message}', log='screenshot')
        # elif screenshot:
        #     add_screenshot(self.driver, self.extra, filename='', message=f'FAIL: {message}', log='screenshot')
        # else:
        #     pass
