import time

import pandas as pd
from Pages.Base import Base
from selenium.webdriver.common.by import By
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot


class AppVersion(Base):
    app_filepath = ReadConfig.getappversionfilepath()

    """Constructor of the AppVersion class"""
    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # instantiate the logger class
        self.logger = LogGen.loggen()
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
    
    def get_expected_application_version(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        version_details = df.loc[df['Application'] == locatorname]['Version'].dropna().to_list()
        return version_details
    
    def check_app_version(self, locatorname, filepath, app_name):
        expected_app_version = self.get_expected_application_version(filepath, app_name)
        actual_app_version = self.base.get_text(locatorname)
        if expected_app_version[0] in actual_app_version:
            self.LogScreenshot.fLogScreenshot(
                message=f"Application version is as expected. Application version is : {actual_app_version}",
                pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(
                message=f"Mismatch found in Applicaiton version. Actual App version is : {actual_app_version} "
                        f"and Expected App version is : {expected_app_version}",
                pass_=False, log=True, screenshot=True)
            raise Exception(f"Mismatch found in Applicaiton version.")

    """Page Actions for LiveSLR Login page"""
    def app_version_check(self, application, about_button, about_text, about_close_btn):
        """
        application login page must be opened before calling this method
        """
        try:
            self.click(about_button, UnivWaitFor=3)
            time.sleep(1)
            self.check_app_version(about_text, self.app_filepath, application)
            time.sleep(1)
            self.click(about_close_btn, UnivWaitFor=3)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Mismatch found in Application Version",
                pass_=False, log=True, screenshot=True)
            raise Exception("Mismatch found in Application Version")