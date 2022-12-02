import time

import pandas as pd
from Pages.Base import Base
from selenium.webdriver.common.by import By
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot


class LoginPage(Base):
    app_filepath = ReadConfig.getappversionfilepath()

    """Constructor of the Login Page class"""
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
    def complete_login(self, username, password, app_url):
        """
        application login page must be opened before calling this method
        """
        try:
            self.LogScreenshot.fLogScreenshot(message=f'Launching the Application: {app_url}',
                                              pass_=True, log=True, screenshot=False)
            # enter username
            self.input_text("username_textbox", username, UnivWaitFor=10)
            self.LogScreenshot.fLogScreenshot(message='Enter username',
                                              pass_=True, log=True, screenshot=True)

            # enter password
            self.input_text("password_textbox", password, UnivWaitFor=3)
            self.LogScreenshot.fLogScreenshot(message='Enter password',
                                              pass_=True, log=True, screenshot=True)

            # submit password
            self.click("login_button", UnivWaitFor=3)
            self.LogScreenshot.fLogScreenshot(message='Submit Credentials',
                                              pass_=True, log=True, screenshot=False)
            
            self.jsclick("launch_live_slr", UnivWaitFor=5)
            self.driver.switch_to.window(self.driver.window_handles[1])
        except Exception:
            pass
        # check whether the login page opened or not
        try:
            self.assertPageTitle("Cytel LiveSLR", UnivWaitFor=10)
            time.sleep(5)
            self.LogScreenshot.fLogScreenshot(
                message=f"Login Successful. Application home page displayed",
                pass_=True, log=True, screenshot=True)
            self.click("about_live_slr", UnivWaitFor=3)
            time.sleep(1)
            self.check_app_version("about_live_slr_text", self.app_filepath, "LiveSLR")
            time.sleep(1)
            self.click("about_live_slr_close", UnivWaitFor=3)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Login Unsuccessful. Please check the application availability and try again",
                pass_=False, log=True, screenshot=True)
            raise Exception("Login Unsuccessful")

    def logout(self):
        """
        this method is to logout from the application and arrive at the login landing page
        """
        try:
            # click on logout
            self.click("liveslr_logout_button", UnivWaitFor=10)
            self.LogScreenshot.fLogScreenshot(message=f"Complete logout",
                                              pass_=True, log=True, screenshot=False)

            self.assertPageTitle("PSE - A Cytel Company - Login", UnivWaitFor=10)
            self.LogScreenshot.fLogScreenshot(message=f"Logout is done. Application Login page is displayed",
                                              pass_=True, log=True, screenshot=True)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Logout Unsuccessful. Please check the application availability and try again",
                pass_=False, log=True, screenshot=False)
            raise Exception("Logout unsuccessful")

    """Page Actions for LiveRef Login page"""
    def liveref_complete_login(self, username, password, app_url):
        """
        application login page must be opened before calling this method
        """
        try:
            self.LogScreenshot.fLogScreenshot(message=f'Launching the Application: {app_url}',
                                              pass_=True, log=True, screenshot=False)
            # enter username
            self.input_text("username_textbox", username, UnivWaitFor=10)
            self.LogScreenshot.fLogScreenshot(message='Enter username',
                                              pass_=True, log=True, screenshot=True)

            # enter password
            self.input_text("password_textbox", password, UnivWaitFor=3)
            self.LogScreenshot.fLogScreenshot(message='Enter password',
                                              pass_=True, log=True, screenshot=True)

            # submit password
            self.click("login_button", UnivWaitFor=3)
            self.LogScreenshot.fLogScreenshot(message='Submit Credentials',
                                              pass_=True, log=True, screenshot=False)
            
            self.jsclick("launch_liveref", UnivWaitFor=5)
            self.driver.switch_to.window(self.driver.window_handles[1])
        except Exception:
            pass
        # check whether the login page opened or not
        try:
            self.assertPageTitle("Cytel LiveRef", UnivWaitFor=10)
            time.sleep(5)
            self.LogScreenshot.fLogScreenshot(
                message=f"Login Successful. Application home page displayed",
                pass_=True, log=True, screenshot=True)
            self.click("about_live_ref", UnivWaitFor=3)
            time.sleep(1)
            self.check_app_version("about_live_ref_text", self.app_filepath, "LiveRef")
            time.sleep(1)
            self.click("about_live_ref_close", UnivWaitFor=3)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Login Unsuccessful. Please check the application availability and try again",
                pass_=False, log=True, screenshot=True)
            raise Exception("Login Unsuccessful")

    def liveref_logout(self):
        """
        this method is to logout from the application and arrive at the login landing page
        """
        try:
            # click on logout
            self.click("liveref_logout_button", UnivWaitFor=10)
            self.LogScreenshot.fLogScreenshot(message=f"Complete logout",
                                              pass_=True, log=True, screenshot=False)

            self.assertPageTitle("PSE - A Cytel Company - Login", UnivWaitFor=10)
            self.LogScreenshot.fLogScreenshot(message=f"Logout is done. Application Login page is displayed",
                                              pass_=True, log=True, screenshot=True)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Logout Unsuccessful. Please check the application availability and try again",
                pass_=False, log=True, screenshot=False)
            raise Exception("Logout unsuccessful")
