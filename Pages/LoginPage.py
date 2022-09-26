import time
from Pages.Base import Base
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot


class LoginPage(Base):

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

    """Page Actions for Login page"""
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
