"""
Test case deals with Login and Logout functionality of LiveSLR application
"""

import time
import pytest

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_Login:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.smoketest
    def test_login_page(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)        
        # Invoking the methods from loginpage
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

    @pytest.mark.smoketest
    def test_logout(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)        
        # Invoking the methods from loginpage
        self.loginPage.driver.get(baseURL)      
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.loginPage.logout("liveslr_logout_button", env)
