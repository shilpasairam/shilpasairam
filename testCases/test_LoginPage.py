"""
Test case deals with Login and Logout functionality of LiveSLR application
"""

import time
import pytest
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_Login:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.smoketest
    def test_login_page(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveSLR Login Scenario"

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

    @pytest.mark.smoketest
    def test_logout(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveSLR Logout Scenario"

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)      
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        loginPage.logout("liveslr_logout_button", env)
