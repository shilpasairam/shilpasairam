"""
Test case deals with Login and Logout functionality of LiveSLR application
"""

import time
import pytest

from Pages.LoginPage import LoginPage
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_Login:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.smoketest
    def test_login_page(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

    @pytest.mark.smoketest
    def test_logout(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)
        self.loginPage.logout()
