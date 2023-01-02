"""
Test case deals with Application version check for LiveSLR and LiveRef
"""

import pytest
from Pages.ApplicationVersionCheck import AppVersion

from Pages.LoginPage import LoginPage
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_AppVersion:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.smoketest
    def test_liveslr_app_version(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of AppVersion class
        self.appver = AppVersion(self.driver, extra)
        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.appver.app_version_check("LiveSLR", "about_live_slr", "about_live_slr_text", "about_live_slr_close")
        self.loginPage.logout("liveslr_logout_button")

    @pytest.mark.smoketest
    def test_liveref_app_version(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of AppVersion class
        self.appver = AppVersion(self.driver, extra)
        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef")
        self.appver.app_version_check("LiveRef", "about_live_ref", "about_live_ref_text", "about_live_ref_close")
        self.loginPage.logout("liveref_logout_button")
