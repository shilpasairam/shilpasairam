"""
Test case deals with Application access check for LiveSLR and LiveRef
"""

import pytest
from re import search
from Pages.Base import Base
from Pages.ApplicationVersionCheck import AppVersion
from Pages.LoginPage import LoginPage

from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_AppAccess:
    # baseURL = ReadConfig.getApplicationURL()
    testdata = ReadConfig.getTestdata("app_access_data")

    # """Constructor of the Test_AppAccess class"""
    # def __init__(self, driver, extra):
    #     # initializing the driver from base class
    #     super().__init__(driver, extra)
    #     # Creating object of loginpage class
    #     self.loginPage = LoginPage(driver, extra)        
    #     # Creating object of AppVersion class
    #     self.appver = AppVersion(driver, extra)        

    @pytest.mark.parametrize(
        "scenarios, name",
        [
            ("scenario1", "Admin User"),
            ("scenario2", "Staff User"),
            ("scenario3", "Client User")
        ]
    )
    @pytest.mark.C37915
    def test_liveslr_application_access(self, scenarios, name, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)        
        # Creating object of AppVersion class
        appver = AppVersion(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveSLR Application Access"

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        # self.loginPage.complete_login(username, password, "launch_live_slr", "Cytel LiveSLR")        
        
        # Validating the application version
        appver.validate_liveslr_page_access(scenarios, self.testdata, baseURL, env)

        # Logging out from the application
        loginPage.logout("liveslr_logout_button", env)        

    @pytest.mark.parametrize(
        "scenarios, name",
        [
            ("scenario1", "Admin User"),
            ("scenario2", "Staff User"),
            ("scenario3", "Client User")
        ]
    )
    @pytest.mark.C37915
    def test_liveref_application_access(self, scenarios, name, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)        
        # Creating object of AppVersion class
        appver = AppVersion(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveRef Application Access"

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        # self.loginPage.complete_login(username, password, "launch_liveref", "Cytel LiveRef")        
        
        # Validating the application version
        appver.validate_liveref_page_access(scenarios, self.testdata, baseURL, env)

        # Logging out from the application
        loginPage.logout("liveref_logout_button", env)        
