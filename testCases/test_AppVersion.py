"""
Test case deals with Application version check for LiveSLR and LiveRef
"""

import pytest
from re import search
from Pages.ApplicationVersionCheck import AppVersion
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_AppVersion:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C29642
    def test_liveslr_app_version(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of AppVersion class
        appver = AppVersion(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveSLR Application Version and Build Number"

        # Clearing the Logs before the test start execution
        base.clear_logs()

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        # Validating the application version
        appver.app_version_check("LiveSLR", "about_live_slr", "about_live_slr_text", "about_live_slr_close", env)

        # Logging out from the application
        loginPage.logout("liveslr_logout_button", env)

    @pytest.mark.C29577
    @pytest.mark.C29826
    def test_liveref_app_version(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)                
        # Creating object of AppVersion class
        appver = AppVersion(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveRef Application Version and Build Number"

        # Clearing the Logs before the test start execution
        base.clear_logs()

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)

        # Validating the application version
        appver.app_version_check("LiveRef", "about_live_ref", "about_live_ref_text", "about_live_ref_close", env)

        # Checking the absence of Glossary option
        res_list = []
        eles = base.select_elements("glossary_nav_bar", env, UnivWaitFor=3)
        for ele in eles:
            res_list.append(ele.text)
        
        if not search("Glossary", res_list[0]):
            LogScreenshot.fLogScreenshot(message=f"Glossary link is not present as expected",
                                              pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Glossary link is present as expected")

        # Logging out from the application
        loginPage.logout("liveref_logout_button", env)
