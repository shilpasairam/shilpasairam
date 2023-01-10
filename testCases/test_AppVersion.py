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
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C29642
    def test_liveslr_app_version(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of AppVersion class
        self.appver = AppVersion(self.driver, extra)

        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")

        # Validating the application version
        self.appver.app_version_check("LiveSLR", "about_live_slr", "about_live_slr_text", "about_live_slr_close")

        # Logging out from the application
        self.loginPage.logout("liveslr_logout_button")

    @pytest.mark.C29577
    @pytest.mark.C29826
    def test_liveref_app_version(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)                
        # Creating object of AppVersion class
        self.appver = AppVersion(self.driver, extra)

        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef")

        # Validating the application version
        self.appver.app_version_check("LiveRef", "about_live_ref", "about_live_ref_text", "about_live_ref_close")

        # Checking the absence of Glossary option
        res_list = []
        eles = self.base.select_elements("glossary_nav_bar", UnivWaitFor=3)
        for ele in eles:
            res_list.append(ele.text)
        
        if not search("Glossary", res_list[0]):
            self.LogScreenshot.fLogScreenshot(message=f"Glossary link is not present as expected", 
                                                pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Glossary link is present as expected")

        # Logging out from the application
        self.loginPage.logout("liveref_logout_button")
