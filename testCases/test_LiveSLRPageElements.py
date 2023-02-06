"""
Test will validate the presence of SLR page elements in LiveSLR Application
"""

import pytest
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_LiveSLRPageElements:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.smoketest
    def test_liveslr_page_ele(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        try:
            self.loginPage.driver.get(baseURL)
            self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
            self.base.go_to_page("SLR_Homepage", env)
            self.base.presence_of_element("SLR_Population", env)
            self.base.presence_of_element("SLR_Type", env)
            self.base.presence_of_element("Data_Report", env)
            self.base.presence_of_element("NMA_Button", env)
            self.base.presence_of_element("Preview_Button", env)
            self.LogScreenshot.fLogScreenshot(message=f"Elements are present in LiveSLR Page",
                                              pass_=True, log=True, screenshot=True)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Complete login and check if application landing page visible",
                pass_=True, log=True, screenshot=True)
            raise Exception("Element Not Found")
