"""
Test will validate the presence of SLR page elements in LiveSLR Application
"""

import pytest

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_LiveSLRPageElements:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.smoketest
    def test_liveslr_page_ele(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        try:
            self.loginPage.driver.get(self.baseURL)
            self.loginPage.complete_login(self.username, self.password)
            self.liveslrpage.go_to_liveslr("SLR_Homepage")
            self.liveslrpage.presence_of_elements("SLR_Population")
            self.liveslrpage.presence_of_elements("SLR_Type")
            self.liveslrpage.presence_of_elements("Data_Report")
            self.liveslrpage.presence_of_elements("NMA_Button")
            self.liveslrpage.presence_of_elements("Preview_Button")
            self.LogScreenshot.fLogScreenshot(message=f"Elements are present in LiveSLR Page",
                                              pass_=True, log=True, screenshot=True)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Complete login and check if application landing page visible",
                pass_=True, log=True, screenshot=True)
            raise Exception("Element Not Found")
