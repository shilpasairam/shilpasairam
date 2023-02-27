"""
Test will validate the presence of SLR page elements in LiveSLR Application
"""

import pytest
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_LiveSLRPageElements:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.smoketest
    def test_liveslr_page_ele(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate presence of elements in LiveSLR Page"

        try:
            loginPage.driver.get(baseURL)
            loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
            base.go_to_page("SLR_Homepage", env)
            base.presence_of_element("SLR_Population", env)
            base.presence_of_element("SLR_Type", env)
            base.presence_of_element("Data_Report", env)
            base.presence_of_element("NMA_Button", env)
            base.presence_of_element("Preview_Button", env)
            LogScreenshot.fLogScreenshot(message=f"Elements are present in LiveSLR Page",
                                              pass_=True, log=True, screenshot=True)
        except Exception:
            LogScreenshot.fLogScreenshot(
                message=f"Complete login and check if application landing page visible",
                pass_=True, log=True, screenshot=True)
            raise Exception("Element Not Found")
