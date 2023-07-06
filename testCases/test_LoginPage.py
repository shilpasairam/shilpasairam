"""
Test case deals with Login and Logout functionality of LiveSLR application
"""

import time
import pytest
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
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

    @pytest.mark.C41188
    def test_portal_contact_info(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveHTA Portal Contact Information"

        expected_contact_info = ('CYTEL\n'
                                 '1050 Winter St\n'
                                 'Suite 2700\n'
                                 'Waltham, MA 02451\n'
                                 'USA')

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        self.driver.switch_to.window(self.driver.window_handles[0])
        base.click("contact_btn", env)
        actual_contact_info = base.get_text("contact_info", env)

        if actual_contact_info == expected_contact_info:
            LogScreenshot.fLogScreenshot(message=f"Correct Cytel company address updated on Contact page.",
                                         pass_=True, log=True, screenshot=True)
        else:
            LogScreenshot.fLogScreenshot(
                message=f"Mismatch found in Cytel company address on Contact page. Expected Contact address is: "
                        f"{expected_contact_info} and Actual Contact address is: {actual_contact_info}",
                pass_=False, log=True, screenshot=True)
            raise Exception(f"Mismatch found in Cytel company address on Contact page")
