"""
Test case deals with validation of Tab Names for different pages under LiveRef application
"""

from re import search
import time
import pytest

from Pages.Base import Base
from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_TabNames:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C29584
    def test_tabname_changes(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)

        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef")

        page_locs = ['searchpublications_button', 'liveref_importpublications_button',
                     'liveref_view_import_status_button', 'liveref_manageindications_button',
                     'liveref_managepopulations_button', 'managesourcedata_button',
                     'liveref_managecatevidence_button', 'liveref_manageproducts_button',
                     'liveref_managecountries_button', 'liveref_dashboard']

        for i in page_locs:
            try:
                self.base.go_to_page(i)
                page_name = self.base.get_text(i)
                self.base.assertPageTitle("Cytel LiveRef", UnivWaitFor=10)
                self.LogScreenshot.fLogScreenshot(message=f"Tab Name for '{page_name}' page is as expected.",
                                                  pass_=True, log=True, screenshot=True)

            except:
                pass

        # Logging out from the application
        self.loginPage.liveref_logout()
