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
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C29584
    @pytest.mark.C29826
    def test_tabname_changes(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)

        # Invoking the methods from loginpage
        # self.loginPage.driver.get(self.baseURL)
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

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in during validation of tab names",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Error in during validation of tab names")

        # Logging out from the application
        self.loginPage.logout("liveref_logout_button")

    @pytest.mark.C29826
    def test_liveref_validate_duplicate_entries_in_admin_panel(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)

        # Invoking the methods from loginpage
        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef")

        try:
            admin_eles = self.base.select_elements("liveref_admin_panel_list")
            admin_eles_text = []
            for i in admin_eles:
                admin_eles_text.append(i.text)
            if len(admin_eles_text) == len(set(admin_eles_text)):
                self.LogScreenshot.fLogScreenshot(message=f"There are no duplicate navigation links on the left panel",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Found duplicate navigation links on the left panel. "
                                                          f"Values are : {admin_eles_text}",
                                                  pass_=False, log=True, screenshot=True)

        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error in during validation of presence of navigation links",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in during validation of presence of navigation links")

        # Logging out from the application
        self.loginPage.logout("liveref_logout_button")
