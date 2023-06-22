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
    # @pytest.mark.C29826
    def test_tabname_changes(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "LiveRef - Validate Tab Name changes in LiveRef Homepage"

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)

        page_locs = ['searchpublications_button', 'liveref_importpublications_button',
                     'liveref_view_import_status_button', 'liveref_manageindications_button',
                     'liveref_managepopulations_button', 'managesourcedata_button',
                     'liveref_managecatevidence_button', 'liveref_manageproducts_button',
                     'liveref_managecountries_button', 'liveref_dashboard']

        for i in page_locs:
            try:
                base.presence_of_admin_page_option(i, env)
                base.go_to_page(i, env)
                page_name = base.get_text(i, env)
                base.assertPageTitle("Cytel LiveRef", UnivWaitFor=10)
                LogScreenshot.fLogScreenshot(message=f"Tab Name for '{page_name}' page is as expected.",
                                                  pass_=True, log=True, screenshot=True)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in during validation of tab names",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Error in during validation of tab names")

        # Logging out from the application
        loginPage.logout("liveref_logout_button", env)

    # @pytest.mark.C29826
    def test_liveref_validate_duplicate_entries_in_admin_panel(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "LiveRef - Validate duplicate entries in LiveRef Homepage - Admin Panel"

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)

        try:
            # admin_eles = base.select_elements("liveref_admin_panel_list", env)
            # admin_eles_text = []
            # for i in admin_eles:
            #     admin_eles_text.append(i.text)
            admin_eles_text = base.get_texts("liveref_admin_panel_list", env)
            if len(admin_eles_text) == len(set(admin_eles_text)):
                LogScreenshot.fLogScreenshot(message=f"There are no duplicate navigation links on the left panel",
                                                  pass_=True, log=True, screenshot=True)
            else:
                LogScreenshot.fLogScreenshot(message=f"Found duplicate navigation links on the left panel. "
                                                          f"Values are : {admin_eles_text}",
                                                  pass_=False, log=True, screenshot=True)

        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in during validation of presence of navigation links",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in during validation of presence of navigation links")

        # Logging out from the application
        loginPage.logout("liveref_logout_button", env)
