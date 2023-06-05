"""
Test will validate the presence of SLR page elements in LiveSLR Application
"""

import pytest
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase

from Pages.LoginPage import LoginPage
from Pages.SLRReportPage import SLRReport
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
            loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL,
                                            env)
            base.presence_of_admin_page_option("SLR_Homepage", env)
            base.go_to_page("SLR_Homepage", env)
            base.presence_of_element("SLR_Population", env)
            base.presence_of_element("SLR_Type", env)
            base.presence_of_element("Category_view", env)
            base.presence_of_element("NMA_Button", env)
            base.presence_of_element("Preview_Button", env)
            LogScreenshot.fLogScreenshot(message=f"Elements are present in LiveSLR Page",
                                         pass_=True, log=True, screenshot=True)
        except Exception:
            LogScreenshot.fLogScreenshot(
                message=f"Complete login and check if application landing page visible",
                pass_=True, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C37558
    @pytest.mark.C38492
    def test_validate_liveslrpage_tooltip(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getsmoketestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Tool tip description in Search LiveSLR Page"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        try:
            # Read population data values
            pop_list = exbase.get_population_data(filepath, 'Sheet1', 'scenario1')
            # Read slrtype data values
            slrtype = exbase.get_slrtype_data(filepath, 'Sheet1', 'scenario1')            

            slrreport.select_data(pop_list[0][0], pop_list[0][1], env)
            slrreport.select_data(slrtype[0][0], slrtype[0][1], env)

            slrreport.validate_liveslrpage_tooltip('scenario1', filepath, env)

        except Exception:
            raise Exception("Unable to select element")
