"""
Test case deals in validating the downloaded report name under LiveRef application
"""
import os
import pytest
from Pages.Base import Base
from Pages.LiveRef_SearchPublicationsPage import SearchPublicationsPage
from Pages.LoginPage import LoginPage
from Pages.SLRReportPage import SLRReport
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen


@pytest.mark.usefixtures("init_driver")
class Test_SearchPublications_DownloadedFilename:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    TestData = ReadConfig.getTestdata("liveref_searchpublications_data")

    @pytest.mark.C29730
    @pytest.mark.C29826
    def test_validate_downloaded_filename(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveRefAppURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        srchpub = SearchPublicationsPage(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "LiveRef - Validate Downloaded Excel report filename"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        scenarios = ['scenario1']
        for i in scenarios:
            try:
                base.go_to_page("searchpublications_button", env)
                base.click("searchpublications_reset_filter", env)
                srchpub.validate_downloaded_filename(i, self.TestData, env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in during validation of downloaded filename",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Error in during validation of downloaded filename")
