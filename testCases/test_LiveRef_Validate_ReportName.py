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
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    TestData = ReadConfig.getTestdata( "liveref_searchpublications_data" )

    @pytest.mark.C29730
    @pytest.mark.C29826
    def test_validate_downloaded_filename(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef")
        scenarios = ['scenario1']
        for i in scenarios:
            try:
                self.base.go_to_page("searchpublications_button")
                self.base.click("searchpublications_reset_filter")
                self.srchpub.validate_downloaded_filename(i, self.TestData)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in during validation of downloaded filename",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Error in during validation of downloaded filename")
