"""
Test case deals in validating the LiveRef label rebrand
"""
import os
import pytest
from Pages.Base import Base
from Pages.LiveRef_Rebrand_Check import LiveRef_Rebrand
from Pages.LiveRef_SearchPublicationsPage import SearchPublicationsPage
from Pages.LoginPage import LoginPage
from Pages.SLRReportPage import SLRReport
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen


@pytest.mark.usefixtures("init_driver")
class Test_LiveRef_Rebrand:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    TestData = ReadConfig.getTestdata( "liveref_searchpublications_data" )

    @pytest.mark.C27354
    def test_validate_liveref_rebrand(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of LiveRef_Rebrand class
        self.rebrand = LiveRef_Rebrand(self.driver, extra)
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
        try:
            self.rebrand.validate_liveref_rebrand(self.TestData)
        except Exception:
            raise Exception("Mismatch found during LiveRef brand validation")
