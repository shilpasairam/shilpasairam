import os

import pytest

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from Pages.ExcludedStudies_liveSLR import ExcludedStudies_liveSLR
from utilities.readProperties import ReadConfig
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_ExcludedStudies_liveSLR:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getexcludedstudiesliveslrpath()

    @pytest.mark.C30712
    def test_presenceof_excludedstudiesliveslr_into_excelreport(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Presence of Excluded studies - LiveSLR Sheet in "
                                                  f"Complete Excel Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                self.exstdy_liveslr.presenceof_excludedstudiesliveslr_into_excelreport(i, self.filepath)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C30713
    @pytest.mark.C31841
    def test_presenceof_columnnames_in_excludedstudiesliveslrtab(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Presence of Excluded studies - LiveSLR Sheet in Complete "
                                                  f"Excel Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for index, i in enumerate(scenarios):
            try:
                self.exstdy_liveslr.presenceof_columnnames_in_excludedstudiesliveslrtab(i, self.filepath, index)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C30715
    @pytest.mark.C31843
    def test_validate_excludedstudiesliveslrtab_and_contents_into_excelreport(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Presence of Excluded studies - LiveSLR Sheet in Complete "
                                                  f"Excel Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for index, i in enumerate(scenarios):
            try:
                self.exstdy_liveslr.validate_excludedstudiesliveslrtab_and_contents_into_excelreport(i, self.filepath,
                                                                                                     index)
            except Exception:
                raise Exception("Unable to select element")
