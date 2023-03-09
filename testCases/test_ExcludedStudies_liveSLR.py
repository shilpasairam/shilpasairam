import os

import pytest
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from Pages.ExcludedStudies_liveSLR import ExcludedStudies_liveSLR
from utilities.readProperties import ReadConfig
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_ExcludedStudies_liveSLR:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getexcludedstudiesliveslrpath()

    @pytest.mark.C30712
    def test_presenceof_excludedstudiesliveslr_into_excelreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Presence of ExcludedStudies_LiveSLR sheet in Complete Excel Report"

        LogScreenshot.fLogScreenshot(message=f"*****Presence of Excluded studies - LiveSLR Sheet in "
                                                  f"Complete Excel Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                exstdy_liveslr.presenceof_excludedstudiesliveslr_into_excelreport(i, self.filepath, env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C30713
    @pytest.mark.C31841
    def test_presenceof_columnnames_in_excludedstudiesliveslrtab(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Presence of ExcludedStudies_LiveSLR sheet column names in Complete Excel Report"

        LogScreenshot.fLogScreenshot(message=f"*****Presence of Excluded studies - LiveSLR Sheet in Complete "
                                                  f"Excel Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for index, i in enumerate(scenarios):
            try:
                exstdy_liveslr.presenceof_columnnames_in_excludedstudiesliveslrtab(i, self.filepath, index, env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C30715
    @pytest.mark.C31843
    def test_validate_excludedstudiesliveslrtab_and_contents_into_excelreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Contents of ExcludedStudies_LiveSLR sheet in Complete Excel Report"

        LogScreenshot.fLogScreenshot(message=f"*****Presence of Excluded studies - LiveSLR Sheet in Complete "
                                                  f"Excel Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for index, i in enumerate(scenarios):
            try:
                exstdy_liveslr.validate_excludedstudiesliveslrtab_and_contents_into_excelreport(i, self.filepath,
                                                                                                     index, env)
            except Exception:
                raise Exception("Unable to select element")
