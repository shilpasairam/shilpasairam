"""
Test will validate Excluded Studies Page

"""

import os
import pytest
from datetime import date

from Pages.LoginPage import LoginPage
from Pages.ExcludedStudiesPage import ExcludedStudiesPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ExcludedStudiesPage:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getexcludedstudiespath()
    slrfilepath = ReadConfig.getslrtestdata()

    def test_add_excluded_study(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of Excluded Studies validation is started***", pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        # Check Manage Excluded Studies option is present in admin page or not
        self.exstdy.presence_of_elements("excluded_studies_link")
        # Go to ExcludedStudies Page
        self.exstdy.go_to_excludedstudies("excluded_studies_link")

        pop_list = ['pop1']

        try:
            for i in pop_list:
                self.exstdy.add_multiple_excluded_study_data(i, self.filepath)
        except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Studies page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Excluded Studies page")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of Excluded Studies validation is completed***", pass_=True, log=True, screenshot=False)

    def test_update_excluded_study(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Updating the existing Excluded Studies file validation is started***", pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.exstdy.go_to_excludedstudies("excluded_studies_link")

        pop_list = ['pop1']

        try:
            for i in pop_list:
                self.exstdy.update_multiple_excluded_study_data(i, self.filepath)
        except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Studies page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Excluded Studies page")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Updating the existing Excluded Studies file validation is completed***", pass_=True, log=True, screenshot=False)

    def test_delete_excluded_study(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Deleting the existing Excluded Studies file validation is started***", pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.exstdy.go_to_excludedstudies("excluded_studies_link")

        pop_list = ['pop1']

        try:
            for i in pop_list:
                self.exstdy.del_multiple_excluded_study_data(i, self.filepath)
        except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Studies page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Excluded Studies page")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Deleting the existing Excluded Studies file validation is completed***", pass_=True, log=True, screenshot=False)

    def test_excluded_study_compare_with_excel_report(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)
        # # Get StudyType and Files path to upload Managae QA Data
        # self.stdy_data = self.exstdy.get_study_file_details(self.filepath)

        # # Removing the files before the test runs
        # if os.path.exists(f'ActualOutputs'):
        #     for root, dirs, files in os.walk(f'ActualOutputs'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Excluded Studies File comparison started***", pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)

        try:
            self.exstdy.compare_excludedstudy_file_with_report(self.filepath, "NewImportLogic_1 - Test_Automation_1", self.slrfilepath)
            self.exstdy.del_after_studyfile_comparison(self.filepath, "NewImportLogic_1 - Test_Automation_1", self.slrfilepath)
        except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Studies page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Error while comparing the Excluded Studies file with Completed Excel Report")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Excluded Studies File comparison completed***", pass_=True, log=True, screenshot=False)
