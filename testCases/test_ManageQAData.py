"""
Test will validate Manage QA Data Page

"""

import os
import pytest
from datetime import date

from Pages.LoginPage import LoginPage
from Pages.ManageQADataPage import ManageQADataPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ManageQADataPage:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getmanageqadatapath()
    slrfilepath = ReadConfig.getslrtestdata()

    # @pytest.mark.manageqadataUI
    def test_add_qa_data(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        self.mngqadata = ManageQADataPage(self.driver, extra)
        # Get StudyType and Files path to upload Managae QA Data
        self.stdy_data = self.mngqadata.get_qa_file_details(self.filepath)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of ManageQAData validation is started***", pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y").replace('/', '')
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngqadata.go_to_manageqadata("manage_qa_data_button")

        try:
            for i in range(1, 4):
                self.mngqadata.add_multiple_manage_qa_data(self.stdy_data, i)
        except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Error while uploading QA File")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of ManageQAData validation is completed***", pass_=True, log=True, screenshot=False)

    # @pytest.mark.manageqadataUI
    def test_overwrite_qa_data(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        self.mngqadata = ManageQADataPage(self.driver, extra)
        # Get StudyType and Files path to Override the existing Managae QA Data
        self.stdy_data_override = self.mngqadata.get_qa_file_details_override(self.filepath)

        self.LogScreenshot.fLogScreenshot(message=f"***Overwriting the ManageQAData validation is started***", pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y").replace('/', '')
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngqadata.go_to_manageqadata("manage_qa_data_button")

        try:
            for j in range(1, 4):
                self.mngqadata.overwrite_multiple_manage_qa_data(self.stdy_data_override, j)
        except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error while overwriting the QA files",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Error while uploading QA File to the same population and same study type")

        self.LogScreenshot.fLogScreenshot(message=f"***Overwriting the ManageQAData validation is completed***", pass_=True, log=True, screenshot=False)
    
    # @pytest.mark.manageqadataUI
    def test_delete_qa_data(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        self.mngqadata = ManageQADataPage(self.driver, extra)
        # Get StudyType and Files path to upload Managae QA Data
        self.stdy_data = self.mngqadata.get_qa_file_details(self.filepath)

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of ManageQAData validation is started***", pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y").replace('/', '')
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngqadata.go_to_manageqadata("manage_qa_data_button")

        try:
            for i in range(1, 4):
                self.mngqadata.del_multiple_manage_qa_data(self.stdy_data, "delete_file_button", "delete_file_popup", i)
        except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage QA Data page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of ManageQAData validation is completed***", pass_=True, log=True, screenshot=False)

    # @pytest.mark.manageqadataendtoendflow
    def test_qafile_compare_with_excelreport(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        self.mngqadata = ManageQADataPage(self.driver, extra)
        # Get StudyType and Files path to upload Managae QA Data
        self.stdy_data = self.mngqadata.get_qa_file_details(self.filepath)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***ManageQAData File comparison started***", pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y").replace('/', '')
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngqadata.go_to_manageqadata("manage_qa_data_button")

        try:
            self.mngqadata.compare_qa_file_with_report(self.stdy_data, "NewImportLogic_1 - Test_Automation_1", self.slrfilepath)
            self.mngqadata.del_data_after_qafile_comparison(self.stdy_data, "NewImportLogic_1 - Test_Automation_1", self.slrfilepath)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage QA Data page",
                pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***ManageQAData File comparison completed***", pass_=True, log=True, screenshot=False)
