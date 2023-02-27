"""
Test will validate Manage QA Data Page

"""

import os
import pytest
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from Pages.ManageQADataPage import ManageQADataPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ManageQADataPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getmanageqadatapath()
    # slrfilepath = ReadConfig.getslrtestdata()

    @pytest.mark.C27360
    def test_access_manageqadata_page_elements(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getmanageqadatapath(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        mngqadata = ManageQADataPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Manage QA Data page accessibility"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Presence of ManageQAData Page Elements check is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manage_qa_data_button", env)

        pop_val = ['pop1']

        try:
            for i in pop_val:
                mngqadata.access_manageqadata_page_elements(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing ManageQAData page elements",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing ManageQAData page elements")
        
        LogScreenshot.fLogScreenshot(message=f"***Presence of ManageQAData Page Elements check is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C27361
    def test_add_qa_data(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getmanageqadatapath(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)                 
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        mngqadata = ManageQADataPage(self.driver, extra)
        # # Get StudyType and Files path to upload Managae QA Data
        # self.stdy_data = self.mngqadata.get_qa_file_details(self.filepath)

        request.node._tcid = caseid
        request.node._title = "Validate Addition of QA Data file"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Addtion of ManageQAData validation is started***",
                                          pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manage_qa_data_button", env)

        pop_val = ['pop1', 'pop2']

        try:
            for i in pop_val:
                mngqadata.add_manage_qa_data_with_invalidfile(i, filepath, env)
                mngqadata.add_multiple_manage_qa_data(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing ManageQAData page",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Error while uploading QA File")
        
        LogScreenshot.fLogScreenshot(message=f"***Addtion of ManageQAData validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C27361
    def test_overwrite_qa_data(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getmanageqadatapath(env)  
        # Instantiate the Base class
        base = Base(self.driver, extra)                
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        mngqadata = ManageQADataPage(self.driver, extra)
        # # Get StudyType and Files path to Override the existing Managae QA Data
        # self.stdy_data_override = self.mngqadata.get_qa_file_details_override(self.filepath)

        request.node._tcid = caseid
        request.node._title = "Validate Overwriting of existing QA Data file"

        LogScreenshot.fLogScreenshot(message=f"***Overwriting the ManageQAData validation is started***",
                                          pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manage_qa_data_button", env)

        pop_val = ['pop1', 'pop2']

        try:
            for j in pop_val:
                mngqadata.overwrite_multiple_manage_qa_data(j, filepath, env)
                mngqadata.del_multiple_manage_qa_data(j, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error while overwriting the QA files",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Error while uploading QA File to the same population and same study type")

        LogScreenshot.fLogScreenshot(message=f"***Overwriting the ManageQAData validation is completed***",
                                          pass_=True, log=True, screenshot=False)
    
    @pytest.mark.C27362
    def test_delete_qa_data(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getmanageqadatapath(env) 
        # Instantiate the Base class
        base = Base(self.driver, extra)                
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        mngqadata = ManageQADataPage(self.driver, extra)
        # # Get StudyType and Files path to upload Managae QA Data
        # self.stdy_data = self.mngqadata.get_qa_file_details(self.filepath)

        request.node._tcid = caseid
        request.node._title = "Validate Deletion of existing QA Data file"

        LogScreenshot.fLogScreenshot(message=f"***Deletion of ManageQAData validation is started***",
                                          pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manage_qa_data_button", env)

        pop_val = ['pop1', 'pop2']

        try:
            for i in pop_val:
                mngqadata.add_multiple_manage_qa_data(i, filepath, env)
                mngqadata.del_multiple_manage_qa_data(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage QA Data page",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Deletion of ManageQAData validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C30247
    @pytest.mark.C27301
    @pytest.mark.C27306
    def test_qafile_compare_with_excelreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getmanageqadatapath(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)                 
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManageQADataPage class
        mngqadata = ManageQADataPage(self.driver, extra)
        # # Get StudyType and Files path to upload Managae QA Data
        # self.stdy_data = self.mngqadata.get_qa_file_details(self.filepath)

        request.node._tcid = caseid
        request.node._title = "Validate content of QA Data file with Complete Excel Report"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***ManageQAData File comparison started***",
                                          pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manage_qa_data_button", env)

        pop_val = ['pop3']

        try:
            for i in pop_val:
                mngqadata.compare_qa_file_with_report(i, filepath, env)
                mngqadata.del_data_after_qafile_comparison(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage QA Data page",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***ManageQAData File comparison completed***",
                                          pass_=True, log=True, screenshot=False)
