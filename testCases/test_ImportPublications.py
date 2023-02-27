"""
Test will validate the Import publications page
"""

import os
import pytest
from Pages.Base import Base
from Pages.ImportPublicationsPage import ImportPublicationPage

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ImportPublicationPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getimportpublicationsdata()

    @pytest.mark.C30246
    @pytest.mark.C27544
    @pytest.mark.C27546
    @pytest.mark.C27381
    @pytest.mark.C28987
    def test_upload_and_del_extraction_template_success(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Upload Extraction Template with Success Icon"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop1']

        for index, i in enumerate(pop_list):
            try:
                imppubpage.upload_file_with_success(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text",
                                            "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is completed***",
                                          pass_=True, log=True, screenshot=False)
    
    @pytest.mark.C30246
    @pytest.mark.C27544
    @pytest.mark.C27546
    @pytest.mark.C27381
    @pytest.mark.C28987
    def test_upload_and_del_extraction_template_failure(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Upload Extraction Template with Failure Icon"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop2']

        for index, i in enumerate(pop_list):
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text",
                                            "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C27547
    def test_upload_extraction_template_with_header_mismatch(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Upload Extraction Template with Column Header Mismatch"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Header Mismatch "
                                                  f"validation is started***", pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop3']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text",
                                            "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Header Mismatch validation "
                                                  f"is completed***", pass_=True, log=True, screenshot=False)

    @pytest.mark.C27379
    def test_upload_extraction_template_with_letters_in_publication_identifier(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Upload Extraction Template with Letters in Publication Identifier column"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with letters in Publication "
                                                  f"Identifier validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop4']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text",
                                            "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with letters in Publication "
                                                  f"Identifier validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C27380
    def test_upload_extraction_template_with_empty_value_in_publication_identifier(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Upload Extraction Template with Empty value in Publication Identifier column"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                                  f"Identifier validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop5']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text",
                                            "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                                  f"Identifier validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C28986
    def test_upload_extraction_template_with_duplicate_value_in_FA18_column(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Upload Extraction Template with Duplicate value in Interventions(per arm) column"

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Duplicate value in Interventions(per arm) validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop6']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text",
                                            "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Duplicate value in Interventions(per arm) validation is completed***",
                                          pass_=True, log=True, screenshot=False)
