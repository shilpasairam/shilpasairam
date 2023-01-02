"""
Test will validate the Import publications page
"""

import os
import pytest
from Pages.ImportPublicationsPage import ImportPublicationPage

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ImportPublicationPage:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getimportpublicationsdata()

    @pytest.mark.C30246
    @pytest.mark.C27544
    @pytest.mark.C27546
    @pytest.mark.C27381
    @pytest.mark.C28987
    def test_upload_and_del_extraction_template(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

        pop_list = ['pop1', 'pop2']

        for index, i in enumerate(pop_list):
            try:
                self.imppubpage.upload_file(i, self.filepath, index)
                self.imppubpage.delete_file(i, self.filepath, "file_status_popup_text",
                                            "upload_table_rows")
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C27547
    def test_upload_extraction_template_with_header_mismatch(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Header Mismatch "
                                                  f"validation is started***", pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

        pop_list = ['pop3']

        for i in pop_list:
            try:
                self.imppubpage.upload_file_with_errors(i, self.filepath)
                self.imppubpage.delete_file(i, self.filepath, "file_status_popup_text",
                                            "upload_table_rows")
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Header Mismatch validation "
                                                  f"is completed***", pass_=True, log=True, screenshot=False)

    @pytest.mark.C27379
    def test_upload_extraction_template_with_letters_in_publication_identifier(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with letters in Publication "
                                                  f"Identifier validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

        pop_list = ['pop4']

        for i in pop_list:
            try:
                self.imppubpage.upload_file_with_errors(i, self.filepath)
                self.imppubpage.delete_file(i, self.filepath, "file_status_popup_text",
                                            "upload_table_rows")
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with letters in Publication "
                                                  f"Identifier validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C27380
    def test_upload_extraction_template_with_empty_value_in_publication_identifier(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                                  f"Identifier validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

        pop_list = ['pop5']

        for i in pop_list:
            try:
                self.imppubpage.upload_file_with_errors(i, self.filepath)
                self.imppubpage.delete_file(i, self.filepath, "file_status_popup_text",
                                            "upload_table_rows")
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                                  f"Identifier validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C28986
    def test_upload_extraction_template_with_duplicate_value_in_FA18_column(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                                  f"Identifier validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

        pop_list = ['pop6']

        for i in pop_list:
            try:
                self.imppubpage.upload_file_with_errors(i, self.filepath)
                self.imppubpage.delete_file(i, self.filepath, "file_status_popup_text",
                                            "upload_table_rows")
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                                  f"Identifier validation is completed***",
                                          pass_=True, log=True, screenshot=False)
