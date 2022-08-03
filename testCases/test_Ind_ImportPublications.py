"""
Test will validate the Import publications page
"""

import time
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
    filepath = ReadConfig.getadminpagedata()

    def test_upload_extraction_template(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)
        # Read extraction sheet values
        self.file_upload = self.imppubpage.get_upload_file_details(self.filepath) 
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

        for index, i in enumerate(self.file_upload):
            try:
                self.imppubpage.select_update("select_update_dropdown", index+3)
                self.LogScreenshot.fLogScreenshot(message=f"Filepath is {i[1]}",
                    pass_=True, log=True, screenshot=False)
                self.imppubpage.upload_file("add_file", i[0], i[1], "upload_button", "file_status_popup_text", "upload_table_rows")
                
                # Keep delete file in separate test case
                # self.imppubpage.delete_file("delete_file", "delete_file_popup", "file_status_popup_text", "upload_table_rows")
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    def test_del_extraction_template(self, extra):
            # Instantiate the logScreenshot class
            self.LogScreenshot = cLogScreenshot(self.driver, extra)
            # Creating object of loginpage class
            self.loginPage = LoginPage(self.driver, extra)
            # Creating object of liveslrpage class
            self.liveslrpage = LiveSLRPage(self.driver, extra)
            # Creating object of ImportPublicationPage class
            self.imppubpage = ImportPublicationPage(self.driver, extra)
            # Read extraction sheet values
            self.file_upload = self.imppubpage.get_upload_file_details(self.filepath) 
            
            self.loginPage.driver.get(self.baseURL)
            self.loginPage.complete_login(self.username, self.password)
            self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

            for i in self.file_upload:
                try:
                    self.imppubpage.delete_file("delete_file", "delete_file_popup", "file_status_popup_text", "upload_table_rows")
                except Exception:
                    self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                        pass_=False, log=True, screenshot=True)
                    raise Exception("Element Not Found")
