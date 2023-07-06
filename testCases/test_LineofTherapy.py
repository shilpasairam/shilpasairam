"""
Test will validate LineofTherapy Page

"""

import os
import pytest
from Pages.Base import Base
from Pages.ImportPublicationsPage import ImportPublicationPage
from Pages.LineOfTherapyPage import LineofTherapyPage

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_LineofTherapyPage:
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # added_lot_data = []
    # edited_lot_data = []

    @pytest.mark.C34623
    def test_oncology_lot(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getmanagelotdata(env)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of LineofTherapy Page class
        lotpage = LineofTherapyPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Manage Population filter 2 page functionalities for Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Addtion, Update and Deletion of Population filter 2 validation is "
                                             f"started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("managelot_button", env)

        pop_val = ['pop1', 'pop2']

        for i in pop_val:
            try:
                added_lot_data = lotpage.add_multiple_lot(i, "add_lot_btn", "managelot_table_rows", filepath, env)
                LogScreenshot.fLogScreenshot(message=f"Added Population filter 2 is {added_lot_data}",
                                             pass_=True, log=True, screenshot=False)
                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_success(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)

                edited_lot_data = lotpage.edit_multiple_lot(i, added_lot_data, "managelot_edit", filepath, env)
                LogScreenshot.fLogScreenshot(message=f"Edited Population filter 2 is {edited_lot_data}",
                                             pass_=True, log=True, screenshot=False)

                lotpage.delete_multiple_manage_lot(edited_lot_data, "managelot_delete", "managelot_delete_popup",
                                                   "managelot_table_rows", env)
                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Population filter 2 page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Addtion, Update and Deletion of Population filter 2 validation is "
                                             f"completed***",
                                     pass_=True, log=True, screenshot=False)
