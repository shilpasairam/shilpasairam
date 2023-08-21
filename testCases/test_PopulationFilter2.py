"""
Test will validate Population filter 2 Page

"""

import os
import pytest
from Pages.Base import Base
from Pages.ImportPublicationsPage import ImportPublicationPage
from Pages.PopulationFilter2Page import PopulationFilter2Page

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_LineofTherapyPage:
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C34623
    def test_oncology_popfilter2(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getmanagePopulationFilter2data(env)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of LineofTherapy Page class
        popfilter2page = PopulationFilter2Page(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Manage Population filter 2 page functionalities for Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Addtion, Update and Deletion of Population filter 2 validation is "
                                             f"started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("managepopfilter2_button", env)

        pop_val = ['pop1', 'pop2']

        for i in pop_val:
            try:
                added_popfilter2_data = popfilter2page.add_multiple_popfilter2(i, "add_popfilter2_btn", "managepopfilter2_table_rows", filepath, env)
                LogScreenshot.fLogScreenshot(message=f"Added Population filter 2 is {added_popfilter2_data}",
                                             pass_=True, log=True, screenshot=False)
                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_success(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)

                edited_popfilter2_data = popfilter2page.edit_multiple_popfilter2(i, added_popfilter2_data, "managepopfilter2_edit", filepath, env)
                LogScreenshot.fLogScreenshot(message=f"Edited Population filter 2 is {edited_popfilter2_data}",
                                             pass_=True, log=True, screenshot=False)

                popfilter2page.delete_multiple_manage_popfilter2(edited_popfilter2_data, "managepopfilter2_delete", "managepopfilter2_delete_popup",
                                                   "managepopfilter2_table_rows", env)
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
