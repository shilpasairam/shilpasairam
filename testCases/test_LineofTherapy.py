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
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getmanagelotdata()
    added_lot_data = []
    edited_lot_data = []

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

        LogScreenshot.fLogScreenshot(message=f"***Addtion, Update and Deletion of Population filter 2 validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("managelot_button", env)

        pop_val = ['pop1', 'pop2']

        for i in pop_val:
            try:
                # base.go_to_page("managelot_button", env)
                added_lot_data = lotpage.add_multiple_lot(i, "add_lot_btn", "managelot_table_rows", filepath, env)
                # self.added_lot_data.append(manage_lot_data)
                LogScreenshot.fLogScreenshot(message=f"Added Population filter 2 is {added_lot_data}",
                                             pass_=True, log=True, screenshot=False)
                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_success(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)

                edited_lot_data = lotpage.edit_multiple_lot(i, added_lot_data, "managelot_edit", filepath, env)
                # self.edited_lot_data.append(manage_lot_data)
                LogScreenshot.fLogScreenshot(message=f"Edited Population filter 2 is {edited_lot_data}",
                                             pass_=True, log=True, screenshot=False)

                # base.go_to_page("managelot_button", env)
                lotpage.delete_multiple_manage_lot(edited_lot_data, "managelot_delete", "managelot_delete_popup",
                                                   "managelot_table_rows", env)
                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Population filter 2 page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Addtion, Update and Deletion of Population filter 2 validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    # @pytest.mark.C34623
    # def test_edit_lot(self, extra, env, request, caseid):
    #     baseURL = ReadConfig.getPortalURL(env)
    #     filepath = ReadConfig.getmanagelotdata(env)
    #     # Instantiate the logScreenshot class
    #     LogScreenshot = cLogScreenshot(self.driver, extra)
    #     # Instantiate the Base class
    #     base = Base(self.driver, extra)         
    #     # Creating object of loginpage class
    #     loginPage = LoginPage(self.driver, extra)
    #     # Creating object of LineofTherapy Page class
    #     lotpage = LineofTherapyPage(self.driver, extra)

    #     request.node._tcid = caseid
    #     request.node._title = "Validate Editing the Population filter 2 data"

    #     LogScreenshot.fLogScreenshot(message=f"***Edit Population filter 2 validation is started***",
    #                                  pass_=True, log=True, screenshot=False)
        
    #     loginPage.driver.get(baseURL)
    #     loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
    #     base.presence_of_admin_page_option("managelot_button", env)
    #     base.go_to_page("managelot_button", env)

    #     pop_val = ['pop1', 'pop2']

    #     result = [(pop_val[i], self.added_lot_data[i]) for i in range(0, len(pop_val))]

    #     for i in result:
    #         try:
    #             manage_lot_data = lotpage.edit_multiple_lot(i[0], i[1], "managelot_edit", filepath, env)
    #             self.edited_lot_data.append(manage_lot_data)
    #             LogScreenshot.fLogScreenshot(message=f"Edited Population filter 2 is {self.edited_lot_data}",
    #                                          pass_=True, log=True, screenshot=False)

    #         except Exception:
    #             LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Population filter 2 page",
    #                                          pass_=False, log=True, screenshot=True)
    #             raise Exception("Element Not Found")

    #     LogScreenshot.fLogScreenshot(message=f"***Edit Population filter 2 validation is completed***",
    #                                  pass_=True, log=True, screenshot=False)

    # @pytest.mark.C34623
    # def test_del_lot(self, extra, env, request, caseid):
    #     baseURL = ReadConfig.getPortalURL(env)
    #     filepath = ReadConfig.getmanagelotdata(env)
    #     # Instantiate the logScreenshot class
    #     LogScreenshot = cLogScreenshot(self.driver, extra)
    #     # Instantiate the Base class
    #     base = Base(self.driver, extra)        
    #     # Creating object of loginpage class
    #     loginPage = LoginPage(self.driver, extra)
    #     # Creating object of LineofTherapy Page class
    #     lotpage = LineofTherapyPage(self.driver, extra)
    #     # Creating object of ImportPublicationPage class
    #     imppubpage = ImportPublicationPage(self.driver, extra)

    #     request.node._tcid = caseid
    #     request.node._title = "Validate Deletion of Population filter 2"

    #     LogScreenshot.fLogScreenshot(message=f"***Deletion of New Population filter 2 validation is started***",
    #                                  pass_=True, log=True, screenshot=False)
        
    #     loginPage.driver.get(baseURL)
    #     loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
    #     base.presence_of_admin_page_option("managelot_button", env)

    #     pop_val = ['pop1', 'pop2']

    #     result = [(pop_val[i], self.edited_lot_data[i]) for i in range(0, len(pop_val))]

    #     for i in result:
    #         try:
    #             base.go_to_page("managelot_button", env)
    #             lotpage.delete_multiple_manage_lot(i[1], "managelot_delete", "managelot_delete_popup",
    #                                                "managelot_table_rows", env)
    #             base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
    #             imppubpage.upload_file_with_errors(i[0], filepath, env)
    #             imppubpage.delete_file(i[0], filepath, "file_status_popup_text", "upload_table_rows", env)

    #         except Exception:
    #             LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Population filter 2 page",
    #                                          pass_=False, log=True, screenshot=True)
    #             raise Exception("Element Not Found")

    #     LogScreenshot.fLogScreenshot(message=f"***Deletion of New Population filter 2 validation is completed***",
    #                                  pass_=True, log=True, screenshot=False)
