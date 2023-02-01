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
    def test_add_lot(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getmanagelotdata(env)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of LineofTherapy Page class
        self.lotpage = LineofTherapyPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of New Line of Therapy validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", env)

        pop_val = ['pop1', 'pop2']

        for i in pop_val:
            try:
                self.base.go_to_page("managelot_button", env)
                manage_lot_data = self.lotpage.add_multiple_lot(i, "add_lot_btn", "managelot_table_rows", filepath, env)
                self.added_lot_data.append(manage_lot_data)
                self.LogScreenshot.fLogScreenshot(message=f"Added Line of Therapy is {self.added_lot_data}",
                                                  pass_=True, log=True, screenshot=False)
                self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn", env)
                self.imppubpage.upload_file_with_success(i, filepath, env)
                self.imppubpage.delete_file(i, filepath, "file_status_popup_text",
                                            "upload_table_rows", env)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Line of Therapy page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of New Line of Therapy validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C34623
    def test_edit_lot(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getmanagelotdata(env)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)         
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of LineofTherapy Page class
        self.lotpage = LineofTherapyPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Edit Line of Therapy validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", env)
        self.base.go_to_page("managelot_button", env)

        pop_val = ['pop1', 'pop2']

        result = [(pop_val[i], self.added_lot_data[i]) for i in range(0, len(pop_val))]

        for i in result:
            try:
                manage_lot_data = self.lotpage.edit_multiple_lot(i[0], i[1], "managelot_edit", filepath, env)
                self.edited_lot_data.append(manage_lot_data)
                self.LogScreenshot.fLogScreenshot(message=f"Edited Line of Therapy is {self.edited_lot_data}",
                                                  pass_=True, log=True, screenshot=False)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Line of Therapy page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Edit Line of Therapy validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C34623
    def test_del_lot(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getmanagelotdata(env)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of LineofTherapy Page class
        self.lotpage = LineofTherapyPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of New Line of Therapy validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", env)

        pop_val = ['pop1', 'pop2']

        result = [(pop_val[i], self.edited_lot_data[i]) for i in range(0, len(pop_val))]

        for i in result:
            try:
                self.base.go_to_page("managelot_button", env)
                self.lotpage.delete_multiple_manage_lot(i[1], "managelot_delete", "managelot_delete_popup",
                                                            "managelot_table_rows", env)
                self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn", env)
                self.imppubpage.upload_file_with_errors(i[0], filepath, env)
                self.imppubpage.delete_file(i[0], filepath, "file_status_popup_text",
                                            "upload_table_rows", env)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Line of Therapy page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of New Line of Therapy validation is completed***",
                                          pass_=True, log=True, screenshot=False)
                                          