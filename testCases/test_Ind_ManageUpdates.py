"""
Test will validate Manage Updates Page

"""

import pytest
from datetime import date, timedelta

from Pages.LoginPage import LoginPage
from Pages.ManageUpdatesPage import ManageUpdatesPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ManageUpdatesPage:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    added_updates_data = []
    edited_updates_data = []

    @pytest.mark.LIVEHTA_568
    def test_add_updates(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngupdpage = ManageUpdatesPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of Population Manageupdates validation is started***", pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y")
        self.day_val = today.day
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)

        for i in range(1, 4):
            try:
                self.mngupdpage.go_to_manageupdates("manageupdates_button")
                manage_update_data = self.mngupdpage.add_multiple_updates(i, "add_update_btn", self.day_val, "manage_update_table_rows", self.dateval)
                self.added_updates_data.append(manage_update_data)
                self.LogScreenshot.fLogScreenshot(message=f"Added population udpate is {self.added_updates_data}", pass_=True, log=True, screenshot=False)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of Population Manageupdates validation is completed***", pass_=True, log=True, screenshot=False)

    @pytest.mark.LIVEHTA_568
    def test_edit_updates(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngupdpage = ManageUpdatesPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Edit Population Manageupdates validation is started***", pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = (today + timedelta(1)).strftime("%m/%d/%Y")
        self.day_val = today.day
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)

        for i in self.added_updates_data:
            try:
                self.mngupdpage.go_to_manageupdates("manageupdates_button")
                manage_update_data = self.mngupdpage.edit_multiple_updates(i, "edit_updates", self.day_val, self.dateval)
                self.edited_updates_data.append(manage_update_data)
                self.LogScreenshot.fLogScreenshot(message=f"Edited population udpate is {self.edited_updates_data}", pass_=True, log=True, screenshot=False)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Edit Population Manageupdates validation is completed***", pass_=True, log=True, screenshot=False)

    @pytest.mark.LIVEHTA_568
    def test_delete_updates(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngupdpage = ManageUpdatesPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of Population Manageupdates validation is started***", pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y").replace('/', '')
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngupdpage.go_to_manageupdates("manageupdates_button")

        for i in self.edited_updates_data:
            try:
                self.mngupdpage.delete_multiple_manage_updates(i, "delete_updates", "delete_updates_popup", "manage_update_table_rows")

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of Population Manageupdates validation is completed***", pass_=True, log=True, screenshot=False)
