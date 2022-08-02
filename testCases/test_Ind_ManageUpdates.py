"""
Test will validate Manage Updates Page

"""

import time
import pytest
from datetime import date
from Pages.ManagePopulationsPage import ManagePopulationsPage

from Pages.LoginPage import LoginPage
from Pages.ManageQADataPage import ManageQADataPage
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

    def test_add_updates(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngupdpage = ManageUpdatesPage(self.driver, extra)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y")  # .replace('/', '')
        self.day_val = today.day
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        # self.mngupdpage.go_to_manageupdates("manageupdates_button")

        for i in range(3):
            try:
                self.mngupdpage.go_to_manageupdates("manageupdates_button")
                self.LogScreenshot.fLogScreenshot(message=f"Date values are: {self.dateval} and {self.day_val}", pass_=True, log=True, screenshot=False)
                manage_update_data = self.mngupdpage.add_multiple_updates("add_update_btn", self.day_val, "manage_update_table_rows", self.dateval)
                self.added_updates_data.append(manage_update_data)
                self.LogScreenshot.fLogScreenshot(message=f"Added population udpate is {self.added_updates_data}", pass_=True, log=True, screenshot=False)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    def test_delete_updates(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngupdpage = ManageUpdatesPage(self.driver, extra)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y").replace('/', '')
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngupdpage.go_to_manageupdates("manageupdates_button")

        for i in self.added_updates_data:
            try:
                self.mngupdpage.delete_multiple_manage_updates(i, "delete_updates", "delete_updates_popup", "manage_update_table_rows")

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")