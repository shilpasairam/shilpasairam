"""
Test will validate Manage Updates Page

"""

import os
import pytest
from datetime import date, timedelta
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from Pages.ManageUpdatesPage import ManageUpdatesPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ManageUpdatesPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    added_updates_data = []
    edited_updates_data = []

    @pytest.mark.smoketest_manageupdate
    def test_add_updates(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getmanageupdatesdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngupdpage = ManageUpdatesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Addition of Manage Update for selected population"

        LogScreenshot.fLogScreenshot(message=f"***Addtion of Population Manageupdates validation is started***",
                                          pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y")
        self.day_val = today.day
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manageupdates_button", env)

        pop_val = ['pop1', 'pop2']

        for i in pop_val:
            try:
                manage_update_data = mngupdpage.add_multiple_updates(i, filepath, "add_update_btn", self.day_val,
                                                                          "manage_update_table_rows", self.dateval, env)
                self.added_updates_data.append(manage_update_data)
                LogScreenshot.fLogScreenshot(message=f"Added population udpate is {self.added_updates_data}",
                                                  pass_=True, log=True, screenshot=False)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Addtion of Population Manageupdates validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.smoketest_manageupdate
    def test_edit_updates(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngupdpage = ManageUpdatesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Editing the existing Manage Update data for selected population"

        LogScreenshot.fLogScreenshot(message=f"***Edit Population Manageupdates validation is started***",
                                          pass_=True, log=True, screenshot=False)

        today = date.today()
        self.day_val = today.day
        # Manipulating the date values when values point to month end
        if self.day_val in [30, 31]:
            self.dateval = (today - timedelta(10)).strftime("%m/%d/%Y")
        else:
            self.dateval = (today + timedelta(1)).strftime("%m/%d/%Y")
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manageupdates_button", env)

        for i in self.added_updates_data:
            try:
                manage_update_data = mngupdpage.edit_multiple_updates(i, "edit_updates", self.day_val,
                                                                           self.dateval, env)
                self.edited_updates_data.append(manage_update_data)
                LogScreenshot.fLogScreenshot(message=f"Edited population udpate is {self.edited_updates_data}",
                                                  pass_=True, log=True, screenshot=False)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Edit Population Manageupdates validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.smoketest_manageupdate
    def test_delete_updates(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngupdpage = ManageUpdatesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Deletion of existing Manage Update data for selected population"

        LogScreenshot.fLogScreenshot(message=f"***Deletion of Population Manageupdates validation is started***",
                                          pass_=True, log=True, screenshot=False)

        today = date.today()
        self.dateval = today.strftime("%m/%d/%Y").replace('/', '')
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("manageupdates_button", env)

        for i in self.edited_updates_data:
            try:
                mngupdpage.delete_multiple_manage_updates(i, "delete_updates", "delete_updates_popup",
                                                               "manage_update_table_rows", env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Deletion of Population Manageupdates validation is "
                                                  f"completed***", pass_=True, log=True, screenshot=False)
