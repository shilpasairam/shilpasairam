"""
Test will validate Manage Updates Page

"""

import os
import pytest
from datetime import date, timedelta
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase

from Pages.LoginPage import LoginPage
from Pages.ManageUpdatesPage import ManageUpdatesPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ManageUpdatesPage:
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    added_updates_data = []
    edited_updates_data = []
    non_onco_added_updates_data = []
    non_onco_edited_updates_data = []    

    @pytest.mark.C38951
    def test_oncology_manage_updates(self, extra, env, request, caseid):
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
        request.node._title = "Validate Manage Update page functionality for selected Oncology population"

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Oncology Population Manageupdates "
                                             f"validation is started***",
                                     pass_=True, log=True, screenshot=False)

        today = date.today()
        dateval = today.strftime("%m/%d/%Y")
        day_val = today.day
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("manageupdates_button", env)
        base.go_to_page("manageupdates_button", env)

        pop_val = ['pop1', 'pop2']

        for i in pop_val:
            try:
                manage_update_data = mngupdpage.add_multiple_updates(i, filepath, "add_update_btn", day_val,
                                                                     "manage_update_table_rows", dateval, env)
                # self.added_updates_data.append(manage_update_data)
                LogScreenshot.fLogScreenshot(message=f"Added population udpate is {manage_update_data}",
                                             pass_=True, log=True, screenshot=False)

                mngupdpage.delete_multiple_manage_updates(manage_update_data, "delete_updates", "delete_updates_popup",
                                                          "manage_update_table_rows", env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Oncology Population Manageupdates "
                                             f"validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    '''
    Commenting this method due to LIVEHTA-2618 implementation
    @pytest.mark.C38951
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
        request.node._title = "Validate Edit the existing Manage Update data for selected Oncology population"

        LogScreenshot.fLogScreenshot(message=f"***Edit Population Manageupdates validation is started***",
                                     pass_=True, log=True, screenshot=False)

        today = date.today()
        day_val = today.day
        # Manipulating the date values when values point to month end
        if day_val in [30, 31]:
            dateval = (today - timedelta(10)).strftime("%m/%d/%Y")
        else:
            dateval = (today + timedelta(1)).strftime("%m/%d/%Y")
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("manageupdates_button", env)
        base.go_to_page("manageupdates_button", env)

        for i in self.added_updates_data:
            try:
                manage_update_data = mngupdpage.edit_multiple_updates(i, "edit_updates", day_val,
                                                                      dateval, "Oncology", env)
                self.edited_updates_data.append(manage_update_data)
                LogScreenshot.fLogScreenshot(message=f"Edited population udpate is {self.edited_updates_data}",
                                             pass_=True, log=True, screenshot=False)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Edit Population Manageupdates validation is completed***",
                                     pass_=True, log=True, screenshot=False)
    '''

    @pytest.mark.C38950
    def test_nononcology_manage_updates(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngupdpage = ManageUpdatesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Manage Update page functionality for selected Non-Oncology population"

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Non-Oncology Population Manageupdates "
                                             f"validation is started***",
                                     pass_=True, log=True, screenshot=False)

        today = date.today()
        dateval = today.strftime("%m/%d/%Y")
        day_val = today.day
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_manageupdatesdata")
        base.presence_of_admin_page_option("manageupdates_button", env)
        base.go_to_page("manageupdates_button", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                manage_update_data = mngupdpage.add_multiple_updates(i, filepath, "add_update_btn", day_val,
                                                                     "manage_update_table_rows", dateval, env)
                # self.non_onco_added_updates_data.append(manage_update_data)
                LogScreenshot.fLogScreenshot(message=f"Added Non-Oncology population udpate is "
                                                     f"{manage_update_data}",
                                             pass_=True, log=True, screenshot=False)

                mngupdpage.delete_multiple_manage_updates(manage_update_data, "delete_updates", "delete_updates_popup",
                                                          "manage_update_table_rows", env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Non-Oncology Population Manageupdates "
                                             f"validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    '''
    Commenting this method due to LIVEHTA-2618 implementation
    @pytest.mark.C38950
    def test_nononcology_edit_updates(self, extra, env, request, caseid):
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
        request.node._title = "Validate Edit the existing Manage Update data for selected Non-Oncology population"

        LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology Population Manageupdates validation is started***",
                                     pass_=True, log=True, screenshot=False)

        today = date.today()
        day_val = today.day
        # Manipulating the date values when values point to month end
        if day_val in [30, 31]:
            dateval = (today - timedelta(10)).strftime("%m/%d/%Y")
        else:
            dateval = (today + timedelta(1)).strftime("%m/%d/%Y")
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("manageupdates_button", env)
        base.go_to_page("manageupdates_button", env)

        for i in self.non_onco_added_updates_data:
            try:
                manage_update_data = mngupdpage.edit_multiple_updates(i, "edit_updates", day_val,
                                                                      dateval, "Non-oncology", env)
                self.non_onco_edited_updates_data.append(manage_update_data)
                LogScreenshot.fLogScreenshot(message=f"Edited Non-Oncology population udpate is "
                                                     f"{self.non_onco_edited_updates_data}",
                                             pass_=True, log=True, screenshot=False)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Updates page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology Population Manageupdates validation is "
                                             f"completed***",
                                     pass_=True, log=True, screenshot=False)
    '''
