"""
Test will validate Manage Populations Page

"""

import os
import pandas as pd
import pytest
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.ManagePopulationsPage import ManagePopulationsPage

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ManagePopultionsPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getmanagepopdatafilepath()
    onco_population_val = []
    onco_edited_population_val = []
    non_onco_population_val = []
    non_onco_edited_population_val = []    

    @pytest.mark.C30244
    def test_add_population(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Addition of New Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1', 'pop2', 'pop3']

        for i in pop_list:
            try:
                added_pop = mngpoppage.add_multiple_population(i, "add_population_btn", self.filepath,
                                                               "manage_pop_table_rows", env)

                self.onco_population_val.append(added_pop)
                LogScreenshot.fLogScreenshot(message=f"Added populations are: {self.onco_population_val}",
                                             pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C30244
    def test_edit_population(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Editing the existing Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1', 'pop2', 'pop3']

        result = [(pop_list[i], self.onco_population_val[i]) for i in range(0, len(pop_list))]

        for i in result:
            try:
                edited_pop = mngpoppage.edit_multiple_population(i[0], i[1], "edit_population", self.filepath, env)

                self.onco_edited_population_val.append(edited_pop)
                LogScreenshot.fLogScreenshot(message=f"Edited populations are: {self.onco_edited_population_val}",
                                             pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C30244
    def test_delete_population(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Deletion of Existing Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        for i in self.onco_edited_population_val:
            try:
                mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup_cancel",
                                                      "manage_pop_table_rows", env)
                mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup_ok",
                                                      "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38392
    def test_add_non_oncology_population_ui_validation(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Field Level Error Messages while adding new Non-Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_managepopulationdata")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                mngpoppage.non_onocolgy_check_field_level_err_msg(i, "add_population_btn", filepath,
                                                                  "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38391
    def test_add_non_oncology_population(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Addition of New Non-Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_managepopulationdata")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                added_pop, tempalte_name = mngpoppage.non_onocolgy_add_population(i, "add_population_btn", filepath,
                                                                                  "manage_pop_table_rows", env)
                self.non_onco_population_val.append(added_pop)
                LogScreenshot.fLogScreenshot(message=f"Added Non-Oncology populations are: "
                                                     f"{self.non_onco_population_val}",
                                             pass_=True, log=True, screenshot=False)
                mngpoppage.non_onocolgy_add_duplicate_population(i, "add_population_btn", filepath,
                                                                    "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38391
    def test_edit_non_oncology_population(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Editing the existing Non-Oncology Population"        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_managepopulationdata")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        result = [(pop_list[i], self.non_onco_population_val[i]) for i in range(0, len(pop_list))]

        for i in result:
            try:
                edited_pop = mngpoppage.non_onocolgy_edit_population(i[0], i[1], "edit_population", filepath, env)

                self.non_onco_edited_population_val.append(edited_pop)
                LogScreenshot.fLogScreenshot(message=f"Edited Non-Oncology populations are: "
                                                     f"{self.non_onco_edited_population_val}",
                                             pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38391
    def test_delete_non_oncology_population(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Deletion of Existing Non-Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        for i in self.non_onco_edited_population_val:
            try:
                mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup_cancel",
                                                      "manage_pop_table_rows", env)
                mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup_ok",
                                                      "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38394
    @pytest.mark.C38857
    def test_validate_non_oncology_population_col_IDs(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Column IDs in downloaded Extraction Template after creating new Non-Oncology " \
                              "Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_managepopulationdata")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                added_pop, template_name = mngpoppage.non_onocolgy_add_population(i, "add_population_btn", filepath,
                                                                                  "manage_pop_table_rows", env)
                LogScreenshot.fLogScreenshot(message=f"Added Non-Oncology populations are {added_pop}",
                                             pass_=True, log=True, screenshot=False)
                mngpoppage.non_onocolgy_endpoint_details_validation(i, filepath, template_name)
                mngpoppage.non_oncology_extraction_file_col_name_validation(i, filepath, template_name)

                mngpoppage.delete_multiple_population(added_pop, "delete_population", "delete_population_popup_ok",
                                                      "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C34902
    def test_verify_new_col_managepopulation_page(self, extra, env, request, caseid):
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Verify Newly added columns 'Oncology/Non-Oncology', 'Custom Endpoints' in Manage " \
                              "Population page. "
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "managepopulation_additional_col_check")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1', 'pop2']

        for i in pop_list:
            try:
                mngpoppage.verify_new_col_managepopulation_page(i, filepath, env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C36193
    def test_non_oncology_edit_population_with_ep_categorical(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Editing the existing Non-Oncology Population with EP - Categorical"        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "edit_ep_categorical_invaliddata")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4', 'scenario5', 'scenario6']

        for i in scenarios:
            try:
                mngpoppage.\
                    non_onocolgy_edit_population_by_uploading_invalid_template(i,
                                                                               'Test_Non_Oncology_Pop_Categorical',
                                                                               'edit_population', filepath,
                                                                               'Categorical', env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C36194
    def test_non_oncology_edit_population_with_ep_continuous(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Editing the existing Non-Oncology Population with EP - Continuous"        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "edit_ep_continuous_invaliddata")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                mngpoppage.\
                    non_onocolgy_edit_population_by_uploading_invalid_template(i, 'Test_Non_Oncology_Pop_Continuous',
                                                                               'edit_population', filepath,
                                                                               'Continuous', env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C36243
    def test_non_oncology_edit_population_with_ep_timetoevent(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Editing the existing Non-Oncology Population with EP - TimetoEvent"        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "edit_ep_timetoevent_invaliddata")
        base.presence_of_admin_page_option("managepopulations_button", env)
        base.go_to_page("managepopulations_button", env)

        scenarios = ['scenario1', 'scenario2']

        for i in scenarios:
            try:
                mngpoppage.\
                    non_onocolgy_edit_population_by_uploading_invalid_template(i, 'Test_Non_Oncology_Pop_TimeToEvent',
                                                                               'edit_population', filepath,
                                                                               'TimetoEvent', env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
