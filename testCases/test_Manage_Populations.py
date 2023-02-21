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
from Pages.OpenLiveSLRPage import LiveSLRPage
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
    def test_add_population(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Addition of Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1', 'pop2', 'pop3']

        for i in pop_list:
            try:
                added_pop = self.mngpoppage.add_multiple_population(i, "add_population_btn", self.filepath,
                                                                    "manage_pop_table_rows", env)

                self.onco_population_val.append(added_pop)
                self.LogScreenshot.fLogScreenshot(message=f"Added populations are {self.onco_population_val}",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
            
        self.LogScreenshot.fLogScreenshot(message=f"***Addition of Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C30244
    def test_edit_population(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Edit the Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1', 'pop2', 'pop3']

        result = [(pop_list[i], self.onco_population_val[i]) for i in range(0, len(pop_list))]

        for i in result:
            try:
                edited_pop = self.mngpoppage.edit_multiple_population(i[0], i[1], "edit_population", self.filepath, env)

                self.onco_edited_population_val.append(edited_pop)
                self.LogScreenshot.fLogScreenshot(message=f"Edited populations are {self.onco_edited_population_val}",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Edit the Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C30244
    def test_delete_population(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("managepopulations_button", env)

        for i in self.onco_edited_population_val:
            try:
                self.mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup",
                                                           "manage_pop_table_rows", env)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C38392
    def test_add_non_oncology_population_ui_validation(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)                
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                mngpoppage.non_onocolgy_check_field_level_err_msg(i, "add_population_btn", filepath[0],
                                                                    "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38391
    def test_add_non_oncology_population(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)                
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                added_pop, tempalte_name = mngpoppage.non_onocolgy_add_population(i, "add_population_btn", filepath[0],
                                                                    "manage_pop_table_rows", env)
                self.non_onco_population_val.append(added_pop)
                LogScreenshot.fLogScreenshot(message=f"Added Non-Oncology populations are {self.non_onco_population_val}",
                                                  pass_=True, log=True, screenshot=False)
                mngpoppage.non_onocolgy_add_duplicate_population(i, "add_population_btn", filepath[0],
                                                                    "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38391
    def test_edit_non_oncology_population(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        result = [(pop_list[i], self.non_onco_population_val[i]) for i in range(0, len(pop_list))]

        for i in result:
            try:
                edited_pop = mngpoppage.non_onocolgy_edit_population(i[0], i[1], "edit_population", filepath[0], env)

                self.non_onco_edited_population_val.append(edited_pop)
                LogScreenshot.fLogScreenshot(message=f"Edited Non-Oncology populations are {self.non_onco_edited_population_val}",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C38391
    def test_delete_non_oncology_population(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))        

        LogScreenshot.fLogScreenshot(message=f"***Deletion of Non-Oncology Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("managepopulations_button", env)

        for i in self.non_onco_edited_population_val:
            try:
                mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup",
                                                           "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Deletion of Non-Oncology Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C38394
    def test_validate_non_oncology_population_ep_details(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)                
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))        
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                added_pop, template_name = mngpoppage.non_onocolgy_add_population(i, "add_population_btn", filepath[0],
                                                                    "manage_pop_table_rows", env)
                LogScreenshot.fLogScreenshot(message=f"Added Non-Oncology populations are {added_pop}",
                                                  pass_=True, log=True, screenshot=False)
                mngpoppage.non_onocolgy_endpoint_details_validation(i, filepath[0], template_name)

                mngpoppage.delete_multiple_population(added_pop, "delete_population", "delete_population_popup",
                                            "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Populations page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
