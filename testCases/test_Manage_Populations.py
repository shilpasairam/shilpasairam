"""
Test will validate Manage Populations Page

"""

import os
import pytest
from Pages.Base import Base
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
    population_val = []
    edited_population_val = []

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

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        LogScreenshot.fLogScreenshot(message=f"***Addition of Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1', 'pop2', 'pop3']

        for i in pop_list:
            try:
                added_pop = mngpoppage.add_multiple_population(i, "add_population_btn", self.filepath,
                                                                    "manage_pop_table_rows", env)

                self.population_val.append(added_pop)
                LogScreenshot.fLogScreenshot(message=f"Added populations are {self.population_val}",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
            
        LogScreenshot.fLogScreenshot(message=f"***Addition of Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)

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

        LogScreenshot.fLogScreenshot(message=f"***Edit the Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("managepopulations_button", env)

        pop_list = ['pop1', 'pop2', 'pop3']

        result = [(pop_list[i], self.population_val[i]) for i in range(0, len(pop_list))]

        for i in result:
            try:
                edited_pop = mngpoppage.edit_multiple_population(i[0], i[1], "edit_population", self.filepath, env)

                self.edited_population_val.append(edited_pop)
                LogScreenshot.fLogScreenshot(message=f"Edited populations are {self.edited_population_val}",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Edit the Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)

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

        LogScreenshot.fLogScreenshot(message=f"***Deletion of Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("managepopulations_button", env)

        for i in self.edited_population_val:
            try:
                mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup",
                                                           "manage_pop_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        LogScreenshot.fLogScreenshot(message=f"***Deletion of Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)
        