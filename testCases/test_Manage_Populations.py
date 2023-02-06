"""
Test will validate Manage Populations Page

"""

import os
import pytest
from Pages.Base import Base
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
    population_val = []
    edited_population_val = []

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

                self.population_val.append(added_pop)
                self.LogScreenshot.fLogScreenshot(message=f"Added populations are {self.population_val}",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
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

        result = [(pop_list[i], self.population_val[i]) for i in range(0, len(pop_list))]

        for i in result:
            try:
                edited_pop = self.mngpoppage.edit_multiple_population(i[0], i[1], "edit_population", self.filepath, env)

                self.edited_population_val.append(edited_pop)
                self.LogScreenshot.fLogScreenshot(message=f"Edited populations are {self.edited_population_val}",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
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

        for i in self.edited_population_val:
            try:
                self.mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup",
                                                           "manage_pop_table_rows", env)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of Population validation is completed***",
                                          pass_=True, log=True, screenshot=False)
        