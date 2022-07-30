"""
Test will validate Manage Populations Page

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
class Test_ManagePopultionsPage:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getmanagepopdatafilepath()
    population_val = []

    def test_upload_extraction_template(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)
        # Read extraction sheet values
        self.file_upload = self.mngpoppage.get_template_file_details(self.filepath)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngpoppage.go_to_managepopulations("managepopulations_button")

        count = 1
        for i in self.file_upload:
            try:
                added_pop = self.mngpoppage.add_multiple_population(count, "add_population_btn", self.filepath, "template_file_upload", i[1], "manage_pop_table_rows")

                self.population_val.append(added_pop)
                self.LogScreenshot.fLogScreenshot(message=f"Added populations are {self.population_val}", pass_=True, log=True, screenshot=False)
                # self.mngpoppage.delete_multiple_population(count, "delete_population", self.filepath, "delete_population_popup", "manage_pop_table_rows")

                count += 1

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    def test_delete_population(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)
        # Read extraction sheet values
        self.file_upload = self.mngpoppage.get_template_file_details(self.filepath)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.mngpoppage.go_to_managepopulations("managepopulations_button")

        for i in self.population_val:
            try:
                self.mngpoppage.delete_multiple_population(i, "delete_population", "delete_population_popup", "manage_pop_table_rows")

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        