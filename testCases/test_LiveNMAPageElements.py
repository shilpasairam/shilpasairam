import os
import random
import time

import pandas as pd
import pytest
from pytest_html.extras import extra
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from webdriver_manager.core import driver

from Pages.Base import Base
from Pages.LiveNMAPage import LiveNMA
from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig, config
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_LiveNMA:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getnmatestdata()

    def test_live_nma_datatable_OS_Switch(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getnmatestdata(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)         
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of slrreport class
        self.nma = LiveNMA(self.driver, extra)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 20)
        # Read population data values
        self.pop_list = self.liveslrpage.get_population_data(filepath)
        # Read slrtype data values
        self.slrtype = self.liveslrpage.get_slrtype_data(filepath)
        # Read Studydesign data values
        self.std_data, self.std_data_chkbox = self.liveslrpage.get_study_design(filepath)
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.liveslrpage.get_reported_variables(filepath)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             print(file)
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    print(file)
                    os.remove(os.path.join(root, file))

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("SLR_Homepage", env)
        for i in self.pop_list:
            try:
                self.base.presence_of_all_elements("slr_pop_panel_eles", env)
                self.nma.select_data(i[0], i[1], env)
                for j in self.slrtype:
                    self.base.presence_of_all_elements("slr_type_panel_eles", env)
                    self.nma.select_data(j[0], j[1], env)
                    self.base.presence_of_element("study_design_section", env)
                    self.nma.select_sub_section(self.std_data[0], self.std_data_chkbox[0], env,
                                                "study_design_section")
                    self.nma.select_sub_section(self.std_data[1], self.std_data_chkbox[1], env,
                                                "study_design_section")
                    self.nma.select_sub_section(self.rpt_data[0], self.rpt_data_chkbox[0], env,
                                                "reported_variable_section")
                    self.nma.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1], env,
                                                "reported_variable_section")
                    self.nma.launch_nma("launch_live_nma", env, UnivWaitFor=5)
                    self.nma.table_display_check("live_nma_switch_1", "live_nma_data_table", env)
                    self.nma.validate_nma_selected_criteria_val(filepath, i[0], "live_nma_study_design",
                                                                "live_nma_reported_variable", "live_nma_pop_data", env)
                    self.nma.form_fill("add_study_button", filepath, "add_button", "live_nma_data_table_rows",
                                       "show_network", env)
                    time.sleep(3)
                    self.nma.driver.close()
                    self.nma.driver.switch_to.window(self.driver.window_handles[1])
            except Exception:
                raise Exception("LiveNMA Page action failed")

    def test_live_nma_datatable_PFS_Switch(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getnmatestdata(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)         
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of slrreport class
        self.nma = LiveNMA(self.driver, extra)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 20)
        # Read population data values
        self.pop_list = self.liveslrpage.get_population_data(filepath)
        # Read slrtype data values
        self.slrtype = self.liveslrpage.get_slrtype_data(filepath)
        # Read Studydesign data values
        self.std_data, self.std_data_chkbox = self.liveslrpage.get_study_design(filepath)
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.liveslrpage.get_reported_variables(filepath)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             print(file)
        #             os.remove(os.path.join(root, file))

        # # Removing the files before the test runs
        # if os.path.exists(f'ActualOutputs'):
        #     for root, dirs, files in os.walk(f'ActualOutputs'):
        #         for file in files:
        #             print(file)
        #             os.remove(os.path.join(root, file))

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("SLR_Homepage", env)
        for i in self.pop_list:
            try:
                self.base.presence_of_all_elements("slr_pop_panel_eles", env)
                self.nma.select_data(i[0], i[1], env)
                for j in self.slrtype:
                    self.base.presence_of_all_elements("slr_type_panel_eles", env)
                    self.nma.select_data(j[0], j[1], env)
                    self.base.presence_of_element("study_design_section", env)
                    self.nma.select_sub_section(self.std_data[0], self.std_data_chkbox[0], env,
                                                "study_design_section")
                    self.nma.select_sub_section(self.std_data[1], self.std_data_chkbox[1], env,
                                                "study_design_section")
                    self.nma.select_sub_section(self.rpt_data[0], self.rpt_data_chkbox[0], env,
                                                "reported_variable_section")
                    self.nma.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1], env,
                                                "reported_variable_section")
                    self.nma.launch_nma("launch_live_nma", env, UnivWaitFor=5)
                    self.nma.table_display_check("live_nma_switch_2", "live_nma_data_table", env)
                    self.nma.validate_nma_selected_criteria_val(filepath, i[0], "live_nma_study_design",
                                                                "live_nma_reported_variable", "live_nma_pop_data", env)
                    self.nma.form_fill("add_study_button", filepath, "add_button", "live_nma_data_table_rows",
                                       "show_network", env)
                    time.sleep(3)
                    self.nma.driver.close()
                    self.nma.driver.switch_to.window(self.driver.window_handles[1])
            except Exception:
                raise Exception("LiveNMA Page action failed")
