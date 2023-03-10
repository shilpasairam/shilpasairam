import os
import random
import time

import pandas as pd
import pytest
from pytest_html.extras import extra
from selenium.webdriver.support import expected_conditions as ec

from Pages.Base import Base
from Pages.LiveNMAPage import LiveNMA
from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.readProperties import ReadConfig, config


@pytest.mark.usefixtures("init_driver")
class Test_LiveNMA:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getnmatestdata()

    @pytest.mark.smoketest_nma
    def test_live_nma_datatable_OS_Switch(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getnmatestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        nma = LiveNMA(self.driver, extra)
        # Read population data values
        pop_list = liveslrpage.get_population_data(filepath)
        # Read slrtype data values
        slrtype = liveslrpage.get_slrtype_data(filepath)
        # Read Studydesign data values
        std_data, std_data_chkbox = liveslrpage.get_study_design(filepath)
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = liveslrpage.get_reported_variables(filepath)

        request.node._tcid = caseid
        request.node._title = "Validate of LiveNMA Datatable with OS Switch Tab"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                base.presence_of_all_elements("slr_pop_panel_eles", env)
                nma.select_data(i[0], i[1], env)
                for j in slrtype:
                    base.presence_of_all_elements("slr_type_panel_eles", env)
                    nma.select_data(j[0], j[1], env)
                    base.presence_of_element("study_design_section", env)
                    nma.select_sub_section(std_data[0], std_data_chkbox[0], env, "study_design_section")
                    nma.select_sub_section(std_data[1], std_data_chkbox[1], env, "study_design_section")
                    nma.select_sub_section(rpt_data[0], rpt_data_chkbox[0], env, "reported_variable_section")
                    nma.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    nma.launch_nma("launch_live_nma", env, UnivWaitFor=5)
                    nma.table_display_check("live_nma_switch_1", "live_nma_data_table", env)
                    nma.validate_nma_selected_criteria_val(filepath, i[0], "live_nma_study_design",
                                                           "live_nma_reported_variable", "live_nma_pop_data", env)
                    nma.form_fill("add_study_button", filepath, "add_button", "live_nma_data_table_rows",
                                  "show_network", env)
                    time.sleep(3)
                    nma.driver.close()
                    nma.driver.switch_to.window(self.driver.window_handles[1])
            except Exception:
                raise Exception("LiveNMA Page action failed")

    @pytest.mark.smoketest_nma
    def test_live_nma_datatable_PFS_Switch(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getnmatestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        nma = LiveNMA(self.driver, extra)
        # Read population data values
        pop_list = liveslrpage.get_population_data(filepath)
        # Read slrtype data values
        slrtype = liveslrpage.get_slrtype_data(filepath)
        # Read Studydesign data values
        std_data, std_data_chkbox = liveslrpage.get_study_design(filepath)
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = liveslrpage.get_reported_variables(filepath)

        request.node._tcid = caseid
        request.node._title = "Validate of LiveNMA Datatable with PFS Switch Tab"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                base.presence_of_all_elements("slr_pop_panel_eles", env)
                nma.select_data(i[0], i[1], env)
                for j in slrtype:
                    base.presence_of_all_elements("slr_type_panel_eles", env)
                    nma.select_data(j[0], j[1], env)
                    base.presence_of_element("study_design_section", env)
                    nma.select_sub_section(std_data[0], std_data_chkbox[0], env, "study_design_section")
                    nma.select_sub_section(std_data[1], std_data_chkbox[1], env, "study_design_section")
                    nma.select_sub_section(rpt_data[0], rpt_data_chkbox[0], env, "reported_variable_section")
                    nma.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    nma.launch_nma("launch_live_nma", env, UnivWaitFor=5)
                    nma.table_display_check("live_nma_switch_2", "live_nma_data_table", env)
                    nma.validate_nma_selected_criteria_val(filepath, i[0], "live_nma_study_design",
                                                           "live_nma_reported_variable", "live_nma_pop_data", env)
                    nma.form_fill("add_study_button", filepath, "add_button", "live_nma_data_table_rows",
                                  "show_network", env)
                    time.sleep(3)
                    nma.driver.close()
                    nma.driver.switch_to.window(self.driver.window_handles[1])
            except Exception:
                raise Exception("LiveNMA Page action failed")
