import os
import re
import time

import pandas as pd
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select

from Pages.Base import Base, fWaitFor
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec


class LiveNMA(Base):
    
    """Constructor of the LiveNMA class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 20)

    def select_data(self, locator, locator_button, env):
        if self.isselected(locator_button, env):
            self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.jsclick(locator, env, UnivWaitFor=10)
            if self.isselected(locator_button, env):
                self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                                  pass_=True, log=True, screenshot=True)

    def select_sub_section(self, locator, locator_button, env, scroll=None):
        if self.scroll(scroll, env):
            if self.isselected(locator_button, env):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.jsclick(locator, env, UnivWaitFor=10)
                if self.isselected(locator_button, env):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=False)
            self.scrollback("SLR_page_header", env)

    def select_all_sub_section(self, locator, locator_button, env, scroll=None):
        if self.scroll(scroll, env):
            if self.isselected(locator_button, env):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.jsclick(locator, env, UnivWaitFor=10)
                if self.isselected(locator_button, env):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=False)
            self.scrollback("SLR_page_header", env)

    def table_display_check(self, nma_data_loc, locator, env):
        self.jsclick(nma_data_loc, env, UnivWaitFor=10)
        if self.isdisplayed(locator, env, UnivWaitFor=20):
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is displayed",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.driver.find_element(getattr(By, self.locatortype(locator, env)), self.locatorpath(locator, env))\
                .is_displayed()
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is displayed with extra wait time",
                                              pass_=True, log=True, screenshot=True)

    def get_nma_selected_criteria_values(self, locator_study, locator_var, env):
        ele1 = self.select_elements(locator_study, env)
        ele2 = self.select_elements(locator_var, env)
        return ele1, ele2

    def validate_nma_selected_criteria_val(self, filepath, pop, locator_study, locator_var, locator_pop, env):
        # Read reportedvariables and studydesign expected data values
        design_val, var_val = self.liveslrpage.get_data_values(filepath)
        # Get the actual values
        act_study_design, act_rep_var = self.get_nma_selected_criteria_values(locator_study, locator_var, env)
        if pop == self.get_text(locator_pop, env):
            self.LogScreenshot.fLogScreenshot(
                message=f"Correct Population is displayed in NMA Page:{self.get_text(locator_pop, env)}",
                pass_=True, log=True, screenshot=False)
        for k in act_study_design:
            if k.text in design_val:
                self.LogScreenshot.fLogScreenshot(
                    message=f"Correct StudyDesign Value is displayed in NMA Page:{k.text}",
                    pass_=True, log=True, screenshot=False)
        for v in act_rep_var:
            if v.text in var_val:
                self.LogScreenshot.fLogScreenshot(
                    message=f"Correct Reported Variable is displayed in NMA Page:{v.text}",
                    pass_=True, log=True, screenshot=False)

    @fWaitFor
    def launch_nma(self, locator, env, UnivWaitFor=0):
        self.click("NMA_Button", env)
        if self.clickable(locator, env):
            self.jsclick(locator, env)
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is clickable",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is not clickable",
                                              pass_=False, log=True, screenshot=False)
        # Switch the driver to LiveNMA tab
        self.driver.switch_to.window(self.driver.window_handles[2])
        try:
            self.presence_of_element("live_nma_data_table", env)
            self.assertPageTitle("Live NMA", UnivWaitFor=30)
            self.LogScreenshot.fLogScreenshot(message=f"LiveNMA Page Opened successfully",
                                              pass_=True, log=True, screenshot=True)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"LiveNMA Page load is not successful. Please try againa",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Login Unsuccessful")

    def form_fill(self, locator, filepath, add_button, trows_locator, network_loc, env):
        # Fetching total rows count before adding study data
        table_rows_before = self.select_elements(trows_locator, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new study: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)
        self.click(locator, env)

        study_values, expected_val = self.liveslrpage.get_addstudy_data(filepath)

        for i in study_values:
            self.input_text(i[0], i[1], env, UnivWaitFor=10)

        ele = self.select_element("reference_dropdown", env)
        select = Select(ele)
        select.select_by_index(1)
        # Adding dropdown value to expected values list as we are selecting the dropdown with index
        expected_val.append(select.first_selected_option.text)
        self.LogScreenshot.fLogScreenshot(message=f"Expected Study values are: {expected_val}",
                                          pass_=True, log=True, screenshot=False)
        
        self.LogScreenshot.fLogScreenshot(message=f"Entered User Study Values are : ",
                                          pass_=True, log=True, screenshot=True)  
        self.click(add_button, env)

        table_rows_after = self.select_elements(trows_locator, env)
        self.LogScreenshot.fLogScreenshot(message=f"Table length after adding a new study: {len(table_rows_after)}",
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                result = []
                td1 = self.select_elements('live_nma_table_row_1', env)
                td2 = self.select_elements('live_nma_table_row_2', env)
                for m in td1:
                    result.append(m.text)

                for n in td2:
                    result.append(n.text)

                self.LogScreenshot.fLogScreenshot(message=f'Actual Study values are: {result}',
                                                  pass_=True, log=True, screenshot=False)

                # Comparison between Expected Study values with Actual Study values
                for values in expected_val:
                    if str(values) in result:
                        self.LogScreenshot.fLogScreenshot(message=f'Correct values displayed on top of table: {values}',
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        raise Exception("Wrong values entered")

                if self.clickable(network_loc, env):
                    self.LogScreenshot.fLogScreenshot(message=f"Show Network Button is clickable",
                                                      pass_=True, log=True, screenshot=False)
                    self.click(network_loc, env)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Show Network Button is not clickable",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception("Show Network Button is not clickable")
                
                self.presence_of_element("network_section", env)
                if self.isdisplayed("network_section", env, UnivWaitFor=30):
                    self.LogScreenshot.fLogScreenshot(message=f"'Network Section' is displayed",
                                              pass_=True, log=True, screenshot=True)
                    
                    panel_ele = self.select_elements("network_panel_txt", env)
                    panel_txt = [m.text for m in panel_ele]
                    self.LogScreenshot.fLogScreenshot(message=f"Panel text from Network section is '{panel_txt}'",
                                              pass_=True, log=True, screenshot=True)

                    if self.clickable("run_nma", env):
                        self.LogScreenshot.fLogScreenshot(message=f"Run NMA Button is clickable",
                                                        pass_=True, log=True, screenshot=False)
                        self.click("run_nma", env)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"Run NMA Button is not clickable",
                                                        pass_=False, log=True, screenshot=False)
                        raise Exception("Run NMA Button is not clickable")
                    
                    if self.isdisplayed("result_forestplot", env, UnivWaitFor=30):
                        self.LogScreenshot.fLogScreenshot(message=f"'Forest plot' is displayed in Result section.",
                                                    pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"'Forest plot' is not displayed in Result section.",
                                                    pass_=False, log=True, screenshot=True)
                        raise Exception("'Forest plot' is not displayed in Result section.")

                    if self.isdisplayed("result_userstudy", env, UnivWaitFor=30):
                        self.LogScreenshot.fLogScreenshot(message=f"'User Study' is displayed in Result section.",
                                                    pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"'User Study' is not displayed in Result section.",
                                                    pass_=False, log=True, screenshot=True)
                        raise Exception("'User Study' is not displayed in Result section.")

                    if self.isdisplayed("result_network", env, UnivWaitFor=30):
                        self.LogScreenshot.fLogScreenshot(message=f"'Network' is displayed in Result section.",
                                                    pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"'Network' is not displayed in Result section.",
                                                    pass_=False, log=True, screenshot=True)
                        raise Exception("'Network' is not displayed in Result section.")

                    if self.clickable("print_screen", env):
                        self.LogScreenshot.fLogScreenshot(message=f"Print Screen Button is clickable",
                                                        pass_=True, log=True, screenshot=False)
                        self.scroll("print_screen", env)
                        self.LogScreenshot.fLogScreenshot(message=f"Print Screen Details : ",
                                                        pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"Print Screen Button is not clickable",
                                                        pass_=False, log=True, screenshot=False)
                        raise Exception("Print Screen Button is not clickable")
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"'Network Section' is not displayed",
                                              pass_=False, log=True, screenshot=True)
                    raise Exception("'Network Section' is not displayed")
        except Exception:
            raise Exception("Failed in adding new study data to table")
