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
    """Constructor of the LiveSLR Page class"""

    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
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

    def select_data(self, locator, locator_button):
        if self.isselected(locator_button):
            self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.jsclick(locator, UnivWaitFor=10)
            if self.isselected(locator_button):
                self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                                  pass_=True, log=True, screenshot=True)

    def select_sub_section(self, locator, locator_button, scroll=None):
        if self.scroll(scroll):
            if self.isselected(locator_button):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.jsclick(locator, UnivWaitFor=10)
                if self.isselected(locator_button):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=False)
            self.scrollback()

    def select_all_sub_section(self, locator, locator_button, scroll=None):
        if self.scroll(scroll):
            if self.isselected(locator_button):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.jsclick(locator, UnivWaitFor=10)
                if self.isselected(locator_button):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=False)
            self.scrollback()

    def table_display_check(self, nma_data_loc, locator):
        self.jsclick(nma_data_loc, UnivWaitFor=10)
        if self.isdisplayed(locator, UnivWaitFor=20):
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is displayed",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).is_displayed()
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is displayed with extra wait time",
                                              pass_=True, log=True, screenshot=True)

    def get_nma_selected_criteria_values(self, locator_study, locator_var):
        ele1 = self.select_elements(locator_study)
        ele2 = self.select_elements(locator_var)
        return ele1, ele2

    def validate_nma_selected_criteria_val(self, filepath, pop, locator_study, locator_var, locator_pop):
        # Read reportedvariables and studydesign expected data values
        design_val, var_val = self.liveslrpage.get_data_values(filepath)
        # Get the actual values
        act_study_design, act_rep_var = self.get_nma_selected_criteria_values(locator_study, locator_var)
        if pop == self.get_text(locator_pop):
            self.LogScreenshot.fLogScreenshot(
                message=f"Correct Population is displayed in NMA Page:{self.get_text(locator_pop)}",
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
    def launch_nma(self, locator, UnivWaitFor=0):
        self.click("NMA_Button")
        if self.clickable(locator):
            self.jsclick(locator)
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is clickable",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is not clickable",
                                              pass_=False, log=True, screenshot=False)
        # Switch the driver to LiveNMA tab
        self.driver.switch_to.window(self.driver.window_handles[2])
        try:
            self.assertPageTitle("Live NMA", UnivWaitFor=30)
            self.LogScreenshot.fLogScreenshot(message=f"LiveNMA Page Opened successfully",
                                              pass_=True, log=True, screenshot=True)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"LiveNMA Page load is not successful. Please try againa",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Login Unsuccessful")

    def form_fill(self, locator, filepath, add_button, trows_locator, network_loc):
        # Fetching total rows count before adding study data
        table_rows_before = self.select_elements(trows_locator)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new study: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)
        self.click(locator)

        study_values, expected_val = self.liveslrpage.get_addstudy_data(filepath)
        self.LogScreenshot.fLogScreenshot(message=f'List values are: {study_values}',
                                          pass_=True, log=True, screenshot=False)
        for i in study_values:
            self.input_text(i[0], i[1], UnivWaitFor=10)
            # self.LogScreenshot.fLogScreenshot(message=f'Enter value for: {i[0]}',
            #                                   pass_=True, log=True, screenshot=False)

        ele = self.select_element("reference_dropdown")
        select = Select(ele)
        select.select_by_index(1)
        # Adding dropdown value to expected values list as we are selecting the dropdown with index
        expected_val.append(select.first_selected_option.text)
        self.LogScreenshot.fLogScreenshot(message=f'Expected values are: {expected_val}',
                                          pass_=True, log=True, screenshot=False)
        # select.select_by_visible_text("Lenalidomide")
        # self.LogScreenshot.fLogScreenshot(message=f'Dropdown value is: {ele.text}\n'
        #                                           f'Dropdown ele value is: {select.first_selected_option.text}',
        #                                   pass_=True, log=True, screenshot=False)
        self.click(add_button)
        # options_list = select.options
        # for i in options_list:
        #     self.LogScreenshot.fLogScreenshot(message=f'DropDown Value is {i.text}',
        #                                       pass_=True, log=True, screenshot=False)
        # Fetching total rows count after adding study data
        table_rows_after = self.select_elements(trows_locator)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new study: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                result = []
                td1 = self.select_elements('live_nma_table_row_1')
                td2 = self.select_elements('live_nma_table_row_2')
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

                if self.clickable(network_loc):
                    self.LogScreenshot.fLogScreenshot(message=f'Show Network Button is clickable', pass_=True, log=True, screenshot=False)
                    self.click(network_loc)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Show Network Button is not clickable', pass_=False, log=True, screenshot=False)
        except Exception:
            raise Exception("Failed in adding new study data to table")