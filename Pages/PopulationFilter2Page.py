import math
from re import T
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ManagePopulationsPage import ManagePopulationsPage
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot


class PopulationFilter2Page(Base):

    """Constructor of the PopulationFilter2 Page class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)
    
    # Reading Population data for PopulationFilter2 Page
    def get_updates_pop_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop = df.loc[df['Name'] == locatorname]['Population'].dropna().to_list()
        return pop

    # Reading Population data for PopulationFilter2 Page
    def get_popfilter2_name(self, filepath, locatorname, columnname):
        df = pd.read_excel(filepath)
        popfilter2_name = df.loc[df['Name'] == locatorname][columnname].dropna().to_list()
        return popfilter2_name
    
    def get_expected_data(self, filepath, columnname):
        file = pd.read_excel(filepath)
        popfilter2_options = list(file[columnname].dropna())
        return popfilter2_options    
    
    # Find the total row data if data is being ordered using Pagination
    def get_table_data(self, table_info, table_next_btn, table_rows_data, env):
        self.refreshpage()
        time.sleep(2)
        # get the count info and extract the total value        
        table_count_info = self.get_text(table_info, env)
        ind1 = table_count_info.index('of')
        ind2 = table_count_info.index('entries')
        total_entries = int(table_count_info[ind1+3:ind2-1])

        # Divide the total entries value with the number of records displayed in a page and round off the result to
        # next nearest integer value
        page_counter = math.ceil(total_entries/10)
        popfilter2_optns_text = self.get_texts(table_rows_data, env)
        
        # Iterate over the remaining pages and append the row data
        for i in range(1, page_counter):
            self.click(table_next_btn, env)
            time.sleep(1)
            options1 = self.select_elements(table_rows_data, env)
            for v in options1:
                popfilter2_optns_text.append(v.text)
        
        return popfilter2_optns_text
    
    def check_popfilter2_ui_elements(self, filepath, env):
        # Read Expected messages
        expected_data = self.get_expected_data(filepath, "Expected_ui_elements")

        pagedetails = [self.get_text("managepopfilter2_page_heading", env), self.get_text("managepopfilter2_pagecontent", env)]

        if expected_data == pagedetails:
            self.LogScreenshot.fLogScreenshot(message=f"Population filter 2 Page name and page content is present "
                                                      f"in UI",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Population filter 2 Page name and page content is not "
                                                      f"present in UI. Expected Page name: '{expected_data}'. "
                                                      f"Actual Page name: '{pagedetails}'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Population filter 2 Page name and page content is not present in UI")            

        if self.isdisplayed("add_popfilter2_btn", env) and self.isdisplayed("managepopfilter2_resultpanel_heading", env):
            self.LogScreenshot.fLogScreenshot(message=f"Add Population filter 2 button and Result panel heading is "
                                                      f"present UI", pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Add Population filter 2 button and Result panel heading is not "
                                                      f"present UI", pass_=False, log=True, screenshot=False)
            raise Exception(f"Add Population filter 2 button and Result panel heading is not present UI")
    
    def add_multiple_popfilter2(self, locatorname, add_popfilter2_button, table_rows, filepath, env):
        expected_add_status_text = "Population filter 2 added successfully"

        self.go_to_page("managepopfilter2_button", env)

        # Read PopulationFilter2 name details from data sheet
        popfilter2_name = self.get_popfilter2_name(filepath, locatorname, "PopulationFilter2_name")

        # Read Expected PopulationFilter2 options from data sheet
        expected_popfilter2_options = self.get_expected_data(filepath, 'Expected_popfilter2_options')    

        # Validate the UI page elements
        self.check_popfilter2_ui_elements(filepath, env)

        # Fetch the complete PopulationFilter2 data from the table
        complete_popfilter2_table_data = self.get_table_data("managepopfilter2_table_rows_info", "managepopfilter2_table_next_btn",
                                                      "managepopfilter2_table_rows_data", env)
        
        # Compare the expected mandatory popfilter2 option with actual table data
        for j in expected_popfilter2_options:
            if j in complete_popfilter2_table_data:
                self.LogScreenshot.fLogScreenshot(message=f"Mandatory PopulationFilter2 '{j}' option is present in the table.",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Mandatory PopulationFilter2 '{j}' option is absent in the table.",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Mandatory PopulationFilter2 '{j}' option is absent in the table.")            
        
        # Fetching total rows count before adding a new PopulationFilter2
        table_rows_before = self.mngpoppage.get_table_length("managepopfilter2_table_rows_info",
                                                             "managepopfilter2_table_next_btn", table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new PopulationFilter2: {table_rows_before}',
                                          pass_=True, log=True, screenshot=True)        
        
        self.scroll_and_click("managepopfilter2_page_heading", env)
        self.click(add_popfilter2_button, env, UnivWaitFor=10)

        self.input_text("popfilter2_name", popfilter2_name[0], env)

        self.click("popfilter2_submit_btn", env)
        time.sleep(2)

        # actual_add_status_text = self.get_text("managepopfilter2_status_text", env, UnivWaitFor=10)
        actual_add_status_text = self.get_status_text("managepopfilter2_status_text", env)
        # time.sleep(2)

        if actual_add_status_text == expected_add_status_text:
            self.LogScreenshot.fLogScreenshot(message=f'Able to add the PopulationFilter2 record',
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(
                message=f'Unable to find status message while adding the PopulationFilter2 record. Actual status message is '
                        f'{actual_add_status_text} and Expected status message is {expected_add_status_text}',
                pass_=False, log=True, screenshot=True)
            raise Exception("Unable to find status message while adding the PopulationFilter2 record")        

        # Fetching total rows count after adding a new PopulationFilter2
        table_rows_after = self.mngpoppage.get_table_length("managepopfilter2_table_rows_info",
                                                            "managepopfilter2_table_next_btn", table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new PopulationFilter2: {table_rows_after}',
                                          pass_=True, log=True, screenshot=True)        

        try:
            if table_rows_after > table_rows_before != table_rows_after:
                self.refreshpage()
                self.input_text("managepopfilter2_search_box", f"{popfilter2_name[0]}", env)
                result = self.get_texts('managepopfilter2_table_row_1', env)
                
                if result[0] == popfilter2_name[0]:
                    self.LogScreenshot.fLogScreenshot(message=f'Added Population filter 2 data is present in table',
                                                      pass_=True, log=True, screenshot=True)
                    popfilter2_data = f"{result[0]}"
                    return popfilter2_data
                else:
                    raise Exception("Population filter 2 data is not added")
            self.clear("managepopfilter2_search_box")
        except Exception:
            raise Exception("Error while adding the Population filter 2")
    
    def edit_multiple_popfilter2(self, locatorname, current_data, edit_upd_button, filepath, env):
        expected_update_status_text = "Population filter 2 updated successfully"
        self.refreshpage()
        time.sleep(2)
        self.go_to_page("managepopfilter2_button", env)

        # Read PopulationFilter2 name details from data sheet
        popfilter2_name = self.get_popfilter2_name(filepath, locatorname, "PopulationFilter2_name")

        self.input_text("managepopfilter2_search_box", current_data, env)
        self.click(edit_upd_button, env, UnivWaitFor=10)

        self.input_text("popfilter2_name", f"{popfilter2_name[0]}_Update", env)

        self.click("sel_update_submit", env)
        time.sleep(2)

        actual_update_status_text = self.get_status_text("managepopfilter2_status_text", env)
        time.sleep(2)

        if actual_update_status_text == expected_update_status_text:
            self.LogScreenshot.fLogScreenshot(message=f'Able to edit the PopulationFilter2 record',
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(
                message=f'Unable to find status message while edit the PopulationFilter2 record. Actual status message is '
                        f'{actual_update_status_text} and Expected status message is {expected_update_status_text}',
                pass_=False, log=True, screenshot=True)
            raise Exception("Unable to find status message while edit the PopulationFilter2 record")

        try:
            self.refreshpage()
            self.input_text("managepopfilter2_search_box", f"{popfilter2_name[0]}_Update", env)
            result = self.get_texts('manage_update_table_row_1', env)
            
            if result[0] == f"{popfilter2_name[0]}_Update":
                self.LogScreenshot.fLogScreenshot(message=f'Edited Population filter 2 data is present in table',
                                                  pass_=True, log=True, screenshot=True)
                popfilter2_data = f"{result[0]}"
                return popfilter2_data
            else:
                raise Exception("Population filter 2 data is not edited")
        except Exception:
            raise Exception("Error while editing the Population filter 2")

    def delete_multiple_manage_popfilter2(self, added_update_val, del_locator, del_locator_popup, tablerows, env):
        expected_del_status_text = "Population filter 2 deleted successfully"
        self.refreshpage()
        time.sleep(2)
        self.go_to_page("managepopfilter2_button", env)

        # Fetching total rows count before deleting a PopulationFilter2
        table_rows_before = self.mngpoppage.get_table_length("managepopfilter2_table_rows_info",
                                                             "managepopfilter2_table_next_btn", tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a PopulationFilter2: {table_rows_before}',
                                          pass_=True, log=True, screenshot=True)

        self.scroll_and_click("managepopfilter2_page_heading", env)
        self.input_text("managepopfilter2_search_box", added_update_val, env)
        self.LogScreenshot.fLogScreenshot(message=f'PopulationFilter2 option selected for deletion is : ',
                                          pass_=True, log=True, screenshot=True)        
        
        self.click(del_locator, env)
        time.sleep(2)
        self.click(del_locator_popup, env)
        time.sleep(2)

        actual_del_status_text = self.get_status_text("managepopfilter2_status_text", env)
        time.sleep(2)
                                          
        if actual_del_status_text == expected_del_status_text:
            self.LogScreenshot.fLogScreenshot(message=f'Able to delete the PopulationFilter2 record',
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(
                message=f'Unable to find status message while delete the PopulationFilter2 record. Actual status message is '
                        f'{actual_del_status_text} and Expected status message is {expected_del_status_text}',
                pass_=False, log=True, screenshot=True)
            raise Exception("Unable to find status message while delete the PopulationFilter2 record")                                          

        # Fetching total rows count after deleting a PopulationFilter2
        table_rows_after = self.mngpoppage.get_table_length("managepopfilter2_table_rows_info",
                                                            "managepopfilter2_table_next_btn", tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a PopulationFilter2: {table_rows_after}',
                                          pass_=True, log=True, screenshot=True)        

        try:
            if table_rows_before > table_rows_after != table_rows_before:
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                                  pass_=True, log=True, screenshot=False)
            self.refreshpage()
            time.sleep(2) 
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in deleting the Population filter 2")
