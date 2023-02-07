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


class LineofTherapyPage(Base):

    """Constructor of the LineofTherapy Page class"""
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
    
    # Reading Population data for LineofTherapy Page
    def get_updates_pop_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop = df.loc[df['Name'] == locatorname]['Population'].dropna().to_list()
        return pop

    # Reading Population data for LineofTherapy Page
    def get_lot_name(self, filepath, locatorname, columnname):
        df = pd.read_excel(filepath)
        lot_name = df.loc[df['Name'] == locatorname][columnname].dropna().to_list()
        return lot_name
    
    def get_expected_data(self, filepath, columnname):
        file = pd.read_excel(filepath)
        lot_options = list(file[columnname].dropna())
        return lot_options    
    
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
        lot_optns_text = []
        options = self.select_elements(table_rows_data, env)
        for k in options:
            lot_optns_text.append(k.text)
        
        # Iterate over the remaining pages and append the row data
        for i in range(1, page_counter):
            self.click(table_next_btn, env)
            time.sleep(1)
            options1 = self.select_elements(table_rows_data, env)
            for v in options1:
                lot_optns_text.append(v.text)
        
        return lot_optns_text    
    
    def check_lot_ui_elements(self, filepath, env):
        # Read Expected messages
        expected_data = self.get_expected_data(filepath, "Expected_ui_elements")

        pagedetails = [self.get_text("managelot_page_heading", env), self.get_text("managelot_pagecontent", env)]

        if expected_data == pagedetails:
            self.LogScreenshot.fLogScreenshot(message=f"Line of Therapy Page name and page content is present in UI",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Line of Therapy Page name and page content is not present in "
                                                      f"UI", pass_=False, log=True, screenshot=False)
            raise Exception(f"Line of Therapy Page name and page content is not present in UI")            

        if self.isdisplayed("add_lot_btn", env) and self.isdisplayed("managelot_resultpanel_heading", env):
            self.LogScreenshot.fLogScreenshot(message=f"Add Line of Therapy button and Result panel heading is "
                                                      f"present UI", pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Add Line of Therapy button and Result panel heading is not "
                                                      f"present UI", pass_=False, log=True, screenshot=False)
            raise Exception(f"Add Line of Therapy button and Result panel heading is not present UI")
    
    def add_multiple_lot(self, locatorname, add_lot_button, table_rows, filepath, env):
        expected_add_status_text = "Line of Therapy added successfully"

        # Read LoT name details from data sheet
        lot_name = self.get_lot_name(filepath, locatorname, "LOT_name")

        # Read Expected LoT options from data sheet
        expected_lot_options = self.get_expected_data(filepath, 'Expected_lot_options')    

        # Validate the UI page elements
        self.check_lot_ui_elements(filepath, env)

        # Fetch the complete LoT data from the table
        complete_lot_table_data = self.get_table_data("managelot_table_rows_info", "managelot_table_next_btn",
                                                      "managelot_table_rows_data", env)
        
        # Compare the expected mandatory lot option with actual table data
        for j in expected_lot_options:
            if j in complete_lot_table_data:
                self.LogScreenshot.fLogScreenshot(message=f"Mandatory LoT '{j}' option is present in the table.",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Mandatory LoT '{j}' option is present in the table.",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Mandatory LoT '{j}' option is present in the table.")            
        
        # Fetching total rows count before adding a new LoT
        table_rows_before = self.mngpoppage.get_table_length("managelot_table_rows_info",
                                                             "managelot_table_next_btn", table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new LoT: {table_rows_before}',
                                          pass_=True, log=True, screenshot=False)        
        
        self.scroll("managelot_page_heading", env)
        self.click(add_lot_button, env, UnivWaitFor=10)

        self.input_text("lot_name", lot_name[0], env)

        self.click("lot_submit_btn", env)
        time.sleep(2)

        actual_add_status_text = self.get_text("managelot_status_text", env, UnivWaitFor=10)
        # time.sleep(2)

        if actual_add_status_text == expected_add_status_text:
            self.LogScreenshot.fLogScreenshot(message=f'Able to add the LoT record',
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(
                message=f'Unable to find status message while adding the LoT record',
                pass_=False, log=True, screenshot=True)
            raise Exception("Unable to find status message while adding the LoT record")        

        # Fetching total rows count after adding a new LoT
        table_rows_after = self.mngpoppage.get_table_length("managelot_table_rows_info",
                                                            "managelot_table_next_btn", table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new LoT: {table_rows_after}',
                                          pass_=True, log=True, screenshot=False)        

        try:
            if table_rows_after > table_rows_before != table_rows_after:
                self.refreshpage()
                result = []
                self.input_text("managelot_search_box", f"{lot_name[0]}", env)
                td1 = self.select_elements('managelot_table_row_1', env)
                for n in td1:
                    result.append(n.text)
                
                if result[0] == lot_name[0]:
                    self.LogScreenshot.fLogScreenshot(message=f'Added Line of Therapy data is present in table',
                                                      pass_=True, log=True, screenshot=True)
                    lot_data = f"{result[0]}"
                    return lot_data
                else:
                    raise Exception("Line of Therapy data is not added")
            self.clear("managelot_search_box")
        except Exception:
            raise Exception("Error while adding the Line of Therapy")
    
    def edit_multiple_lot(self, locatorname, current_data, edit_upd_button, filepath, env):
        self.refreshpage()
        time.sleep(5)

        # Read LoT name details from data sheet
        lot_name = self.get_lot_name(filepath, locatorname, "LOT_name")

        self.input_text("managelot_search_box", current_data, env)
        self.click(edit_upd_button, env, UnivWaitFor=10)

        self.input_text("lot_name", f"{lot_name[0]}_Update", env)

        self.click("sel_update_submit", env)
        time.sleep(2)

        update_text = self.get_text("managelot_status_text", env, UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Line of Therapy updated successfully", update_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to edit the LoT record',
                                          pass_=True, log=True, screenshot=True)

        try:
            self.refreshpage()
            result = []
            self.input_text("managelot_search_box", f"{lot_name[0]}_Update", env)
            td1 = self.select_elements('manage_update_table_row_1', env)
            for n in td1:
                result.append(n.text)
            
            if result[0] == f"{lot_name[0]}_Update":
                self.LogScreenshot.fLogScreenshot(message=f'Edited Line of Therapy data is present in table',
                                                  pass_=True, log=True, screenshot=True)
                lot_data = f"{result[0]}"
                return lot_data
            else:
                raise Exception("Line of Therapy data is not edited")
        except Exception:
            raise Exception("Error while editing the Line of Therapy")

    def delete_multiple_manage_lot(self, added_update_val, del_locator, del_locator_popup, tablerows, env):
        self.refreshpage()
        time.sleep(2)

        # Fetching total rows count before deleting a LoT
        table_rows_before = self.mngpoppage.get_table_length("managelot_table_rows_info",
                                                             "managelot_table_next_btn", tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a LoT: {table_rows_before}',
                                          pass_=True, log=True, screenshot=False)

        self.scroll("managelot_page_heading", env)
        self.input_text("managelot_search_box", added_update_val, env)
        self.LogScreenshot.fLogScreenshot(message=f'LoT option selected for deletion is : ',
                                          pass_=True, log=True, screenshot=True)        
        
        self.click(del_locator, env)
        time.sleep(2)
        self.click(del_locator_popup, env)
        time.sleep(2)

        del_text = self.get_text("managelot_status_text", env, UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Line of Therapy deleted successfully", del_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to delete the LoT record',
                                          pass_=True, log=True, screenshot=True)

        # Fetching total rows count after deleting a LoT
        table_rows_after = self.mngpoppage.get_table_length("managelot_table_rows_info",
                                                            "managelot_table_next_btn", tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a LoT: {table_rows_after}',
                                          pass_=True, log=True, screenshot=False)        

        try:
            if table_rows_before > table_rows_after != table_rows_before:
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                                  pass_=True, log=True, screenshot=False)
            self.refreshpage()
            time.sleep(2) 
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in deleting the Line of Therapy")
