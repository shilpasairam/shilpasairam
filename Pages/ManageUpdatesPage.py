from datetime import date
import random
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.ManagePopulationsPage import ManagePopulationsPage
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class ManageUpdatesPage(Base):
    
    # filepath = ReadConfig.getmanageupdatesdata()

    """Constructor of the ManageUpdates Page class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of ExtendedBase class
        self.exbase = ExtendedBase(self.driver, self.extra)         
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    # Reading Population data for ManageUpdates Page
    def get_updates_pop_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop = df.loc[df['Name'] == locatorname]['Population'].dropna().to_list()
        return pop

    def add_updates(self, manage_update_pg, add_upd_button, sel_pop_val, date_val, table_rows, dateval_to_search):
        self.click(manage_update_pg)
        self.refreshpage()
        table_ele = self.select_element("sel_table_entries_dropdown")
        select = Select(table_ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before adding a new update
        table_rows_before = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new update: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.click(add_upd_button, UnivWaitFor=10)

        pop_ele = self.select_element("sel_pop_update_dropdown")
        select = Select(pop_ele)
        select.select_by_visible_text(sel_pop_val)

        self.click("sel_update_date")
        # self.input_text("sel_update_date", date_val)
        self.select_calendar_date(date_val)

        self.click("sel_update_type")
        update_type_ele = self.select_element("sel_update_type")
        select = Select(update_type_ele)
        select.select_by_index(1)

        self.click("sel_update_submit")
        time.sleep(1)

        add_text = self.get_text("manage_update_status_text", UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Update added successfully", add_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to add the updates',
                                          pass_=True, log=True, screenshot=True)

        table_ele = self.select_element("sel_table_entries_dropdown")
        select = Select(table_ele)
        select.select_by_visible_text("100")

        # Fetching total rows count after adding a new update
        table_rows_after = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new update: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                self.refreshpage()
                result = []
                self.input_text("manage_update_search_box", f"{sel_pop_val} - {dateval_to_search}")
                td1 = self.select_elements('manage_update_table_row_1')
                for n in td1:
                    result.append(n.text)

                self.LogScreenshot.fLogScreenshot(message=f'Table data after adding a new update: {result}',
                                                  pass_=True, log=True, screenshot=False)
                
                if result[0] == sel_pop_val:
                    self.LogScreenshot.fLogScreenshot(message=f'Population update data is present in table',
                                                      pass_=True, log=True, screenshot=False)
                    update_data = f"{result[0]} - {result[2]} - {result[1].replace('0', '')}"
                    return update_data
                else:
                    raise Exception("Population update data is not added")
            self.clear("manage_update_search_box")
        except Exception:
            raise Exception("Error while adding the population updates")

    def delete_manage_update(self, manage_update_pg, sel_pop_val, del_locator, del_locator_popup, tablerows):
        self.click(manage_update_pg)
        self.refreshpage()
        time.sleep(2)
        ele = self.select_element("sel_table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a update
        table_rows_before = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a update: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.input_text("manage_update_search_box", sel_pop_val)
        
        self.click(del_locator)
        time.sleep(2)
        self.click(del_locator_popup)
        time.sleep(2)

        del_text = self.get_text("manage_update_status_text", UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Update deleted successfully", del_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to delete the updates',
                                          pass_=True, log=True, screenshot=True)

        ele = self.select_element("sel_table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count after deleting a update
        table_rows_after = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a update: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                                  pass_=True, log=True, screenshot=False)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in deleting the update")
    
    def add_multiple_updates(self, locatorname, filepath, add_upd_button, date_val, table_rows, dateval_to_search, env):
        expected_status_text = "Update added successfully"
        # Read population details from data sheet
        pop_details = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_name',
                                                   'indication_type')

        # Fetching total rows count before adding a new update
        table_rows_before = self.mngpoppage.get_table_length("manage_update_table_rows_info",
                                                             "manage_update_table_next_btn", table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new update: {table_rows_before}',
                                          pass_=True, log=True, screenshot=True)

        self.scroll("manageupdates_page_heading", env)
        self.click(add_upd_button, env, UnivWaitFor=10)

        pop_ele = self.select_element("sel_pop_update_dropdown", env)
        select = Select(pop_ele)
        select.select_by_visible_text(pop_details[0][0])
        sel_pop_val = select.first_selected_option.text

        self.click("sel_update_date", env)
        self.select_calendar_date(date_val)

        self.LogScreenshot.fLogScreenshot(message=f"Entered update details are : ",
                                            pass_=True, log=True, screenshot=True)         

        '''As per LIVEHTA-1657 changes commenting the below section'''
        # self.click("sel_update_type")
        # update_type_ele = self.select_element("sel_update_type")
        # select = Select(update_type_ele)
        # select.select_by_index(1)

        self.click("sel_update_submit", env)
        time.sleep(2)

        actual_status_text = self.get_status_text("manage_update_status_text", env)
        # time.sleep(2)

        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Addition of New Update is success",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding New Update. "
                                                      f"Actual status message is {actual_status_text} and Expected "
                                                      f"status message is {expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while adding New Update")        

        # Fetching total rows count after adding a new update
        table_rows_after = self.mngpoppage.get_table_length("manage_update_table_rows_info",
                                                            "manage_update_table_next_btn", table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new update: {table_rows_after}',
                                          pass_=True, log=True, screenshot=True)        

        try:
            if table_rows_after > table_rows_before != table_rows_after:
                self.refreshpage()
                result = []
                self.input_text("manage_update_search_box", f"{sel_pop_val} - {dateval_to_search}", env)
                td1 = self.select_elements('manage_update_table_row_1', env)
                for n in td1:
                    result.append(n.text)
                
                if result[0] == sel_pop_val and result[1] == pop_details[0][1]:
                    self.LogScreenshot.fLogScreenshot(message=f'Population update data is present in table',
                                                      pass_=True, log=True, screenshot=True)
                    update_data = f"{result[0]} - {result[2]}"
                    return update_data
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Population update data is not added. Actual update data is : {result}',
                                                    pass_=False, log=True, screenshot=False)                    
                    raise Exception("Population update data is not added")
            self.clear("manage_update_search_box", env)
        except Exception:
            raise Exception("Error while adding the population updates")
    
    def edit_multiple_updates(self, current_data, edit_upd_button, edit_date_val, dateval_to_search, indication, env):
        expected_status_text = "Update updated successfully"

        self.input_text("manage_update_search_box", current_data, env)
        self.LogScreenshot.fLogScreenshot(message=f"Population opted for performing the edit operation is : ",
                                            pass_=True, log=True, screenshot=True)        
        self.click(edit_upd_button, env, UnivWaitFor=10)

        pop_ele = self.select_element("sel_pop_update_dropdown", env)
        select = Select(pop_ele)
        sel_pop_val = select.first_selected_option.text

        self.click("sel_update_date", env)
        # Manipulating the date values when values point to month end
        if edit_date_val in [30, 31]:
            edit_date_val = edit_date_val - 10
        else:
            edit_date_val = edit_date_val + 1
        self.select_calendar_date(edit_date_val)

        '''As per LIVEHTA-1657 changes commenting the below section'''
        # self.click("sel_update_type")
        # update_type_ele = self.select_element("sel_update_type")
        # select = Select(update_type_ele)
        # select.select_by_index(2)

        # if select.first_selected_option.text == 'Congress search':
        #     congress_ele = self.select_element("sel_congress_meeting")
        #     select1 = Select(congress_ele)
        #     select1.select_by_index(2)

        self.click("sel_update_submit", env)
        time.sleep(2)

        actual_status_text = self.get_status_text("manage_update_status_text", env)
        # time.sleep(2)

        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Editing the existing Update is success",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while editing the "
                                                      f"Update data. Actual status message is {actual_status_text} "
                                                      f"and Expected status message is {expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while editing the Update data")        

        try:
            self.refreshpage()
            result = []
            self.input_text("manage_update_search_box", f"{sel_pop_val} - {dateval_to_search}", env)
            td1 = self.select_elements('manage_update_table_row_1', env)
            for n in td1:
                result.append(n.text)
            
            if result[0] == sel_pop_val and result[1] == indication:
                self.LogScreenshot.fLogScreenshot(message=f'Edited Population update data is present in table',
                                                  pass_=True, log=True, screenshot=True)
                update_data = f"{result[0]} - {result[2]}"
                return update_data
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Population update data is not edited. Actual update data is : {result}',
                                    pass_=False, log=True, screenshot=False)
                raise Exception("Population update data is not edited")
        except Exception:
            raise Exception("Error while editing the population updates")

    def delete_multiple_manage_updates(self, added_update_val, del_locator, del_locator_popup, tablerows, env):
        expected_status_text = "Update deleted successfully"

        # Fetching total rows count before deleting a update
        table_rows_before = self.mngpoppage.get_table_length("manage_update_table_rows_info",
                                                             "manage_update_table_next_btn", tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a update: {table_rows_before}',
                                          pass_=True, log=True, screenshot=True)

        self.scroll("manageupdates_page_heading", env)
        self.input_text("manage_update_search_box", added_update_val, env)
        self.LogScreenshot.fLogScreenshot(message=f"Population opted for performing the delete operation is : ",
                                            pass_=True, log=True, screenshot=True)        
        
        self.click(del_locator, env)
        time.sleep(2)
        self.click(del_locator_popup, env)
        time.sleep(2)

        actual_status_text = self.get_status_text("manage_update_status_text", env)
        # time.sleep(2)

        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Deleting the existing Update is success",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting the "
                                                      f"Update data Actual status message is {actual_status_text} "
                                                      f"and Expected status message is {expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while deleting the Update data")        

        # Fetching total rows count after deleting a update
        table_rows_after = self.mngpoppage.get_table_length("manage_update_table_rows_info",
                                                            "manage_update_table_next_btn", tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a update: {table_rows_after}',
                                          pass_=True, log=True, screenshot=True)        

        try:
            if table_rows_before > table_rows_after != table_rows_before:
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                                  pass_=True, log=True, screenshot=False)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in deleting the update")
