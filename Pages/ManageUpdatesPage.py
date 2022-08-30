from datetime import date
import random
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ManagePopulationsPage import ManagePopulationsPage
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class ManageUpdatesPage(Base):

    """Constructor of the ManageUpdates Page class"""
    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
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

    def go_to_manageupdates(self, locator):
        self.click(locator, UnivWaitFor=10)
        time.sleep(5)

    def add_updates(self, manage_update_pg, add_upd_button, sel_pop_val, date_val, table_rows, dateval_to_search):
        self.click(manage_update_pg)
        self.refreshpage()
        table_ele = self.select_element("sel_table_entries_dropdown")
        select = Select(table_ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before adding a new population
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

        # Fetching total rows count after adding a new population
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
        except:
            raise Exception("Error while adding the population updates")

    def delete_manage_update(self, manage_update_pg, sel_pop_val, del_locator, del_locator_popup, tablerows):
        self.click(manage_update_pg)
        self.refreshpage()
        time.sleep(2)
        ele = self.select_element("sel_table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a file from top of the table
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

        # Fetching total rows count before deleting a file from top of the table
        table_rows_after = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a update: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                            pass_=True, log=True, screenshot=False)                    
        except:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                            pass_=False, log=True, screenshot=False)  
            raise Exception("Error in deleting the update")
    
    def add_multiple_updates(self, pop_index, add_upd_button, date_val, table_rows, dateval_to_search):
        self.refreshpage()
        time.sleep(5)
        table_ele = self.select_element("sel_table_entries_dropdown")
        select = Select(table_ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before adding a new population
        table_rows_before = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new update: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.click(add_upd_button, UnivWaitFor=10)

        pop_ele = self.select_element("sel_pop_update_dropdown")
        select = Select(pop_ele)
        # options_list_res = []
        # options_list = select.options
        # for xy in options_list:
        #     options_list_res.append(xy.text)
        # self.LogScreenshot.fLogScreenshot(message=f'Options list values are: {options_list_res} and Length is : {len(options_list_res)}',
        #                                   pass_=True, log=True, screenshot=False)
        # pop_index = random.randint(1, len(options_list_res)-1)
        # self.LogScreenshot.fLogScreenshot(message=f'Index value is {pop_index}',
        #                                   pass_=True, log=True, screenshot=False)
        select.select_by_index(pop_index)
        sel_pop_val = select.first_selected_option.text

        self.click("sel_update_date")
        self.select_calendar_date(date_val)

        self.click("sel_update_type")
        update_type_ele = self.select_element("sel_update_type")
        select = Select(update_type_ele)
        select.select_by_index(1)

        self.click("sel_update_submit")
        time.sleep(2)

        add_text = self.get_text("manage_update_status_text", UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Update added successfully", add_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to add the updates record',
                                    pass_=True, log=True, screenshot=True)

        table_ele = self.select_element("sel_table_entries_dropdown")
        select = Select(table_ele)
        select.select_by_visible_text("100")

        # Fetching total rows count after adding a new population
        table_rows_after = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new update: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                    self.refreshpage()
                    self.LogScreenshot.fLogScreenshot(message=f'Text to search: {sel_pop_val} - {dateval_to_search}',
                                          pass_=True, log=True, screenshot=False)
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
                        update_data = f"{result[0]} - {result[1]}"
                        return update_data
                    else:
                        raise Exception("Population update data is not added")
            self.clear("manage_update_search_box")
        except:
            raise Exception("Error while adding the population updates")
    
    def edit_multiple_updates(self, current_data, edit_upd_button, date_val, dateval_to_search):
        self.refreshpage()
        time.sleep(5)

        self.input_text("manage_update_search_box", current_data)
        self.click(edit_upd_button, UnivWaitFor=10)

        pop_ele = self.select_element("sel_pop_update_dropdown")
        select = Select(pop_ele)
        sel_pop_val = select.first_selected_option.text

        self.click("sel_update_date")
        self.select_calendar_date(date_val+1)

        self.click("sel_update_type")
        update_type_ele = self.select_element("sel_update_type")
        select = Select(update_type_ele)
        select.select_by_index(2)

        if select.first_selected_option.text == 'Congress search':
            congress_ele = self.select_element("sel_congress_meeting")
            select1 = Select(congress_ele)
            select1.select_by_index(2)

        self.click("sel_update_submit")
        time.sleep(2)

        update_text = self.get_text("manage_update_status_text", UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Update updated successfully", update_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to edit the updates record',
                                    pass_=True, log=True, screenshot=True)

        try:
            self.refreshpage()
            self.LogScreenshot.fLogScreenshot(message=f'Text to search: {sel_pop_val} - {dateval_to_search}',
                                    pass_=True, log=True, screenshot=False)
            result = []
            self.input_text("manage_update_search_box", f"{sel_pop_val} - {dateval_to_search}")
            td1 = self.select_elements('manage_update_table_row_1')
            for n in td1:
                result.append(n.text)

            self.LogScreenshot.fLogScreenshot(message=f'Table data after editing the update: {result}',
                                    pass_=True, log=True, screenshot=False)
            
            if result[0] == sel_pop_val:
                self.LogScreenshot.fLogScreenshot(message=f'Edited Population update data is present in table',
                                    pass_=True, log=True, screenshot=False)
                update_data = f"{result[0]} - {result[1]}"
                return update_data
            else:
                raise Exception("Population update data is not eduted")
        except:
            raise Exception("Error while adding the population updates")

    def delete_multiple_manage_updates(self, added_update_val, del_locator, del_locator_popup, tablerows):
        self.refreshpage()
        time.sleep(2)
        ele = self.select_element("sel_table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_before = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a update: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.input_text("manage_update_search_box", added_update_val)
        
        self.click(del_locator)
        time.sleep(2)
        self.click(del_locator_popup)
        time.sleep(2)

        del_text = self.get_text("manage_update_status_text", UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Update deleted successfully", del_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to delete the updates record',
                                    pass_=True, log=True, screenshot=True)

        ele = self.select_element("sel_table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_after = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a update: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                            pass_=True, log=True, screenshot=False)                    
        except:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                            pass_=False, log=True, screenshot=False)  
            raise Exception("Error in deleting the update")
