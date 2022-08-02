import os
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class ManagePopulationsPage(Base):

    """Constructor of the ManagePopulations Page class"""
    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    def go_to_managepopulations(self, locator):
        self.click(locator, UnivWaitFor=10)
        time.sleep(5)

    def get_template_file_details(self, filepath):
        file = pd.read_excel(filepath)
        sheet_name = list(file['manage_population_file_name'].dropna())
        sheet_path = list(os.getcwd()+file['manage_population_file_to_upload'].dropna())
        manage_pop_template = [(sheet_name[i], sheet_path[i]) for i in range(0, len(sheet_name))]
        return manage_pop_template
    
    def get_pop_data(self, filepath):
        file = pd.read_excel(filepath)
        data = list(file['Add_population_field'].dropna())
        value = list(file['Add_population_value'].dropna())
        result = [(data[i], value[i]) for i in range(0, len(data))]
        return result, value

    def add_population(self, add_locator, filepath, upload_loc, upload_file, table_rows):
        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before adding a new population
        table_rows_before = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new population: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.click(add_locator, UnivWaitFor=10)

        # Read population details from data sheet
        new_pop_data, new_pop_val = self.get_pop_data(filepath)

        for j in new_pop_data:
            self.input_text(j[0], j[1], UnivWaitFor=10)
            
        self.input_text(upload_loc, upload_file)
        self.click("submit_button")
        time.sleep(2)

        add_text = self.get_text("file_status_popup_text", UnivWaitFor=10)
        self.LogScreenshot.fLogScreenshot(message=f'Message popup: {add_text}',
                                          pass_=True, log=True, screenshot=False)
                                          
        self.assertText("Population added successfully", add_text)

        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count after adding a new population
        table_rows_after = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new population: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                    result = []
                    self.input_text("search_button", new_pop_val[1])
                    td1 = self.select_elements('manage_pop_table_row_1')
                    for m in td1:
                        result.append(m.text)

                    self.LogScreenshot.fLogScreenshot(message=f'Table data after adding a new population: {result}',
                                            pass_=True, log=True, screenshot=False)
                    
                    if result[1] == new_pop_val[1]:
                        self.LogScreenshot.fLogScreenshot(message=f'Population data is present in table',
                                            pass_=True, log=True, screenshot=False)
                        population = f"{result[0]} - {result[1]}"
                        return population
                    else:
                        raise Exception("Population data is not added")
            self.clear("search_button")
        except:
            raise Exception("Error while adding the population")

    def delete_population(self, manage_pop_pg, del_locator, filepath, del_locator_popup, tablerows):
        self.click(manage_pop_pg)
        self.refreshpage()
        time.sleep(2)
        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_before = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a population: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        # Read extraction sheet values
        new_pop_data, new_pop_val = self.get_pop_data(filepath)

        self.input_text("search_button", new_pop_val[1])
        
        self.click(del_locator)
        time.sleep(2)
        self.click(del_locator_popup)
        time.sleep(1)
        
        del_text = self.get_text("file_status_popup_text", UnivWaitFor=10)
        self.LogScreenshot.fLogScreenshot(message=f'Message popup: {del_text}',
                                          pass_=True, log=True, screenshot=False)

        self.assertText("Population deleted successfully", del_text)

        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_after = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a population: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                            pass_=True, log=True, screenshot=False)                    
        except:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                            pass_=False, log=True, screenshot=False)  
            raise Exception("Error in deleting the population")

    def add_multiple_population(self, counter, add_locator, filepath, upload_loc, upload_file, table_rows):
        self.refreshpage()
        time.sleep(2)
        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before adding a new population
        table_rows_before = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new population: {len(table_rows_before)}',
                                        pass_=True, log=True, screenshot=False)

        self.click(add_locator, UnivWaitFor=10)

        # Read population details from data sheet
        new_pop_data, new_pop_val = self.get_pop_data(filepath)

        for j in new_pop_data:
            self.input_text(j[0], f'{j[1]}_{counter}', UnivWaitFor=10)
            
        self.input_text(upload_loc, upload_file)
        self.click("submit_button")
        time.sleep(2)

        add_text = self.get_text("file_status_popup_text", UnivWaitFor=10)
        self.LogScreenshot.fLogScreenshot(message=f'Message popup: {add_text}',
                                        pass_=True, log=True, screenshot=False)
                                        
        self.assertText("Population added successfully", add_text)

        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count after adding a new population
        table_rows_after = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new population: {len(table_rows_after)}',
                                        pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                    result = []
                    self.input_text("search_button", f'{new_pop_val[1]}_{counter}')
                    td1 = self.select_elements('manage_pop_table_row_1')
                    for m in td1:
                        result.append(m.text)

                    self.LogScreenshot.fLogScreenshot(message=f'Table data after adding a new population: {result}',
                                            pass_=True, log=True, screenshot=False)
                    
                    if result[1] == f'{new_pop_val[1]}_{counter}':
                        self.LogScreenshot.fLogScreenshot(message=f'Population data is present in table',
                                            pass_=True, log=True, screenshot=False)
                        population = f"{result[1]}"
                        return population
                    else:
                        raise Exception("Population data is not added")
            self.clear("search_button")
            self.refreshpage()
            time.sleep(2)
        except:
            raise Exception("Error while adding the population")

    def delete_multiple_population(self, pop_value, del_locator, del_locator_popup, tablerows):
        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_before = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a population: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.input_text("search_button", pop_value)
        
        self.click(del_locator)
        time.sleep(2)
        self.click(del_locator_popup)
        time.sleep(2)
        
        del_text = self.get_text("file_status_popup_text", UnivWaitFor=10)
        self.LogScreenshot.fLogScreenshot(message=f'Message popup: {del_text}',
                                          pass_=True, log=True, screenshot=False)

        self.assertText("Population deleted successfully", del_text)
        time.sleep(2)

        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_after = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a population: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                            pass_=True, log=True, screenshot=False)
            self.refreshpage()
            time.sleep(2)                
        except:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                            pass_=False, log=True, screenshot=False)  
            raise Exception("Error in deleting the population")