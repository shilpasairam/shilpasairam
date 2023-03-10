import math
import os
import time
import openpyxl
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class ManagePopulationsPage(Base):

    """Constructor of the ManagePopulations Page class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of ExtendedBase class
        self.exbase = ExtendedBase(self.driver, extra)        
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    # def go_to_managepopulations(self, locator, env):
    #     self.click(locator, env, UnivWaitFor=10)
    #     time.sleep(5)

    # def get_template_file_details(self, filepath):
    #     file = pd.read_excel(filepath)
    #     sheet_name = list(file['manage_population_file_name'].dropna())
    #     sheet_path = list(os.getcwd()+file['manage_population_file_to_upload'].dropna())
    #     manage_pop_template = [(sheet_name[i], sheet_path[i]) for i in range(0, len(sheet_name))]
    #     return manage_pop_template
    
    # def get_pop_data(self, filepath):
    #     file = pd.read_excel(filepath)
    #     data = list(file['Add_population_field'].dropna())
    #     value = list(file['Add_population_value'].dropna())
    #     result = [(data[i], value[i]) for i in range(0, len(data))]
    #     return result, value

    def get_template_file_details(self, filepath, locatorname, column_name):
        df = pd.read_excel(filepath)
        upload_sheet_path = os.getcwd()+(df.loc[df['Name'] == locatorname][column_name].to_list()[0])
        # manage_pop_template = [(sheet_name[i], sheet_path[i]) for i in range(0, len(sheet_name))]
        return upload_sheet_path
    
    def get_pop_data(self, filepath, locatorname, column_name):
        df = pd.read_excel(filepath)
        data = df.loc[df['Name'] == locatorname]['Add_population_field'].to_list()
        value = df.loc[df['Name'] == locatorname][column_name].to_list()
        result = [(data[i], value[i]) for i in range(0, len(data))]
        return result, value

    def add_population(self, add_locator, filepath, upload_loc, upload_file, table_rows):
        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count before adding a new population
        table_rows_before = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new population: '
                                                  f'{len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.click(add_locator, UnivWaitFor=10)

        # Read population details from data sheet
        new_pop_data, new_pop_val = self.get_pop_data(filepath)

        for j in new_pop_data:
            self.input_text(j[0], j[1], UnivWaitFor=10)
            
        self.input_text(upload_loc, upload_file)
        self.click("submit_button")
        time.sleep(2)

        add_text = self.get_text("population_status_popup_text", UnivWaitFor=10)
        time.sleep(2)
                                          
        self.assertText("Population added successfully", add_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to add the population record',
                                          pass_=True, log=True, screenshot=True)

        ele = self.select_element("table_entries_dropdown")
        select = Select(ele)
        select.select_by_visible_text("100")

        # Fetching total rows count after adding a new population
        table_rows_after = self.select_elements(table_rows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new population: '
                                                  f'{len(table_rows_after)}',
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
        except Exception:
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
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a population: '
                                                  f'{len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        # Read extraction sheet values
        new_pop_data, new_pop_val = self.get_pop_data(filepath)

        self.input_text("search_button", new_pop_val[1])
        
        self.click(del_locator)
        time.sleep(2)
        self.click(del_locator_popup)
        time.sleep(1)
        
        del_text = self.get_text("population_status_popup_text", UnivWaitFor=10)
        time.sleep(2)

        self.assertText("Population deleted successfully", del_text)
        self.LogScreenshot.fLogScreenshot(message=f'Able to delete the population record',
                                          pass_=True, log=True, screenshot=True)

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
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in deleting the population")

    # Find the total row count if data is being ordered using Pagination
    def get_table_length(self, table_info, table_next_btn, table_rows, env):
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
        # Get the length of row from the landing page
        initial_rows_count = self.select_elements(table_rows, env)
        table_row_count = len(initial_rows_count)
        # Iterate over the remaining pages and append the row counts
        for i in range(1, page_counter):
            self.click(table_next_btn, env)
            time.sleep(1)
            next_rows_count = self.select_elements(table_rows, env)
            table_row_count += len(next_rows_count)
        
        return table_row_count

    def add_multiple_population(self, locatorname, add_locator, filepath, table_rows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Addition of Oncology Population validation is started***",
                                          pass_=True, log=True, screenshot=False)        
        expected_status_text = "Population added successfully"
        
        # Fetching total rows count before adding a new population
        table_rows_before = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                  table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new population: '
                                                  f'{table_rows_before}',
                                          pass_=True, log=True, screenshot=False)        

        self.scroll("managepopulation_page_heading", env)
        self.click(add_locator, env, UnivWaitFor=10)

        # Read the file name and path required to upload
        upload_file_path = self.get_template_file_details(filepath, locatorname, 'manage_population_file_to_upload')
        # Read population details from data sheet
        new_pop_data, new_pop_val = self.get_pop_data(filepath, locatorname, 'Add_population_value')

        for j in new_pop_data:
            self.input_text(j[0], f'{j[1]}', env, UnivWaitFor=10)
            
        self.input_text("template_file_upload", upload_file_path, env)
        self.click("submit_button", env)
        time.sleep(2)

        # actual_status_text = self.get_text("population_status_popup_text", env, UnivWaitFor=10)
        actual_status_text = self.get_status_text("population_status_popup_text", env)
        # time.sleep(2)

        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Addition of New Population is success",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding New Population. "
                                                      f"Actual status message is {actual_status_text} and Expected "
                                                      f"status message is {expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while adding New Population")

        # Fetching total rows count after adding a new population
        table_rows_after = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                 table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new population: '
                                                  f'{table_rows_after}',
                                          pass_=True, log=True, screenshot=False)        

        try:
            if table_rows_after > table_rows_before != table_rows_after:
                result = []
                self.scroll("managepopulation_page_heading", env)
                self.input_text("search_button", f'{new_pop_val[1]}', env)
                td1 = self.select_elements('manage_pop_table_row_1', env)
                for m in td1:
                    result.append(m.text)

                self.LogScreenshot.fLogScreenshot(message=f'Table data after adding a new population: {result}',
                                                  pass_=True, log=True, screenshot=False)
                
                if result[2] == f'{new_pop_val[1]}':
                    self.LogScreenshot.fLogScreenshot(message=f'Population data is present in table',
                                                      pass_=True, log=True, screenshot=False)
                    population = f"{result[2]}"
                    return population
                else:
                    raise Exception("Population data is not added")
            self.clear("search_button", env)
            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Addition of Oncology Population validation is started***",
                                              pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Error while adding the population")

    def edit_multiple_population(self, locatorname, pop_name, edit_locator, filepath, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Edit Oncology population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        expected_status_text = "Population updated successfully"

        self.input_text("search_button", f'{pop_name}', env)
        self.click(edit_locator, env, UnivWaitFor=10)
        # Read the file name and path required to upload
        upload_file_path = self.get_template_file_details(filepath, locatorname,
                                                          'manage_population_updatedfile_to_upload')
        # Read population details from data sheet
        edit_pop_data, edit_pop_val = self.get_pop_data(filepath, locatorname, 'Edit_population_value')

        for j in edit_pop_data:
            self.input_text(j[0], f'{j[1]}', env, UnivWaitFor=10)
            
        self.input_text("template_file_upload", upload_file_path, env)
        self.click("submit_button", env)
        time.sleep(2)

        # actual_status_text = self.get_text("population_status_popup_text", env, UnivWaitFor=10)
        actual_status_text = self.get_status_text("population_status_popup_text", env)
        # time.sleep(2)
        
        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Editing the existing Population is success",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while editing the "
                                                      f"Population data. Actual status message is {actual_status_text} "
                                                      f"and Expected status message is {expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while editing the Population data")

        try:
            result = []
            self.input_text("search_button", f'{edit_pop_val[1]}', env)
            td1 = self.select_elements('manage_pop_table_row_1', env)
            for m in td1:
                result.append(m.text)

            self.LogScreenshot.fLogScreenshot(message=f'Table data after editing the existing population: {result}',
                                              pass_=True, log=True, screenshot=False)
            
            if result[2] == f'{edit_pop_val[1]}':
                self.LogScreenshot.fLogScreenshot(message=f'Edited Population data is present in table',
                                                  pass_=True, log=True, screenshot=False)
                population = f"{result[2]}"
                return population
            
            self.clear("search_button", env)
            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Edit Oncology population validation is completed***",
                                              pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Error while editing the population")

    def delete_multiple_population(self, pop_value, del_locator, del_locator_popup, tablerows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of Population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        expected_status_text = "Population deleted successfully"
        
        # Fetching total rows count before deleting a new population
        table_rows_before = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                  tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a population: '
                                                  f'{table_rows_before}',
                                          pass_=True, log=True, screenshot=False)        

        self.scroll("managepopulation_page_heading", env)
        self.input_text("search_button", pop_value, env)

        self.LogScreenshot.fLogScreenshot(message=f"Selected Population for Deletion is : ",
                                          pass_=True, log=True, screenshot=True)
        
        result = []
        td1 = self.select_elements('manage_pop_table_row_1', env)
        for m in td1:
            result.append(m.text)
        
        expected_popup_msg = f"Are you sure you want to delete population '{result[0]} - {result[2]}' ?"
        
        self.click(del_locator, env)
        time.sleep(2)

        # Check the Pop-up message details
        actual_popup_msg = self.get_text("delete_popup_msg", env)

        if actual_popup_msg == expected_popup_msg:
            self.LogScreenshot.fLogScreenshot(message=f"Delete Confirmation Pop-up window : ",
                                              pass_=True, log=True, screenshot=True)
            self.click(del_locator_popup, env)
            time.sleep(2)
            if del_locator_popup == 'delete_population_popup_cancel':
                self.clear("search_button", env)
                self.refreshpage()
                time.sleep(2)
            else:
                actual_status_text = self.get_status_text("population_status_popup_text", env)
                if actual_status_text == expected_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"Deleting the existing Population is success",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting the "
                                                              f"Population data. Actual status message is "
                                                              f"{actual_status_text} and Expected status message is "
                                                              f"{expected_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while deleting the Population data")
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Delete pop-up message. Actual popup "
                                                      f"message is : {actual_popup_msg}. Expected popup message is : "
                                                      f"{expected_popup_msg}.",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Mismatch found in Delete pop-up message.")

        # Fetching total rows count after deleting a new population
        table_rows_after = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                 tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a population: {table_rows_after}',
                                          pass_=True, log=True, screenshot=False)        

        try:
            if table_rows_before == table_rows_after:
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not deleted as expected after '
                                                          f'clicking on cancel button.',
                                                  pass_=True, log=True, screenshot=False)
            elif table_rows_before > table_rows_after != table_rows_before:
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Record deletion is not successful. Table length before "
                                                          f"deletion is : '{table_rows_before}' and Table lenght "
                                                          f"after deletion is : '{table_rows_after}'.",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Record deletion is not successful.")
            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Deletion of Population validation is completed***",
                                              pass_=True, log=True, screenshot=False)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f'Population Record deletion is not successful',
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in deleting the population")

    def non_onocolgy_check_field_level_err_msg(self, locatorname, add_locator, filepath, table_rows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Field Level Validation while adding new Non-Oncology "
                                                  f"population is started***", pass_=True, log=True, screenshot=False)

        # Read required population and endpoint details
        field_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'population_field')
        ep1_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'ep1_field')
        ep2_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'ep2_field')
        ep3_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'ep3_field')
        expected_field_level_err_msg = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1',
                                                                           'expected_filed_level_err_msg')

        # Fetching total rows count before adding a new population
        table_rows_before = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                  table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new population: '
                                                  f'{table_rows_before}',
                                          pass_=True, log=True, screenshot=False)        

        self.scroll("managepopulation_page_heading", env)
        self.click(add_locator, env, UnivWaitFor=10)

        # Select Non-oncology tab
        self.click("non_oncology_tab", env)

        # Navigate through all fields without entering any values
        for i in field_locs:
            self.input_text(i, f'{Keys.TAB}', env, UnivWaitFor=10)
        
        self.click("add_ep_btn", env)

        self.LogScreenshot.fLogScreenshot(message=f'Name, Client and Description Field Level Error Messages :',
                                          pass_=True, log=True, screenshot=True)

        # Navigate through all fields without entering any values
        for j in ep1_locs:
            self.input_text(j, f'{Keys.TAB}', env, UnivWaitFor=10)

        self.click("add_ep_btn", env)

        self.LogScreenshot.fLogScreenshot(message=f'Endpoint1 Field Level Error Messages :',
                                          pass_=True, log=True, screenshot=True)
        
        # Navigate through all fields without entering any values
        for k in ep2_locs:
            self.input_text(k, f'{Keys.TAB}', env, UnivWaitFor=10)

        self.click("add_ep_btn", env)

        self.LogScreenshot.fLogScreenshot(message=f'Endpoint2 Field Level Error Messages :',
                                          pass_=True, log=True, screenshot=True)

        # Navigate through all fields without entering any values
        for v in ep3_locs:
            self.input_text(v, f'{Keys.TAB}', env, UnivWaitFor=10)

        self.LogScreenshot.fLogScreenshot(message=f'Endpoint3 Field Level Error Messages :',
                                          pass_=True, log=True, screenshot=True)

        # Check presence of Add Endpoint button after adding 3 endpoint details
        if not self.isvisible("add_ep_btn", env, "Add Endpoint button"):
            self.LogScreenshot.fLogScreenshot(message='Add Endpoint button is absent after adding 3 Endpoints',
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message='Add Endpoint button is present after adding 3 Endpoints',
                                              pass_=False, log=True, screenshot=False)
            raise Exception('Add Endpoint button is present after adding 3 Endpoints')

        # Read the field level error messages from all the fields
        actual_field_level_err_msg = []
        err_eles = self.select_elements("field_level_err_msgs", env)
        for i in err_eles:
            actual_field_level_err_msg.append(i.text)

        comparison_result = self.exbase.list_comparison_between_reports_data(expected_field_level_err_msg,
                                                                             actual_field_level_err_msg)

        if len(comparison_result) == 0:
            self.LogScreenshot.fLogScreenshot(message=f'Field Level Error Messages displayed correctly. Error '
                                                      f'messages are {actual_field_level_err_msg}',
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f'Mismatch found in Field Level Error Messages. Mismatch values '
                                                      f'are arranged in following order -> Expected Error Message, '
                                                      f'Actual Error Message. {comparison_result}',
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Field Level Error Messages")

        # Check whether Save and Download Master Extraction Template button is disabled when no details
        # are entered in fields
        if not self.isenabled("submit_button", env):
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is '
                                                      'disabled as expected', pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is not '
                                                      'disabled. Please recheck.',
                                              pass_=False, log=True, screenshot=False)
            raise Exception('Save and Download Master Extraction Template button is not disabled. Please recheck.')            
        
        # Click on Cancel button
        self.click("cancel_btn", env)

        # Fetching total rows count after clicking on cancel button
        table_rows_after = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                 table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after clicking on cancel button: '
                                                  f'{table_rows_after}',
                                          pass_=True, log=True, screenshot=False)        

        try:
            if table_rows_after == table_rows_before:
                self.LogScreenshot.fLogScreenshot(message=f'Table length is not increased after clicking on Cancel '
                                                          f'button as expected.',
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Table length is increased after clicking on Cancel button.',
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Table length is increased after clicking on Cancel button.")

            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Field Level Validation while adding new Non-Oncology "
                                                      f"population is completed***",
                                              pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Error while adding the population")

    def non_onocolgy_add_population(self, locatorname, add_locator, filepath, table_rows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Addition of new Non-Oncology population validation is "
                                                  f"started***", pass_=True, log=True, screenshot=False)

        expected_status_text = "Population added successfully"

        # Read required population and endpoint details
        pop_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_field',
                                                   'population_name')
        ep1_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep1_field', 'ep1_name')
        ep2_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep2_field', 'ep2_name')
        ep3_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep3_field', 'ep3_name')

        # Fetching total rows count before adding a new population
        table_rows_before = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                  table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new population: '
                                                  f'{table_rows_before}',
                                          pass_=True, log=True, screenshot=False)        

        self.scroll("managepopulation_page_heading", env)
        self.click(add_locator, env, UnivWaitFor=10)

        # Select Non-oncology tab
        self.click("non_oncology_tab", env)

        # Navigate through all fields without entering any values
        for i in pop_locs:
            self.input_text(i[0], i[1], env, UnivWaitFor=10)

        client_ele = self.select_element("non_oncology_clientid_drpdwn", env)
        select = Select(client_ele)
        select.select_by_visible_text('LIVEHTA Automation')
        sel_client_val = select.first_selected_option.text

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Name, Client and Description are :',
                                          pass_=True, log=True, screenshot=True)

        # Navigate through all fields without entering any values
        for j in ep1_locs:
            self.input_text(j[0], j[1], env, UnivWaitFor=10)
        
        self.click('ep1_type_cat', env)

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Endpoint1 :',
                                          pass_=True, log=True, screenshot=True)
        
        self.click("add_ep_btn", env)       
        
        # Navigate through all fields without entering any values
        for k in ep2_locs:
            self.input_text(k[0], k[1], env, UnivWaitFor=10)

        self.click('ep2_type_cont', env)

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Endpoint2 :',
                                          pass_=True, log=True, screenshot=True)
        
        self.click("add_ep_btn", env)

        # Navigate through all fields without entering any values
        for v in ep3_locs:
            self.input_text(v[0], v[1], env, UnivWaitFor=10)

        self.click('ep3_type_event', env)

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Endpoint3 :',
                                          pass_=True, log=True, screenshot=True)

        # Save and Download Master Extraction Template button should be enabled after entering all 3 endpoint details
        if self.isenabled("submit_button", env):
            self.click('submit_button', env)
            time.sleep(2)

            actual_status_text = self.get_status_text("population_status_popup_text", env)

            if actual_status_text == expected_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"Addition of New Non-Oncology Population is success",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding New "
                                                          f"Non-Oncology Population. Actual status message is "
                                                          f"{actual_status_text} and Expected status message is "
                                                          f"{expected_status_text}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Unable to find status message while adding New Non-Oncology Population")
            
            template_name = self.exbase.get_latest_filename(UnivWaitFor=180)
            if template_name == f"LIVEHTA Automation-{pop_locs[0][1]}-Master Template.xlsx":
                self.LogScreenshot.fLogScreenshot(message=f"Correct Template is downloaded. Tempalte name is "
                                                          f"{template_name}",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Downloaded Filename is {template_name}, Expectedname is "
                                                          f"LIVEHTA Automation-{pop_locs[0][1]}-Master Template",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Downloaded Filename is : {template_name}, Expectedname is : "
                                f"LIVEHTA Automation-{pop_locs[0][1]}-Master Template")
        else:
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is '
                                                      'disabled after entering all the mandatory details. Please '
                                                      'recheck.', pass_=False, log=True, screenshot=False)
            raise Exception('Save and Download Master Extraction Template button is disabled after entering all the '
                            'mandatory details. Please recheck.')

        # Fetching total rows count after adding a new population
        table_rows_after = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                 table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding a new population: '
                                                  f'{table_rows_after}',
                                          pass_=True, log=True, screenshot=False)        

        try:
            if table_rows_after > table_rows_before != table_rows_after:
                result = []
                self.scroll("managepopulation_page_heading", env)
                self.input_text("search_button", f'{pop_locs[0][1]}', env)
                td1 = self.select_elements('manage_pop_table_row_1', env)
                for m in td1:
                    result.append(m.text)

                self.LogScreenshot.fLogScreenshot(message=f'Table data after adding a new population: {result}',
                                                  pass_=True, log=True, screenshot=False)
                
                if result[0] == f'{pop_locs[0][1]}' and result[3] == f'{pop_locs[1][1]}':
                    self.LogScreenshot.fLogScreenshot(message=f'Population data is present in table',
                                                      pass_=True, log=True, screenshot=False)
                    population = f"{result[0]}"
                    return population, template_name
                else:
                    raise Exception("Population data is not added")
            self.clear("search_button", env)
            self.refreshpage()
            time.sleep(2)

            self.LogScreenshot.fLogScreenshot(message=f"***Addition of new Non-Oncology population validation is "
                                                      f"completed***", pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Error while adding the population")

    def non_onocolgy_add_duplicate_population(self, locatorname, add_locator, filepath, table_rows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Adding duplicate entry of existing Non-Oncology population "
                                                  f"validation is started***", pass_=True, log=True, screenshot=False)
        expected_status_text = "A population with the same name and unique company id already exists."

        # Read required population and endpoint details
        pop_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_field',
                                                   'population_name')
        ep1_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep1_field', 'ep1_name')
        ep2_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep2_field', 'ep2_name')
        ep3_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep3_field', 'ep3_name')

        # Fetching total rows count before adding a new population
        table_rows_before = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                  table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding a new population: '
                                                  f'{table_rows_before}',
                                          pass_=True, log=True, screenshot=False)        

        self.scroll("managepopulation_page_heading", env)
        self.click(add_locator, env, UnivWaitFor=10)

        # Select Non-oncology tab
        self.click("non_oncology_tab", env)

        # Navigate through all fields without entering any values
        for i in pop_locs:
            self.input_text(i[0], i[1], env, UnivWaitFor=10)

        client_ele = self.select_element("non_oncology_clientid_drpdwn", env)
        select = Select(client_ele)
        select.select_by_visible_text('LIVEHTA Automation')
        sel_client_val = select.first_selected_option.text

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Name, Client and Description are :',
                                          pass_=True, log=True, screenshot=True)

        # Navigate through all fields without entering any values
        for j in ep1_locs:
            self.input_text(j[0], j[1], env, UnivWaitFor=10)
        
        self.click('ep1_type_cat', env)

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Endpoint1 :',
                                          pass_=True, log=True, screenshot=True)
        
        self.click("add_ep_btn", env)       
        
        # Navigate through all fields without entering any values
        for k in ep2_locs:
            self.input_text(k[0], k[1], env, UnivWaitFor=10)

        self.click('ep2_type_cat', env)

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Endpoint2 :',
                                          pass_=True, log=True, screenshot=True)
        
        self.click("add_ep_btn", env)

        # Navigate through all fields without entering any values
        for v in ep3_locs:
            self.input_text(v[0], v[1], env, UnivWaitFor=10)

        self.click('ep3_type_cat', env)

        self.LogScreenshot.fLogScreenshot(message=f'Details entered for Endpoint3 :',
                                          pass_=True, log=True, screenshot=True)

        # Check whether Save and Download Master Extraction Template button is disabled when no details are entered
        # in fields
        if self.isenabled("submit_button", env):
            self.click('submit_button', env)
            time.sleep(2)

            actual_status_text = self.get_status_text("population_status_popup_text", env)

            if actual_status_text == expected_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"Non-Oncology Population '{pop_locs[0][1]}' is already "
                                                          f"present. Duplicate entry is not allowed.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding duplicate "
                                                          f"Non-Oncology Population. Actual status message is "
                                                          f"{actual_status_text} and Expected status message is "
                                                          f"{expected_status_text}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Unable to find status message while adding duplicate Non-Oncology Population")
        else:
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is '
                                                      'disabled after entering all the mandatory details. Please '
                                                      'recheck.', pass_=False, log=True, screenshot=False)
            raise Exception('Save and Download Master Extraction Template button is disabled after entering all the '
                            'mandatory details. Please recheck.')

        # Click on Cancel button
        self.click("cancel_btn", env)

        # Fetching total rows count after clicking on cancel button
        table_rows_after = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                 table_rows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after clicking on cancel button: '
                                                  f'{table_rows_after}',
                                          pass_=True, log=True, screenshot=False)        

        try:
            if table_rows_after == table_rows_before:
                self.LogScreenshot.fLogScreenshot(message=f'Table length is not increased while trying to add '
                                                          f'duplicate entry.', pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Table length is increased while trying to add duplicate '
                                                          f'entry.', pass_=False, log=True, screenshot=False)
                raise Exception("Table length is increased while trying to add duplicate entry.")

            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Adding duplicate entry of existing Non-Oncology population "
                                                      f"validation is completed***",
                                              pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Error while adding the population")

    def non_onocolgy_edit_population(self, locatorname, pop_name, edit_locator, filepath, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology population validation is started***",
                                          pass_=True, log=True, screenshot=False)
        expected_status_text = "Population updated successfully"
        # expected_status_text = "Project updated successfully"

        # Read required population and endpoint details
        pop_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_field',
                                                   'edit_population_name')

        self.input_text("search_button", f'{pop_name}', env)
        self.click(edit_locator, env, UnivWaitFor=10)

        for j in pop_locs:
            self.input_text(j[0], f'{j[1]}', env, UnivWaitFor=10)
        
        self.click("template_download_btn", env)
        time.sleep(2)
        template_name = self.exbase.get_latest_filename(UnivWaitFor=180)
            
        self.input_text("non_onco_template_file_upload", f'{os.getcwd()}//ActualOutputs//{template_name}', env)
        self.click("submit_button", env)
        time.sleep(1)

        # actual_status_text = self.get_text("population_status_popup_text", env, UnivWaitFor=10)
        actual_status_text = self.get_status_text("population_status_popup_text", env)
        
        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Editing the existing Non-Oncology Population is success",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while editing the "
                                                      f"Non-Oncology Population data. Actual status message is "
                                                      f"{actual_status_text} and Expected status message is "
                                                      f"{expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while editing the Non-Oncology Population data")

        try:
            result = []
            self.input_text("search_button", f'{pop_locs[0][1]}', env)
            td1 = self.select_elements('manage_pop_table_row_1', env)
            for m in td1:
                result.append(m.text)

            self.LogScreenshot.fLogScreenshot(message=f'Table data after editing the existing Non-Oncology population: '
                                                      f'{result}', pass_=True, log=True, screenshot=False)
            
            if result[0] == f'{pop_locs[0][1]}' and result[3] == f'{pop_locs[1][1]}':
                self.LogScreenshot.fLogScreenshot(message=f'Edited Non-Oncology Population data is present in table',
                                                  pass_=True, log=True, screenshot=False)
                population = f"{result[0]}"
                return population
            
            self.clear("search_button", env)
            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology population validation is completed***",
                                              pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Error while editing the population")

    def non_onocolgy_endpoint_details_validation(self, locatorname, filepath, template_name):
        self.LogScreenshot.fLogScreenshot(message=f"***Validation of Endpoint Details in Extraction Template for "
                                                  f"newly created Non-Oncology population is started***",
                                          pass_=True, log=True, screenshot=False)

        # Read required population and endpoint details
        ep1_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep1_field', 'ep1_name')
        ep2_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep2_field', 'ep2_name')
        ep3_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep3_field', 'ep3_name')

        self.LogScreenshot.fLogScreenshot(message=f"Extraction Tempalte name is {template_name}",
                                          pass_=True, log=True, screenshot=False)

        # Check the Endpoint details in extraction template
        template_data = openpyxl.load_workbook(f'ActualOutputs//{template_name}')
        template_sheet = template_data['Extraction sheet upload']

        # Validate Endpoint name and Endpoint abbreviation details
        if template_sheet['AN2'].value == ep1_locs[0][1] and template_sheet['AO2'].value == ep1_locs[1][1]:
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 1 details is present in Extraction template. "
                                                      f"Endpoint 1 Name: {template_sheet['AN2'].value} and "
                                                      f"Endpoint 1 Abbreviation: {template_sheet['AO2'].value}",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 1 details. "
                                                      f"Kindly recheck and try again. Actual Endpoint 1 Name: "
                                                      f"'{template_sheet['AN2'].value}', Actual Endpoint 1 "
                                                      f"Abbreviation: '{template_sheet['AO2'].value}' and Expected "
                                                      f"Endpoint 1 Name: '{ep1_locs[0][1]}', Expected Endpoint 1 "
                                                      f"Abbreviation: '{ep1_locs[1][1]}'.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 1 details. Kindly recheck and "
                            f"try again.")

        if template_sheet['BL2'].value == ep2_locs[0][1] and template_sheet['BM2'].value == ep2_locs[1][1]:
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 2 details is present in Extraction template. "
                                                      f"Endpoint 2 Name: {template_sheet['BL2'].value} and Endpoint 2 "
                                                      f"Abbreviation: {template_sheet['BM2'].value}",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 2 details. "
                                                      f"Kindly recheck and try again. Actual Endpoint 2 Name: "
                                                      f"{template_sheet['BL2'].value} and Actual Endpoint 2 "
                                                      f"Abbreviation: {template_sheet['BM2'].value} and Expected "
                                                      f"Endpoint 2 Name: '{ep2_locs[0][1]}', Expected Endpoint 2 "
                                                      f"Abbreviation: '{ep2_locs[1][1]}'.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 2 details. Kindly recheck and "
                            f"try again.")

        if template_sheet['CQ2'].value == ep3_locs[0][1] and template_sheet['CR2'].value == ep3_locs[1][1]:
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 3 details is present in Extraction template. "
                                                      f"Endpoint 3 Name: {template_sheet['CQ2'].value} and "
                                                      f"Endpoint 3 Abbreviation: {template_sheet['CR2'].value}",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 3 details. "
                                                      f"Kindly recheck and try again. Actual Endpoint 3 Name: "
                                                      f"{template_sheet['CQ2'].value} and Actual Endpoint 3 "
                                                      f"Abbreviation: {template_sheet['CR2'].value} and Expected "
                                                      f"Endpoint 3 Name: '{ep3_locs[0][1]}', Expected Endpoint 3 "
                                                      f"Abbreviation: '{ep3_locs[1][1]}'.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 3 details. Kindly recheck and "
                            f"try again.")

        # Validate Endpoint Type details
        if template_sheet['AY3'].value == "Endpoint Result-Categorical Variable-Within Arm" and \
                template_sheet['BD3'].value == "Endpoint Result-Categorical Variable-Between Arm":
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 1 Type details is present in Extraction template. "
                                                      f"Endpoint 1 Type Details: '{template_sheet['AY3'].value}' and "
                                                      f"'{template_sheet['BD3'].value}'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 1 Type "
                                                      f"details. Kindly recheck and try again. Actual Endpoint 1 "
                                                      f"Type Details: '{template_sheet['AY3'].value}' and "
                                                      f"'{template_sheet['BD3'].value}'. Expected Endpoint 1 Type "
                                                      f"Details: 'Endpoint Result-Categorical Variable-Within Arm' "
                                                      f"and 'Endpoint Result-Categorical Variable-Between Arm'.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 1 Type details. Kindly recheck and "
                            f"try again.")

        if template_sheet['BW3'].value == "Endpoint Result-Continuous Variable-Within Arm" and \
                template_sheet['CI3'].value == "Endpoint Result-Continuous Variable-Between Arm":
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 2 Type details is present in Extraction template. "
                                                      f"Endpoint 2 Type Details: '{template_sheet['BW3'].value}' and "
                                                      f"'{template_sheet['CI3'].value}'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 2 Type "
                                                      f"details. Kindly recheck and try again. Actual Endpoint 2 Type "
                                                      f"Details: '{template_sheet['BW3'].value}' and "
                                                      f"'{template_sheet['CI3'].value}'. Expected Endpoint 2 Type "
                                                      f"Details: 'Endpoint Result-Continuous Variable-Within Arm' and "
                                                      f"'Endpoint Result-Continuous Variable-Between Arm'.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 2 Type details. Kindly recheck and "
                            f"try again.")

        if template_sheet['DB3'].value == "Endpoint Result-Time-to-event Variable":
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 3 Type details is present in Extraction template. "
                                                      f"Endpoint 3 Type Details: '{template_sheet['DB3'].value}'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 3 Type "
                                                      f"details. Kindly recheck and try again. Actual Endpoint 3 "
                                                      f"Type Details: '{template_sheet['DB3'].value}'. Expected "
                                                      f"Endpoint 3 Type Details: 'Endpoint Result-Time-to-event "
                                                      f"Variable'.", pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 3 Type details. Kindly recheck and "
                            f"try again.")

        self.LogScreenshot.fLogScreenshot(message=f"***Validation of Endpoint Details in Extraction Template for "
                                                  f"newly created Non-Oncology population is completed***",
                                          pass_=True, log=True, screenshot=False)

    def verify_new_col_managepopulation_page(self, locatorname, filepath, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Validation of Newly added columns in Manage Population page "
                                                  f"is started***", pass_=True, log=True, screenshot=False)

        # Read required population and endpoint details
        testdata = self.exbase.get_triple_col_data(filepath, locatorname, 'Sheet1', 'population_name',
                                                   'indication_type', 'endpoints')

        for i in testdata:
            col_name = []
            col_eles = self.select_elements('manage_pop_table_col_names', env)
            for m in col_eles:
                col_name.append(m.text)
            
            if col_name[1] == 'Oncology/Non-oncology' and col_name[4] == 'Custom Endpoints':
                self.LogScreenshot.fLogScreenshot(message=f"'Oncology/Non-Oncology' column present in between Name "
                                                          f"and Unique ID columns. 'Custom Endpoints' column present "
                                                          f"in between Description and Last Update columns.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Newly added columns 'Oncology/Non-Oncology', 'Custom "
                                                          f"Endpoints' are not aligned in the desired position. "
                                                          f"Actual column names are {col_name}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Newly added columns 'Oncology/Non-Oncology', 'Custom Endpoints' are not aligned in "
                                f"the desired position. Kindy recheck the table column names under Manage "
                                f"Population page")
            
            result = []
            self.input_text("search_button", f'{i[0]}', env)
            td1 = self.select_elements('manage_pop_table_row_1', env)
            for n in td1:
                result.append(n.text)
            
            if result[1] == f'{i[1]}' and result[4] == f'{i[2]}':
                self.LogScreenshot.fLogScreenshot(message=f"For population -> '{i[0]}', values are displayed "
                                                          f"correctly. 'Oncology/Non-Oncology' -> {result[1]} and "
                                                          f"'Custom Endpoints' -> {result[4]}.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"For population -> '{i[0]}', values are not displayed "
                                                          f"correctly. 'Oncology/Non-Oncology' -> {result[1]} and "
                                                          f"'Custom Endpoints' -> {result[4]}.",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"For population -> '{i[0]}', values are not displayed correctly under newly added "
                                f"columns.")

            self.clear("search_button", env)
            self.refreshpage()
            time.sleep(2)
        
        self.LogScreenshot.fLogScreenshot(message=f"***Validation of Newly added columns in Manage Population page "
                                                  f"is completed***", pass_=True, log=True, screenshot=False)

    def non_onocolgy_edit_population_by_uploading_invalid_template(self, locatorname, pop_name, edit_locator, filepath,
                                                                   ep, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology population with invalid template validation "
                                                  f"for Endpoint {ep} -> {locatorname} is started***",
                                          pass_=True, log=True, screenshot=False)

        # Read expected error messages
        expected_err_msg = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'expected_error_msgs')
        
        # Read required population and endpoint details
        pop_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_field',
                                                   'edit_population_name')

        # Read template file
        invalid_template = self.exbase.get_template_file_details(filepath, locatorname, 'Files_to_upload')

        self.input_text("search_button", pop_name, env)
        self.click(edit_locator, env, UnivWaitFor=10)

        for j in pop_locs:
            self.input_text(j[0], f'{j[1]}', env, UnivWaitFor=10)

        self.input_text("non_onco_template_file_upload", f'{invalid_template}', env)
        self.click("submit_button", env)
        time.sleep(1)
        self.scrollback("non_onco_error_msg_heading", env)

        error_msg_heading = self.get_text('non_onco_error_msg_heading', env)
        error_msg_details = self.get_text('non_onco_error_msg_details', env)                                              

        actual_err_msg = []
        error_list_eles = self.select_elements('non_onco_list_of_errors', env)
        for k in error_list_eles:
            actual_err_msg.append(k.text)
        
        if error_msg_heading == 'Upload failed' and error_msg_details == f'There are {len(actual_err_msg)} columns ' \
                                                                         f'with errors.':
            self.LogScreenshot.fLogScreenshot(message=f"Error Message Heading and Error Message Information is "
                                                      f"displayed. Message Heading is '{error_msg_heading}' and "
                                                      f"Message Information is '{error_msg_details}'",
                                              pass_=True, log=True, screenshot=True)
            error_comparison_res = self.exbase.list_comparison_between_reports_data(expected_err_msg, actual_err_msg)
            if len(error_comparison_res) == 0:
                self.LogScreenshot.fLogScreenshot(message=f'Upload Error Messages displayed correctly. Error messages '
                                                          f'are {actual_err_msg}',
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Mismatch found in Upload Error Messages. Mismatch values '
                                                          f'are arranged in following order -> Expected Error Message,'
                                                          f' Actual Error Message. {error_comparison_res}',
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Mismatch found in Upload Error Messages")
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Error Message Heading and Error Message "
                                                      f"Information. Actual Message Heading is '{error_msg_heading}' "
                                                      f"and Actual Message Information is '{error_msg_details}'",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Mismatch found in Error Message Heading and Error Message Information")    

        self.click("cancel_btn", env)
        
        self.LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology population with invalid template validation "
                                                  f"for Endpoint {ep} -> {locatorname} is completed***",
                                          pass_=True, log=True, screenshot=False)

    def non_oncology_extraction_file_col_name_validation(self, locatorname, filepath, master_template):
        self.LogScreenshot.fLogScreenshot(message=f"***Non-Oncology Extraction file column IDs for Study "
                                                  f"Characteristics, Treatment Characteristics, Patient "
                                                  f"Characteristics, Clinical Study Characteristics validation "
                                                  f"is started***", pass_=True, log=True, screenshot=False)

        # Read expected column IDs
        e_stdy_char_cols = self.exbase.get_data_values(filepath, 'study_char_cols')
        e_treatment_char_cols = self.exbase.get_data_values(filepath, 'treatment_char_cols')
        e_patient_char_cols = self.exbase.get_data_values(filepath, 'patient_char_cols')
        e_clinical_study_char_cols = self.exbase.get_data_values(filepath, 'clinical_study_char_cols')
        e_ep1_cols = self.exbase.get_data_values(filepath, 'ep1_cols')
        e_ep2_cols = self.exbase.get_data_values(filepath, 'ep2_cols')
        e_ep3_cols = self.exbase.get_data_values(filepath, 'ep3_cols')
        e_other_clinical_eps = self.exbase.get_data_values(filepath, 'other_clinical_eps')
        e_clinical_safety_cols = self.exbase.get_data_values(filepath, 'clinical_safety_cols')
        e_clinical_safety_maping = self.exbase.get_data_values(filepath, 'clinical_safety_mapping')

        # Get master template file path
        master_file_path = f'ActualOutputs//{master_template}'

        # Read the actual column IDs
        a_stdy_char_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload',
                                         usecols="A:T")
        a_treatment_char_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload',
                                              usecols="U:X")
        a_patient_char_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload',
                                            usecols="Y:AJ")
        a_clinical_study_char_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload',
                                                   usecols="AK:AM")
        a_ep1_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload', usecols="AN:BK")
        a_ep2_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload', usecols="BL:CP")
        a_ep3_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload', usecols="CQ:DM")
        a_other_clinical_eps = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload',
                                             usecols="DN:DO")
        a_clinical_safety_cols = pd.read_excel(master_file_path, skiprows=3, sheet_name='Extraction sheet upload',
                                               usecols="DP:DV")
        a_clinical_safety_maping = pd.read_excel(master_file_path, skiprows=4, sheet_name='Extraction sheet upload',
                                                 usecols="DP:DV")

        # Compare the Expected and Actual Column IDs
        res1 = self.exbase.list_comparison_between_reports_data(e_stdy_char_cols, a_stdy_char_cols.columns.values)
        res2 = self.exbase.list_comparison_between_reports_data(e_treatment_char_cols,
                                                                a_treatment_char_cols.columns.values)
        res3 = self.exbase.list_comparison_between_reports_data(e_patient_char_cols,
                                                                a_patient_char_cols.columns.values)
        res4 = self.exbase.list_comparison_between_reports_data(e_clinical_study_char_cols,
                                                                a_clinical_study_char_cols.columns.values)
        res5 = self.exbase.list_comparison_between_reports_data(e_ep1_cols, a_ep1_cols.columns.values)
        res6 = self.exbase.list_comparison_between_reports_data(e_ep2_cols, a_ep2_cols.columns.values)
        res7 = self.exbase.list_comparison_between_reports_data(e_ep3_cols, a_ep3_cols.columns.values)
        res8 = self.exbase.list_comparison_between_reports_data(e_other_clinical_eps,
                                                                a_other_clinical_eps.columns.values)
        res9 = self.exbase.list_comparison_between_reports_data(e_clinical_safety_cols,
                                                                a_clinical_safety_cols.columns.values)
        res10 = self.exbase.list_comparison_between_reports_data(e_clinical_safety_maping,
                                                                 a_clinical_safety_maping.columns.values)

        res = {'Study Characteristics': [res1, a_stdy_char_cols.columns.values],
               'Treatment Characteristics': [res2, a_treatment_char_cols.columns.values],
               'Patient Characteristics': [res3, a_patient_char_cols.columns.values],
               'Clinical Study Characteristics': [res4, a_clinical_study_char_cols.columns.values],
               'Endpoint 1': [res5, a_ep1_cols.columns.values],
               'Endpoint 2': [res6, a_ep2_cols.columns.values],
               'Endpoint 3': [res7, a_ep3_cols.columns.values],
               'Other Clinical Endpoint': [res8, a_other_clinical_eps.columns.values],
               'Clinical Saftey Column IDs': [res9, a_clinical_safety_cols.columns.values],
               'Clinical Saftey Column Mapping': [res10, a_clinical_safety_maping.columns.values]}

        for k, v in res.items():
            if len(v[0]) == 0:
                self.LogScreenshot.fLogScreenshot(message=f"'{k}' column IDs are present in Extraction Template. "
                                                          f"IDs are : '{v[1]}'",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in '{k}' column IDs. Mismatch values are "
                                                          f"arranged in following order -> Expected Column IDs, "
                                                          f"Actual Column IDs. {v[0]}",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Mismatch found in '{k}' column IDs")

        self.LogScreenshot.fLogScreenshot(message=f"***Non-Oncology Extraction file column IDs for Study "
                                                  f"Characteristics, Treatment Characteristics, "
                                                  f"Patient Characteristics, Clinical Study Characteristics "
                                                  f"validation is completed***",
                                          pass_=True, log=True, screenshot=False)
