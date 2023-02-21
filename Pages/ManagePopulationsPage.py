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
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding New Population. Actual status message is {actual_status_text} and Expected status message is {expected_status_text}",
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
        except Exception:
            raise Exception("Error while adding the population")

    def edit_multiple_population(self, locatorname, pop_name, edit_locator, filepath, env):
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
                                                      f"Population data. Actual status message is {actual_status_text} and Expected status message is {expected_status_text}",
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
        except Exception:
            raise Exception("Error while editing the population")

    def delete_multiple_population(self, pop_value, del_locator, del_locator_popup, tablerows, env):
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
        
        self.click(del_locator, env)
        time.sleep(2)
        self.click(del_locator_popup, env)
        time.sleep(2)
        
        # actual_status_text = self.get_text("population_status_popup_text", env, UnivWaitFor=10)
        actual_status_text = self.get_status_text("population_status_popup_text", env)
        # time.sleep(2)

        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Deleting the existing Population is success",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting the "
                                                      f"Population data. Actual status message is {actual_status_text} and Expected status message is {expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while deleting the Population data")

        # Fetching total rows count after deleting a new population
        table_rows_after = self.get_table_length("manage_pop_table_rows_info", "manage_pop_table_next_btn",
                                                 tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a population: {table_rows_after}',
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
            raise Exception("Error in deleting the population")

    def non_onocolgy_check_field_level_err_msg(self, locatorname, add_locator, filepath, table_rows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Field Level Validation while adding new Non-Oncology population is started***",
                                        pass_=True, log=True, screenshot=False)

        # Read required population and endpoint details
        field_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'population_field')
        ep1_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'ep1_field')
        ep2_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'ep2_field')
        ep3_locs = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'ep3_field')
        expected_field_level_err_msg = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'expected_filed_level_err_msg')

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
            self.LogScreenshot.fLogScreenshot(message='Add Endpoint button is absent after adding 3 Endpoints', pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message='Add Endpoint button is present after adding 3 Endpoints', pass_=False, log=True, screenshot=False)
            raise Exception('Add Endpoint button is present after adding 3 Endpoints')

        # Read the field level error messages from all the fields
        actual_field_level_err_msg = []
        err_eles = self.select_elements("field_level_err_msgs", env)
        for i in err_eles:
            actual_field_level_err_msg.append(i.text)

        comparison_result = self.exbase.list_comparison_between_reports_data(expected_field_level_err_msg, actual_field_level_err_msg)

        if len(comparison_result) == 0:
            self.LogScreenshot.fLogScreenshot(message=f'Field Level Error Messages displayed correctly. Error messages are {actual_field_level_err_msg}',
                                            pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f'Mismatch found in Field Level Error Messages. Mismatch values are arranged in following order -> Actual Error Message, Expected Error Message. {comparison_result}',
                                            pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Field Level Error Messages")

        # Check whether Save and Download Master Extraction Template button is disabled when no details are entered in fields
        if not self.isenabled("submit_button", env):
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is disabled as expected', pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is not disabled. Please recheck.', pass_=False, log=True, screenshot=False)
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
                self.LogScreenshot.fLogScreenshot(message=f'Table length is not increased after clicking on Cancel button as expected.',
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Table length is increased after clicking on Cancel button.',
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Table length is increased after clicking on Cancel button.")

            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Field Level Validation while adding new Non-Oncology population is completed***",
                                            pass_=True, log=True, screenshot=False)            
        except Exception:
            raise Exception("Error while adding the population")

    def non_onocolgy_add_population(self, locatorname, add_locator, filepath, table_rows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Addition of new Non-Oncology population is started***",
                                        pass_=True, log=True, screenshot=False)

        expected_status_text = "Population added successfully"

        # Read required population and endpoint details
        pop_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_field', 'population_name')
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

        # Check whether Save and Download Master Extraction Template button is disabled when no details are entered in fields
        if self.isenabled("submit_button", env):
            self.click('submit_button', env)
            time.sleep(2)

            actual_status_text = self.get_status_text("population_status_popup_text", env)

            if actual_status_text == expected_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"Addition of New Non-Oncology Population is success",
                                                    pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding New Non-Oncology Population. Actual status message is {actual_status_text} and Expected status message is {expected_status_text}",
                                                    pass_=False, log=True, screenshot=True)
                raise Exception(f"Unable to find status message while adding New Non-Oncology Population")
            
            template_name = self.exbase.get_latest_filename(UnivWaitFor=180)
            if template_name == f"LIVEHTA Automation-{pop_locs[0][1]}-Master Template.xlsx":
                self.LogScreenshot.fLogScreenshot(message=f"Correct Template is downloaded. Tempalte name is {template_name}",
                                pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Downloaded Filename is {template_name}, Expectedname is "
                                            f"LIVEHTA Automation-{pop_locs[0][1]}-Master Template",
                                    pass_=False, log=True, screenshot=False)           
        else:
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is disabled after entering all the mandatory details. Please recheck.', pass_=False, log=True, screenshot=False)
            raise Exception('Save and Download Master Extraction Template button is disabled after entering all the mandatory details. Please recheck.')

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

            self.LogScreenshot.fLogScreenshot(message=f"***Addition of new Non-Oncology population is completed***",
                                            pass_=True, log=True, screenshot=False)            
        except Exception:
            raise Exception("Error while adding the population")

    def non_onocolgy_add_duplicate_population(self, locatorname, add_locator, filepath, table_rows, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Adding duplicate entry of existing Non-Oncology population is started***",
                                        pass_=True, log=True, screenshot=False)
        expected_status_text = "A population with the same name and unique company id already exists."

        # Read required population and endpoint details
        pop_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_field', 'population_name')
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

        # Check whether Save and Download Master Extraction Template button is disabled when no details are entered in fields
        if self.isenabled("submit_button", env):
            self.click('submit_button', env)
            time.sleep(2)

            actual_status_text = self.get_status_text("population_status_popup_text", env)

            if actual_status_text == expected_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"Non-Oncology Population '{pop_locs[0][1]}' is already present. Duplicate entry is not allowed.",
                                                    pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding duplicate Non-Oncology Population. Actual status message is {actual_status_text} and Expected status message is {expected_status_text}",
                                                    pass_=False, log=True, screenshot=True)
                raise Exception(f"Unable to find status message while adding duplicate Non-Oncology Population")
        else:
            self.LogScreenshot.fLogScreenshot(message='Save and Download Master Extraction Template button is disabled after entering all the mandatory details. Please recheck.', pass_=False, log=True, screenshot=False)
            raise Exception('Save and Download Master Extraction Template button is disabled after entering all the mandatory details. Please recheck.')

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
                self.LogScreenshot.fLogScreenshot(message=f'Table length is not increased while trying to add duplicate entry.',
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Table length is increased while trying to add duplicate entry.',
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Table length is increased while trying to add duplicate entry.")

            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Adding duplicate entry of existing Non-Oncology population is completed***",
                                        pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Error while adding the population")

    def non_onocolgy_edit_population(self, locatorname, pop_name, edit_locator, filepath, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology population is started***",
                                        pass_=True, log=True, screenshot=False)        
        expected_status_text = "Population updated successfully"
        # expected_status_text = "Project updated successfully"

        # Read required population and endpoint details
        pop_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'population_field', 'edit_population_name')               

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
                                                      f"Non-Oncology Population data. Actual status message is {actual_status_text} and Expected status message is {expected_status_text}",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while editing the Non-Oncology Population data")

        try:
            result = []
            self.input_text("search_button", f'{pop_locs[0][1]}', env)
            td1 = self.select_elements('manage_pop_table_row_1', env)
            for m in td1:
                result.append(m.text)

            self.LogScreenshot.fLogScreenshot(message=f'Table data after editing the existing Non-Oncology population: {result}',
                                              pass_=True, log=True, screenshot=False)
            
            if result[0] == f'{pop_locs[0][1]}' and result[3] == f'{pop_locs[1][1]}':
                self.LogScreenshot.fLogScreenshot(message=f'Edited Non-Oncology Population data is present in table',
                                                  pass_=True, log=True, screenshot=False)
                population = f"{result[0]}"
                return population
            
            self.clear("search_button", env)
            self.refreshpage()
            time.sleep(2)
            self.LogScreenshot.fLogScreenshot(message=f"***Edit Non-Oncology population is completed***",
                                            pass_=True, log=True, screenshot=False)             
        except Exception:
            raise Exception("Error while editing the population")

    def non_onocolgy_endpoint_details_validation(self, locatorname, filepath, template_name):
        self.LogScreenshot.fLogScreenshot(message=f"***Validation of Endpoint Details in Extraction Template for newly created Non-Oncology population is started***",
                                        pass_=True, log=True, screenshot=False)

        # Read required population and endpoint details
        ep1_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep1_field', 'ep1_name')
        ep2_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep2_field', 'ep2_name')
        ep3_locs = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'ep3_field', 'ep3_name')

        self.LogScreenshot.fLogScreenshot(message=f"Extraction Tempalte name is {template_name}", pass_=True, log=True, screenshot=False)

        # Check the Endpoint details in extraction template
        template_data = openpyxl.load_workbook(f'ActualOutputs//{template_name}')
        template_sheet = template_data['Extraction sheet upload']

        # Validate Endpoint name and Endpoint abbreviation details
        if template_sheet['AN2'].value == ep1_locs[0][1] and template_sheet['AO2'].value == ep1_locs[1][1]:
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 1 details is present in Extraction template. Endpoint 1 Name: {template_sheet['AN2'].value} and Endpoint 1 Abbreviation: {template_sheet['AO2'].value}",
                                            pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 1 details. Kindly recheck and try again. Actual Endpoint 1 Name: '{template_sheet['AN2'].value}', Actual Endpoint 1 Abbreviation: '{template_sheet['AO2'].value}' and Expected Endpoint 1 Name: '{ep1_locs[0][1]}', Expected Endpoint 1 Abbreviation: '{ep1_locs[1][1]}'.",
                                            pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 1 details. Kindly recheck and try again.")

        if template_sheet['BL2'].value == ep2_locs[0][1] and template_sheet['BM2'].value == ep2_locs[1][1]:
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 2 details is present in Extraction template. Endpoint 2 Name: {template_sheet['BL2'].value} and Endpoint 2 Abbreviation: {template_sheet['BM2'].value}",
                                            pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 2 details. Kindly recheck and try again. Actual Endpoint 2 Name: {template_sheet['BL2'].value} and Actual Endpoint 2 Abbreviation: {template_sheet['BM2'].value} and Expected Endpoint 2 Name: '{ep2_locs[0][1]}', Expected Endpoint 2 Abbreviation: '{ep2_locs[1][1]}'.",
                                            pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 2 details. Kindly recheck and try again.")

        if template_sheet['CQ2'].value == ep3_locs[0][1] and template_sheet['CR2'].value == ep3_locs[1][1]:
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 3 details is present in Extraction template. Endpoint 3 Name: {template_sheet['CQ2'].value} and Endpoint 3 Abbreviation: {template_sheet['CR2'].value}",
                                            pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 3 details. Kindly recheck and try again. Actual Endpoint 3 Name: {template_sheet['CQ2'].value} and Actual Endpoint 3 Abbreviation: {template_sheet['CR2'].value} and Expected Endpoint 3 Name: '{ep3_locs[0][1]}', Expected Endpoint 3 Abbreviation: '{ep3_locs[1][1]}'.",
                                            pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 3 details. Kindly recheck and try again.")

        # Validate Endpoint Type details
        if template_sheet['AY3'].value == "Endpoint Result-Categorical Variable-Within Arm" and template_sheet['BD3'].value == "Endpoint Result-Categorical Variable-Between Arm":
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 1 Type details is present in Extraction template. Endpoint 1 Type Details: '{template_sheet['AY3'].value}' and '{template_sheet['BD3'].value}'",
                                            pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 1 Type details. Kindly recheck and try again. Actual Endpoint 1 Type Details: '{template_sheet['AY3'].value}' and '{template_sheet['BD3'].value}'. Expected Endpoint 1 Type Details: 'Endpoint Result-Categorical Variable-Within Arm' and 'Endpoint Result-Categorical Variable-Between Arm'.",
                                            pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 1 Type details. Kindly recheck and try again.")

        if template_sheet['BW3'].value == "Endpoint Result-Continuous Variable-Within Arm" and template_sheet['CI3'].value == "Endpoint Result-Continuous Variable-Between Arm":
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 2 Type details is present in Extraction template. Endpoint 2 Type Details: '{template_sheet['BW3'].value}' and '{template_sheet['CI3'].value}'",
                                            pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 2 Type details. Kindly recheck and try again. Actual Endpoint 2 Type Details: '{template_sheet['BW3'].value}' and '{template_sheet['CI3'].value}'. Expected Endpoint 2 Type Details: 'Endpoint Result-Continuous Variable-Within Arm' and 'Endpoint Result-Continuous Variable-Between Arm'.",
                                            pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 2 Type details. Kindly recheck and try again.")

        if template_sheet['DB3'].value == "Endpoint Result-Time-to-event Variable":
            self.LogScreenshot.fLogScreenshot(message=f"Endpoint 3 Type details is present in Extraction template. Endpoint 3 Type Details: '{template_sheet['DB3'].value}'",
                                            pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Extraction template -> Endpoint 3 Type details. Kindly recheck and try again. Actual Endpoint 3 Type Details: '{template_sheet['DB3'].value}'. Expected Endpoint 3 Type Details: 'Endpoint Result-Time-to-event Variable'.",
                                            pass_=False, log=True, screenshot=False)
            raise Exception(f"Mismatch found in Extraction template -> Endpoint 3 Type details. Kindly recheck and try again.")            

        self.LogScreenshot.fLogScreenshot(message=f"***Validation of Endpoint Details in Extraction Template for newly created Non-Oncology population is completed***",
                                        pass_=True, log=True, screenshot=False)
