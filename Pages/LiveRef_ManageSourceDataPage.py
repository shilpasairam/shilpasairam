from datetime import date
import os
import random
from re import search
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


class ManageSourceDataPage(Base):

    """Constructor of the ManageSourceData Page class"""
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

    def go_to_managesourcedata(self, locator, env):
        self.click(locator, env, UnivWaitFor=10)
        time.sleep(5)
    
    def get_details(self, locatorname, filepath, column_locator, column_name):
        df = pd.read_excel(filepath)
        data = df.loc[df['Name'] == locatorname][column_locator].dropna().to_list()
        value = df.loc[df['Name'] == locatorname][column_name].dropna().to_list()
        result = [(data[i], value[i]) for i in range(0, len(data))]
        return result, value
    
    def get_file_details(self, locatorname, filepath, logo_column_name, temp_column_name):
        df = pd.read_excel(filepath)
        logo_file_path = os.getcwd()+(df.loc[df['Name'] == locatorname][logo_column_name].to_list()[0])
        template_file_path = os.getcwd()+(df.loc[df['Name'] == locatorname][temp_column_name].to_list()[0])
        return logo_file_path, template_file_path
    
    def add_invalid_managesourcedata(self, locatorname, filepath, env):
        expected_error_msg = "A problem occurred while uploading the file. The columns DateTime_Location, " \
                             "Author_LastName, Main_Message, Study Type/GVD Chapter, Study_Sub_Type/GVD Section, " \
                             "Reported_Data_Variables, Scales, Main_Results are not valid for this template"
        # Read manage source details from data sheet
        source_data, source_value = self.get_details(locatorname, filepath, 'Add_source_field', 'Add_source_value')

        # Read filepaths to upload
        logo_path, template_path = self.get_file_details(locatorname, filepath, 'Source_Logo', 'Source_Template')

        self.click("add_sourcedata_btn", env, UnivWaitFor=10)
        time.sleep(1)

        for j in source_data:
            self.input_text(j[0], f'{j[1]}', env, UnivWaitFor=10)
            
        self.input_text("source_logo_file", logo_path, env)
        self.input_text("source_template_file", template_path, env)
        time.sleep(1)
        self.LogScreenshot.fLogScreenshot(message=f"Entered Source Details are: ",
                                          pass_=True, log=True, screenshot=True)

        self.click("source_save_button", env)
        time.sleep(3)

        actual_err_text = self.get_text("source_invaliddata_msg", env)
        
        if search(expected_error_msg, actual_err_text):
            self.LogScreenshot.fLogScreenshot(message=f"Error Message is displayed as expected. Error messgae is : "
                                                      f"{actual_err_text}", pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Different Error Message has occurred. Error message is : "
                                                      f"{actual_err_text}", pass_=False, log=True, screenshot=True)
            raise Exception(f"Different Error Message has occurred.")

    def add_valid_managesourcedata(self, locatorname, filepath, tablerows, env):
        expected_msg = "Record added successfully"
        # Read manage source details from data sheet
        source_data, source_value = self.get_details(locatorname, filepath, 'Add_source_field', 'Add_source_value')

        # Read filepaths to upload
        logo_path, template_path = self.get_file_details(locatorname, filepath, 'Source_Logo', 'Source_Template')

        ele = self.select_element("sel_table_entries_dropdown", env)
        select = Select(ele)
        select.select_by_visible_text("All")

        # Fetching total rows count before adding a new manage source data
        table_rows_before = self.select_elements(tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before adding Manage Source Data: '
                                                  f'{len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.click("add_sourcedata_btn", env, UnivWaitFor=10)
        time.sleep(1)

        for j in source_data:
            self.input_text(j[0], f'{j[1]}', env, UnivWaitFor=10)
            
        self.input_text("source_logo_file", logo_path, env)
        self.input_text("source_template_file", template_path, env)
        time.sleep(1)
        self.LogScreenshot.fLogScreenshot(message=f"Entered Source Details are: ",
                                          pass_=True, log=True, screenshot=True)

        self.click("source_save_button", env)

        actual_msg = self.get_text("sourcedata_add_success_msg", env, UnivWaitFor=10)
        
        if search(expected_msg, actual_msg):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of Data is successfully added",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Error while adding Manage Source of Data",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Error while adding Manage Source of Data")
        
        ele = self.select_element("sel_table_entries_dropdown", env)
        select = Select(ele)
        select.select_by_visible_text("All")

        # Fetching total rows count after adding a new manage source data
        table_rows_after = self.select_elements(tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after adding Manage Source Data: '
                                                  f'{len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                result = []
                self.input_text("sourcedata_search_box", f'{source_value[1]}_{date.today().year}', env)
                # self.LogScreenshot.fLogScreenshot(message=f'Table data after adding a new manage source of data : ',
                #                     pass_=True, log=True, screenshot=True)
                td1 = self.select_elements('sourcedata_table_row_1', env)
                for m in td1:
                    result.append(m.text)
                
                if result[2] == f'{source_value[1]}_{date.today().year}':
                    self.LogScreenshot.fLogScreenshot(message=f'Newly added Manage Source data is present in table',
                                                      pass_=True, log=True, screenshot=True)
                    source_code = f"{result[2]}"
                    return source_code
                else:
                    raise Exception("Manage Source data is not added")
            self.clear("sourcedata_search_box", env)
            self.refreshpage()
            time.sleep(2)
        except Exception:
            raise Exception("Error while adding the manage source data")

    def edit_valid_managesourcedata(self, locatorname, src_code, filepath, edit_locator, env):
        expected_msg = "Record updated successfully"
        self.refreshpage()
        time.sleep(2)        
        # Read manage source details from data sheet
        edit_source_data, edit_source_value = self.get_details(locatorname, filepath, 'Add_source_field',
                                                               'Edit_source_value')

        # Read filepaths to upload
        logo_path, template_path = self.get_file_details(locatorname, filepath, 'Source_Logo', 'Source_Template')

        self.input_text("sourcedata_search_box", f'{src_code}', env)
        self.LogScreenshot.fLogScreenshot(message=f'Selected record details for updation : ',
                                          pass_=True, log=True, screenshot=True)

        self.click(edit_locator, env, UnivWaitFor=10)
        time.sleep(1)

        for j in edit_source_data:
            self.input_text(j[0], f'{j[1]}', env, UnivWaitFor=10)
            
        self.input_text("source_logo_file", logo_path, env)
        self.input_text("source_template_file", template_path, env)
        time.sleep(1)
        self.LogScreenshot.fLogScreenshot(message=f"Entered Source Details are: ",
                                          pass_=True, log=True, screenshot=True)

        self.click("source_save_button", env)

        actual_msg = self.get_text("sourcedata_edit_success_msg", env, UnivWaitFor=10)
        
        if search(expected_msg, actual_msg):
            self.LogScreenshot.fLogScreenshot(message=f"Manage Source of Data is successfully updated",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Error while updating Manage Source of Data",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Error while updating Manage Source of Data")

        try:
            result = []
            self.input_text("sourcedata_search_box", f'{edit_source_value[1]}_{date.today().year}', env)
            # self.LogScreenshot.fLogScreenshot(message=f'Table data after updating manage source of data : ',
            #                         pass_=True, log=True, screenshot=True)
            td1 = self.select_elements('sourcedata_table_row_1', env)
            for m in td1:
                result.append(m.text)
            
            if result[2] == f'{edit_source_value[1]}_{date.today().year}':
                self.LogScreenshot.fLogScreenshot(message=f'Updated Manage Source data is present in table',
                                                  pass_=True, log=True, screenshot=True)
                source_code = f"{result[2]}"
                return source_code

            self.clear("sourcedata_search_box", env)
            self.refreshpage()
            time.sleep(2)
        except Exception:
            raise Exception("Error while updating the manage source data")

    def delete_managesourcedata(self, src_code, tablerows, env):
        expected_status_text = "Source successfully deleted"
        ele = self.select_element("sel_table_entries_dropdown", env)
        select = Select(ele)
        select.select_by_visible_text("All")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_before = self.select_elements(tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a manage source data: '
                                                  f'{len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)

        self.input_text("sourcedata_search_box", src_code, env)
        self.LogScreenshot.fLogScreenshot(message=f'Selected record details for deletion : ',
                                          pass_=True, log=True, screenshot=True)
        
        self.click("sourcedata_delete", env)
        time.sleep(1)
        self.click("sourcedata_delete_popup_confirm", env)
        time.sleep(1)
        
        actual_status_text = self.get_text("sourcedata_del_success_msg", env, UnivWaitFor=10)
        self.click("sourcedata_del_success_msg_ok", env)

        if actual_status_text == expected_status_text:
            self.LogScreenshot.fLogScreenshot(message=f"Deleting the existing Manage source data is success",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting the "
                                                      f"Manage source data",
                                              pass_=False, log=True, screenshot=True)
            raise Exception(f"Unable to find status message while deleting the Manage source data")

        ele = self.select_element("sel_table_entries_dropdown", env)
        select = Select(ele)
        select.select_by_visible_text("All")

        # Fetching total rows count before deleting a file from top of the table
        table_rows_after = self.select_elements(tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a manage source data: '
                                                  f'{len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                                  pass_=True, log=True, screenshot=False)
            self.refreshpage()
            time.sleep(2)                
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error in deleting the manage source data")
