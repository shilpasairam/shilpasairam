import os
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class ImportPublicationPage(Base):

    """Constructor of the ImportPublication Page class"""
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

    # def go_to_importpublications(self, locator, button, env):
    #     self.click(locator, env, UnivWaitFor=10)
    #     self.jsclick(button, env)
    #     time.sleep(3)

    # Reading Population data for Excluded Studies Page
    def get_file_details_to_upload(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop_name = df.loc[df['Name'] == locatorname]['Population_name'].dropna().to_list()
        path = df.loc[df['Name'] == locatorname]['Files_to_upload'].dropna().to_list()
        filename = df.loc[df['Name'] == locatorname]['Expected_File_names'].dropna().to_list()
        result = [[pop_name[i], os.getcwd() + path[i], filename[i]] for i in range(0, len(pop_name))]
        return result
    
    def upload_file(self, locatorname, filepath, index):
        expected_upload_status_text = "File(s) uploaded successfully"
        # Read population details from data sheet
        pop_data = self.get_file_details_to_upload(filepath, locatorname)

        for i in pop_data:
            ele = self.select_element("select_update_dropdown")
            time.sleep(2)
            select = Select(ele)
            select.select_by_visible_text(i[0])

            # Fetching total rows count before uploading a new file
            table_rows_before = self.select_elements("upload_table_rows")
            self.LogScreenshot.fLogScreenshot(message=f'Table length before uploading a new file: '
                                                      f'{len(table_rows_before)}',
                                              pass_=True, log=True, screenshot=False)
            
            jscmd = ReadConfig.get_remove_att_JScommand(16, 'hidden')
            self.jsclick_hide(jscmd)
            self.input_text("add_file", i[1])
            try:
                self.jsclick("upload_button")
                time.sleep(3)
                actual_upload_status_text = self.get_text("file_status_popup_text", UnivWaitFor=30)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'File upload is success for Population : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading '
                                                              f'Extraction File for Population : {i[0]}.',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message during Extraction file uploading")

                # Fetching total rows count after uploading a new file
                table_rows_after = self.select_elements("upload_table_rows")
                self.LogScreenshot.fLogScreenshot(message=f'Table length after uploading a new file: '
                                                          f'{len(table_rows_after)}',
                                                  pass_=True, log=True, screenshot=False)

                if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                    result = []
                    td1 = self.select_elements('upload_table_row_1')
                    for m in td1:
                        result.append(m.text)
                    
                    if i[2] in result:
                        self.LogScreenshot.fLogScreenshot(message=f'Correct file with expected filename is being '
                                                                  f'uploaded: {i[2]}',
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        raise Exception("Wrong file is uploaded")

                # Validating the upload status icon
                status_icon = ["file_upload_status_pass", "file_upload_status_failure"]
                if status_icon[index] == "file_upload_status_pass":
                    time.sleep(10)
                    if self.isdisplayed("file_upload_status_pass", UnivWaitFor=180):
                        self.LogScreenshot.fLogScreenshot(message=f'File uploading is done with Success Icon',
                                                          pass_=True, log=True, screenshot=True)
                    else:
                        raise Exception("Error while uploading the extraction file")
                elif status_icon[index] == "file_upload_status_failure":
                    time.sleep(10)
                    if self.isdisplayed("file_upload_status_failure", UnivWaitFor=180):
                        self.LogScreenshot.fLogScreenshot(message=f'File uploading is done with Failure Icon',
                                                          pass_=True, log=True, screenshot=True)
                        self.click("view_action", UnivWaitFor=10)
                        time.sleep(2)
                        td = self.select_elements('error_data_table')
                        error_data = []
                        for n in td:
                            error_data.append(n.text)
                        self.LogScreenshot.fLogScreenshot(message=f'Excel sheet contains the following errors: '
                                                                  f'{error_data}',
                                                          pass_=True, log=True, screenshot=True)
                        self.click("back_to_view_action_btn", UnivWaitFor=10)
                    else:
                        raise Exception("Error while uploading the extraction file")
                self.refreshpage()
                time.sleep(5)
            except Exception:
                raise Exception("Error while uploading")

    def delete_file(self, locatorname, filepath, msg_popup, tablerows, env):
        expected_delete_status_text = "Import status deleted successfully"
        # Read population details from data sheet
        pop_data = self.get_file_details_to_upload(filepath, locatorname)

        # Fetching total rows count before deleting a file from top of the table
        table_rows_before = self.select_elements(tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a file: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)
        
        self.refreshpage()
        time.sleep(5)

        for i in pop_data:
            result = []
            td1 = self.select_elements('upload_table_row_1', env)
            for m in td1:
                result.append(m.text)
            
            # Check the uploaded filename before deleting the record
            if i[2] in result:
                self.LogScreenshot.fLogScreenshot(message=f"Uploaded Filename '{i[2]}' is present in the table. "
                                                          f"Performing the delete operation.",
                                                  pass_=True, log=True, screenshot=True)
            
                self.click("delete_file", env)
                time.sleep(2)
                self.click("delete_file_popup", env)
                time.sleep(2)

                # actual_delete_status_text = self.get_text(msg_popup, env, UnivWaitFor=30)
                actual_delete_status_text = self.get_status_text(msg_popup, env)
                
                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'Extraction File Deletion is success.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while deleting '
                                                              f'Extraction File. Actual status message is {actual_delete_status_text} and Expected status message is {expected_delete_status_text}',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Error during Extraction File Deletion")

                # Fetching total rows count before deleting a file from top of the table
                table_rows_after = self.select_elements(tablerows, env)
                self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a file: '
                                                          f'{len(table_rows_after)}',
                                                  pass_=True, log=True, screenshot=False)

                try:
                    if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                        self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                                          pass_=True, log=True, screenshot=False)
                except Exception:
                    self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception("Error in deleting the imported file")
            else:
                raise Exception("No file uploaded to perform delete operation")

    def upload_file_with_errors(self, locatorname, filepath, env):
        expected_upload_status_text = "File(s) uploaded successfully"
        # Read population details from data sheet
        pop_data = self.get_file_details_to_upload(filepath, locatorname)

        # Read expected error messages data
        col1_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'error_msg_col1')
        # Converting list values from float to int
        col1_data = [int(x) for x in col1_data]

        col2_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'error_msg_col2')

        expected_err_msg = [[col1_data[i], col2_data[i]] for i in range(0, len(col1_data))]        

        for i in pop_data:
            ele = self.select_element("select_update_dropdown", env)
            time.sleep(2)
            select = Select(ele)
            select.select_by_visible_text(i[0])

            # Fetching total rows count before uploading a new file
            table_rows_before = self.select_elements("upload_table_rows", env)
            self.LogScreenshot.fLogScreenshot(message=f'Table length before uploading a new file: '
                                                      f'{len(table_rows_before)}',
                                              pass_=True, log=True, screenshot=False)
            
            jscmd = ReadConfig.get_remove_att_JScommand(16, 'hidden')
            self.jsclick_hide(jscmd)
            self.input_text("add_file", i[1], env)
            try:
                self.jsclick("upload_button", env)
                time.sleep(2)
                # actual_upload_status_text = self.get_text("file_status_popup_text", env, UnivWaitFor=30)
                actual_upload_status_text = self.get_status_text("file_status_popup_text", env)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'File upload is success for Population : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading '
                                                              f'Extraction File for Population : {i[0]}. Actual status message is {actual_upload_status_text} and Expected status message is {expected_upload_status_text}',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message during Extraction file uploading")

                # Fetching total rows count after uploading a new file
                table_rows_after = self.select_elements("upload_table_rows", env)
                self.LogScreenshot.fLogScreenshot(message=f'Table length after uploading a new file: '
                                                          f'{len(table_rows_after)}',
                                                  pass_=True, log=True, screenshot=False)

                if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                    result = []
                    td1 = self.select_elements('upload_table_row_1', env)
                    for m in td1:
                        result.append(m.text)
                    
                    if i[2] in result:
                        self.LogScreenshot.fLogScreenshot(message=f'Correct file with expected filename is '
                                                                  f'being uploaded: {i[2]}',
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        raise Exception("Wrong file is uploaded")

                # Validating the upload status icon
                time.sleep(10)
                if self.isdisplayed("file_upload_status_failure", env, UnivWaitFor=180):
                    self.LogScreenshot.fLogScreenshot(message=f'File uploading is done with Failure Icon',
                                                      pass_=True, log=True, screenshot=True)
                    self.click("view_action", env, UnivWaitFor=10)
                    time.sleep(2)
                    td1 = self.select_elements('error_data_table_col1', env)
                    error_data_col1 = []
                    for k in td1:
                        error_data_col1.append(k.text)
                    # Converting list values from string to int
                    error_data_col1 = [int(x) for x in error_data_col1]

                    td2 = self.select_elements('error_data_table_col2', env)
                    error_data_col2 = []
                    for v in td2:
                        error_data_col2.append(v.text)
                    
                    actual_err_msg = [[error_data_col1[i], error_data_col2[i]] for i in range(0, len(error_data_col1))]

                    # self.LogScreenshot.fLogScreenshot(message=f'Excel sheet contains the following errors: '
                    #                                           f'{error_data}',
                    #                                   pass_=True, log=True, screenshot=True)
                    expected_err_msg = self.sort_nested_list(expected_err_msg, 0)
                    actual_err_msg = self.sort_nested_list(actual_err_msg, 0)

                    expected_err_msg = self.sort_nested_list(expected_err_msg, 1)
                    actual_err_msg = self.sort_nested_list(actual_err_msg, 1)

                    if expected_err_msg == actual_err_msg:
                        self.LogScreenshot.fLogScreenshot(message=f"Upload failed as expected error message matches with actual results.",
                                                        pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Error Messages while uploading invalid data. Expected Error Messages : '{expected_err_msg}'. Actual Error Message. '{actual_err_msg}'",
                                                        pass_=False, log=True, screenshot=True)
                        raise Exception(f"Mismatch found in Error Messages while uploading invalid data")                                                      
                    self.click("back_to_view_action_btn", env, UnivWaitFor=10)
                else:
                    raise Exception("Error while uploading the extraction file")
                self.refreshpage()
                time.sleep(5)
            except Exception:
                raise Exception("Error while uploading")

    def upload_file_with_success(self, locatorname, filepath, env):
        expected_upload_status_text = "File(s) uploaded successfully"
        # Read population details from data sheet
        pop_data = self.get_file_details_to_upload(filepath, locatorname)

        for i in pop_data:
            ele = self.select_element("select_update_dropdown", env)
            time.sleep(2)
            select = Select(ele)
            select.select_by_visible_text(i[0])

            # Fetching total rows count before uploading a new file
            table_rows_before = self.select_elements("upload_table_rows", env)
            self.LogScreenshot.fLogScreenshot(message=f'Table length before uploading a new file: '
                                                      f'{len(table_rows_before)}',
                                              pass_=True, log=True, screenshot=False)
            
            jscmd = ReadConfig.get_remove_att_JScommand(16, 'hidden')
            self.jsclick_hide(jscmd)
            self.input_text("add_file", i[1], env)
            try:
                self.jsclick("upload_button", env)
                time.sleep(2)
                # actual_upload_status_text = self.get_text("file_status_popup_text", env, UnivWaitFor=30)
                actual_upload_status_text = self.get_status_text("file_status_popup_text", env)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'File upload is success for Population : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading '
                                                              f'Extraction File for Population : {i[0]}. Actual status message is {actual_upload_status_text} and Expected status message is {expected_upload_status_text}',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message during Extraction file uploading")

                # Fetching total rows count after uploading a new file
                table_rows_after = self.select_elements("upload_table_rows", env)
                self.LogScreenshot.fLogScreenshot(message=f'Table length after uploading a new file: '
                                                          f'{len(table_rows_after)}',
                                                  pass_=True, log=True, screenshot=False)

                if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                    result = []
                    td1 = self.select_elements('upload_table_row_1', env)
                    for m in td1:
                        result.append(m.text)
                    
                    if i[2] in result:
                        self.LogScreenshot.fLogScreenshot(message=f'Correct file with expected filename is being '
                                                                  f'uploaded: {i[2]}',
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        raise Exception("Wrong file is uploaded")

                # Validating the upload status icon
                time.sleep(10)
                if self.isdisplayed("file_upload_status_pass", env, UnivWaitFor=180):
                    self.LogScreenshot.fLogScreenshot(message=f'File uploading is done with Success Icon',
                                                      pass_=True, log=True, screenshot=True)
                    self.click("view_action", env, UnivWaitFor=10)
                    time.sleep(2)
                    self.LogScreenshot.fLogScreenshot(message=f'Success Status Details : ',
                                                      pass_=True, log=True, screenshot=True)
                    self.click("back_to_view_action_btn", env, UnivWaitFor=10)
                else:
                    raise Exception("Error while uploading the extraction file")
                self.refreshpage()
                time.sleep(5)
            except Exception:
                raise Exception("Error while uploading")
