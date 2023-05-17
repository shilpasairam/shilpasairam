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

    # Reading Population data for Excluded Studies Page
    def get_file_details_to_upload(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop_name = df.loc[df['Name'] == locatorname]['Population_name'].dropna().to_list()
        path = df.loc[df['Name'] == locatorname]['Files_to_upload'].dropna().to_list()
        filename = df.loc[df['Name'] == locatorname]['Expected_File_names'].dropna().to_list()
        result = [[pop_name[i], os.getcwd() + path[i], filename[i]] for i in range(0, len(pop_name))]
        return result

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
                                                              f'Extraction File for Population : {i[0]}. '
                                                              f'Actual status message is {actual_upload_status_text} '
                                                              f'and Expected status message is '
                                                              f'{expected_upload_status_text}',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message during Extraction file uploading")

                # Fetching total rows count after uploading a new file
                table_rows_after = self.select_elements("upload_table_rows", env)
                self.LogScreenshot.fLogScreenshot(message=f'Table length after uploading a new file: '
                                                          f'{len(table_rows_after)}',
                                                  pass_=True, log=True, screenshot=False)

                if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                    self.LogScreenshot.fLogScreenshot(message=f"Record count is incremented after uploading the "
                                                              f"extraction file.",
                                                      pass_=True, log=True, screenshot=False)
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
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Record count is not incremented after uploading the "
                                                              f"extraction file.",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Record count is not incremented after uploading the extraction file.")

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
                                                              f'Extraction File for Population : {i[0]}. Actual '
                                                              f'status message is {actual_upload_status_text} and '
                                                              f'Expected status message is '
                                                              f'{expected_upload_status_text}',
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

                    expected_err_msg = self.sort_nested_list(expected_err_msg, 0)
                    actual_err_msg = self.sort_nested_list(actual_err_msg, 0)

                    expected_err_msg = self.sort_nested_list(expected_err_msg, 1)
                    actual_err_msg = self.sort_nested_list(actual_err_msg, 1)

                    comparison_result = self.exbase.list_comparison_between_reports_data(sorted(col2_data),
                                                                                         sorted(error_data_col2))

                    if len(comparison_result) == 0 and expected_err_msg == actual_err_msg:
                        self.LogScreenshot.fLogScreenshot(message=f"Upload is Successful. Expected error messages "
                                                                  f"matches with Actual error messages.",
                                                          pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Error Messages while uploading "
                                                                  f"invalid data. Mismatch values are arranged in "
                                                                  f"following order -> Expected Error Message, "
                                                                  f"Actual Error Message. {comparison_result}",
                                                          pass_=False, log=True, screenshot=True)
                        raise Exception(f"Mismatch found in Error Messages while uploading invalid data")

                    self.click("back_to_view_action_btn", env, UnivWaitFor=10)
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
                time.sleep(4)

                # actual_delete_status_text = self.get_text(msg_popup, env, UnivWaitFor=30)
                actual_delete_status_text = self.get_status_text(msg_popup, env)
                
                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'Extraction File Deletion is success.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while deleting '
                                                              f'Extraction File. Actual status message is '
                                                              f'{actual_delete_status_text} and Expected status '
                                                              f'message is {expected_delete_status_text}',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Error during Extraction File Deletion")

                # Fetching total rows count after deleting a file from top of the table
                table_rows_after = self.select_elements(tablerows, env)
                self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a file: '
                                                          f'{len(table_rows_after)}',
                                                  pass_=True, log=True, screenshot=False)

                try:
                    if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                        self.LogScreenshot.fLogScreenshot(message=f"Record count is decremented after deleting the "
                                                                  f"extraction file.",
                                                          pass_=True, log=True, screenshot=False)
                except Exception:
                    self.LogScreenshot.fLogScreenshot(message=f"Record count is not decremented after deleting the "
                                                              f"extraction file.",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Record count is not decremented after deleting the extraction file.")
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find the uploaded Filename '{i[2]}' in first "
                                                          f"row of the table. Hence aborting the delete operation.",
                                                  pass_=False, log=True, screenshot=True)                
                raise Exception(f"Unable to find the uploaded Filename '{i[2]}' in first row of the table.")

    def upload_file_for_same_population(self, locatorname, filepath, env):
        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction File for the existing population validation "
                                                  f"is started***", pass_=True, log=True, screenshot=False)
        # expected_err_msg = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'error_msg')
        
        # # Read population details from data sheet
        # pop_data = self.get_file_details_to_upload(filepath, locatorname)

        try:
            self.upload_file_with_errors(locatorname, filepath, env)
            self.delete_file(locatorname, filepath, "file_status_popup_text", "upload_table_rows", env)
            # for i in pop_data:
            #     ele = self.select_element("select_update_dropdown", env)
            #     time.sleep(2)
            #     select = Select(ele)
            #     select.select_by_visible_text(i[0])

            #     time.sleep(1)
            #     actual_status_text = self.get_status_text("file_status_popup_text", env)                                                    

            #     if actual_status_text == expected_err_msg[0]:
            #         self.LogScreenshot.fLogScreenshot(message=f"User is not allowed to upload extraction file for "
            #                                                   f"existing population with same update.",
            #                                           pass_=True, log=True, screenshot=True)
            #     else:
            #         self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading "
            #                                                   f"Extraction File for existing Population : {i[0]}. "
            #                                                   f"Actual status message is {actual_status_text} and "
            #                                                   f"Expected status message is {expected_err_msg[0]}",
            #                                           pass_=False, log=True, screenshot=True)
            #         raise Exception("Unable to find status message while uploading Extraction File for existing "
            #                         "Population with same update")
                
            #     jscmd = ReadConfig.get_remove_att_JScommand(16, 'hidden')
            #     self.jsclick_hide(jscmd)
            #     self.input_text("add_file", i[1], env)
            
            #     if not self.isenabled("upload_button", env):
            #         self.LogScreenshot.fLogScreenshot(message=f"Upload file button is not clickable while trying to "
            #                                                   f"upload extraction file for existing population with "
            #                                                   f"same update as expected.",
            #                                           pass_=True, log=True, screenshot=True)
            #         self.LogScreenshot.fLogScreenshot(message=f"*****Hence Performing the delete operation for the "
            #                                                   f"successfully uploaded record and trying to re-upload "
            #                                                   f"the extraction for the same population.*****",
            #                                           pass_=True, log=True, screenshot=False)
            #         self.delete_file(locatorname, filepath, "file_status_popup_text", "upload_table_rows", env)
            #         self.upload_file_with_success(locatorname, filepath, env)
            #     else:
            #         self.jsclick("upload_button", env)
            #         time.sleep(2)
            #         actual_upload_status_text = self.get_status_text("file_status_popup_text", env)
            #         self.LogScreenshot.fLogScreenshot(message=f"Upload file button is clickable. File is uploaded "
            #                                                   f"with success toaster message "
            #                                                   f"'{actual_upload_status_text}'",
            #                                           pass_=False, log=True, screenshot=True)
            #         raise Exception(f"Upload file button is clickable for the Population which has already update "
            #                         f"which is not expected")
            #     self.refreshpage()
            #     time.sleep(5)
        except Exception:
            raise Exception("Error while uploading")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Upload Extraction File for the existing population validation "
                                                  f"is completed***", pass_=True, log=True, screenshot=False)
