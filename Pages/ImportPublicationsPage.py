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


class ImportPublicationPage(Base):

    """Constructor of the ImportPublication Page class"""
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

    def go_to_importpublications(self, locator, button):
        self.click(locator, UnivWaitFor=10)
        self.jsclick(button)
        time.sleep(3)

    def get_upload_file_details(self, filepath):
        file = pd.read_excel(filepath)
        sheet_name = list(file['Expected_File_names'].dropna())
        sheet_path = list(os.getcwd()+file['Files_to_upload'].dropna())
        admin_page_data = [(sheet_name[i], sheet_path[i]) for i in range(0, len(sheet_name))]
        return admin_page_data
    
    def select_update(self, locator, pop_index):
        ele = self.select_element(locator)
        time.sleep(2)
        select = Select(ele)
        select.select_by_index(pop_index)

    def upload_file(self, locator, expected_filename, filepath, upload, msg_popup, tablerows):
        # Fetching total rows count before uploading a new file
        table_rows_before = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before uploading a new file: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)
        
        jscmd = ReadConfig.getJScommand()
        self.jsclick_hide(jscmd)
        self.input_text(locator, filepath)
        try:
            self.jsclick(upload)
            time.sleep(3)
            upload_text = self.get_text(msg_popup, UnivWaitFor=30)
            self.LogScreenshot.fLogScreenshot(message=f'Message popup: {upload_text}',
                                          pass_=True, log=True, screenshot=False)

            self.assertText("File(s) uploaded successfully", upload_text)

            # Fetching total rows count after uploading a new file
            table_rows_after = self.select_elements(tablerows)
            self.LogScreenshot.fLogScreenshot(message=f'Table length after uploading a new file: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

            if len(table_rows_after) > len(table_rows_before) != len(table_rows_after):
                result = []
                td1 = self.select_elements('upload_table_row_1')
                for m in td1:
                    result.append(m.text)
                
                if expected_filename in result:
                    self.LogScreenshot.fLogScreenshot(message=f'Correct file with expected filename is being uploaded: {expected_filename}',
                                          pass_=True, log=True, screenshot=False)
                else:
                    raise Exception("Wrong file is uploaded")

            if self.isdisplayed("file_upload_status_pass", UnivWaitFor=60):
                self.LogScreenshot.fLogScreenshot(message=f'File uploading is done',
                                          pass_=True, log=True, screenshot=False)
            elif self.isdisplayed("file_upload_status_failure", UnivWaitFor=60):
                self.click("view_action", UnivWaitFor=10)
                time.sleep(2)
                td = self.select_elements('error_data_table')
                error_data = []
                for n in td:
                    error_data.append(n.text)
                self.LogScreenshot.fLogScreenshot(message=f'Excel sheet contains the following errors: {error_data}',
                                          pass_=False, log=True, screenshot=False)
            else:
                raise Exception("Error while uploading the extraction file")
            self.refreshpage()
            time.sleep(5)
        except:
            raise Exception("Error while uploading")

    def delete_file(self, del_locator, del_locator_popup, msg_popup, tablerows):
        # Fetching total rows count before deleting a file from top of the table
        table_rows_before = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a file: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)
        
        self.click(del_locator)
        time.sleep(2)
        self.click(del_locator_popup)
        time.sleep(2)

        del_text = self.get_text(msg_popup, UnivWaitFor=10)
        self.LogScreenshot.fLogScreenshot(message=f'Message popup: {del_text}',
                                          pass_=True, log=True, screenshot=False)

        self.assertText("Import status deleted successfully", del_text)

        # Fetching total rows count before deleting a file from top of the table
        table_rows_after = self.select_elements(tablerows)
        self.LogScreenshot.fLogScreenshot(message=f'Table length after deleting a file: {len(table_rows_after)}',
                                          pass_=True, log=True, screenshot=False)

        try:
            if len(table_rows_before) > len(table_rows_after) != len(table_rows_before):
                self.LogScreenshot.fLogScreenshot(message=f'Record deletion is successful',
                                            pass_=True, log=True, screenshot=False)  
            self.refreshpage()
            time.sleep(5)                  
        except:
            self.LogScreenshot.fLogScreenshot(message=f'Record deletion is not successful',
                                            pass_=False, log=True, screenshot=False)  
            raise Exception("Error in deleting the imported file")
