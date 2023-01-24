import os
from pathlib import Path
import time

import docx
import openpyxl
import pandas as pd
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from pandas.core.common import flatten

from Pages.Base import Base, fWaitFor
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


class ExtendedBase(Base):

    """Constructor of the ExtendedBase class"""
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
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 20)

    # Read Population data for LIVESLR Page
    def get_population_data(self, filepath, sheet, locatorname):
        df = pd.read_excel(filepath, sheet_name=sheet)
        pop = df.loc[df['Name'] == locatorname]['Population'].dropna().to_list()
        pop_button = df.loc[df['Name'] == locatorname]['Population_Radio_button'].dropna().to_list()
        result = [[pop[i], pop_button[i]] for i in range(0, len(pop))]
        return result
    
    # Read SLRTYPE data for LIVESLR Page
    def get_slrtype_data(self, filepath, sheet, locatorname):
        df = pd.read_excel(filepath, sheet_name=sheet)
        slrtype = df.loc[df['Name'] == locatorname]['slrtype'].dropna().to_list()
        slrtype_button = df.loc[df['Name'] == locatorname]['slrtype_Radio_button'].dropna().to_list()
        result = [[slrtype[i], slrtype_button[i]] for i in range(0, len(slrtype))]
        return result
    
    # Read expected test data file for comparison
    def get_source_template(self, filepath, sheet, locatorname):
        file = pd.read_excel(filepath, sheet_name=sheet)
        expectedfilepath = (os.getcwd()+(file.loc[file['Name'] == locatorname]['ExpectedSourceTemplateFile'].
                                         dropna())).to_list()
        return expectedfilepath
    
    # Reading Population data for Excluded Studies Page
    def get_file_details_to_upload(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop_name = df.loc[df['Name'] == locatorname]['Population_name'].dropna().to_list()
        path = df.loc[df['Name'] == locatorname]['Files_to_upload'].dropna().to_list()
        filename = df.loc[df['Name'] == locatorname]['Expected_File_names'].dropna().to_list()
        result = [[pop_name[i], os.getcwd() + path[i], filename[i]] for i in range(0, len(pop_name))]
        return result    
    
    def upload_file(self, pop_name, file_to_upload):
        expected_upload_status_text = "File(s) uploaded successfully"
    
        ele = self.select_element("select_update_dropdown")
        time.sleep(2)
        select = Select(ele)
        select.select_by_visible_text(pop_name)
        
        jscmd = ReadConfig.get_remove_att_JScommand(16, 'hidden')
        self.jsclick_hide(jscmd)
        self.input_text("add_file", file_to_upload)
        try:
            self.jsclick("upload_button", UnivWaitFor=30)
            time.sleep(4)
            if self.isdisplayed("file_status_popup_text"):
                actual_upload_status_text = self.get_text("file_status_popup_text", UnivWaitFor=30)
            else:
                time.sleep(2)
                actual_upload_status_text = self.get_text("file_status_popup_text", UnivWaitFor=30)
            
            if actual_upload_status_text == expected_upload_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"File upload is success for Population : {pop_name}. "
                                                          f"Extraction Filename is '{Path(f'{file_to_upload}').stem}'",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading Extraction '
                                                          f'File for Population : {pop_name}.',
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message during Extraction file uploading")

            time.sleep(10)
            if self.isdisplayed("file_upload_status_pass", UnivWaitFor=180):
                self.LogScreenshot.fLogScreenshot(message=f'File uploading is done with Success Icon',
                                                  pass_=True, log=True, screenshot=True)
            else:
                raise Exception("Error while uploading the extraction file")

            self.refreshpage()
            time.sleep(5)
        except Exception:
            raise Exception("Error while uploading")
            
    def delete_file(self, expected_filename):
        expected_delete_status_text = "Import status deleted successfully"        
        self.refreshpage()
        time.sleep(5)

        result = []
        td1 = self.select_elements('upload_table_row_1')
        for m in td1:
            result.append(m.text)

        # Check the uploaded filename before deleting the record
        if expected_filename in result:
            self.LogScreenshot.fLogScreenshot(message=f"Uploaded Filename '{expected_filename}' is present in "
                                                      f"the table. Performing the delete operation.",
                                              pass_=True, log=True, screenshot=True)
        
            self.click("delete_file")
            time.sleep(2)
            self.click("delete_file_popup")
            time.sleep(3)

            actual_delete_status_text = self.get_text("file_status_popup_text", UnivWaitFor=30)
            
            if actual_delete_status_text == expected_delete_status_text:
                self.LogScreenshot.fLogScreenshot(message=f'Extraction File Deletion is success.',
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while deleting Extraction '
                                                          f'File', pass_=False, log=True, screenshot=True)
                raise Exception("Error during Extraction File Deletion")   
        else:
            self.LogScreenshot.fLogScreenshot(message=f'Unable to find uploaded filename in table. Uploaded filename '
                                                      f'is {expected_filename}',
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Unable to find uploaded filename in table")