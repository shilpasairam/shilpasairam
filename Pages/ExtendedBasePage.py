import glob
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

    def select_data(self, locator, locator_button, env):
        time.sleep(3)
        if self.isselected(locator_button, env):
            self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.jsclick(locator, env, UnivWaitFor=10)
            if self.isselected(locator_button, env):
                self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                                  pass_=True, log=True, screenshot=True)

    def select_sub_section(self, locator, locator_button, env, scroll=None):
        if self.scroll(scroll, env, UnivWaitFor=20):
            if self.isselected(locator_button, env):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.jsclick(locator, env, UnivWaitFor=10)
                if self.isselected(locator_button, env):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=True)
            self.scrollback("SLR_page_header", env)

    def select_all_sub_section(self, locator, locator_button, env, scroll=None):
        if self.scroll(scroll, env, UnivWaitFor=20):
            if self.isselected(locator_button, env):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.jsclick(locator, env, UnivWaitFor=10)
                if self.isselected(locator_button, env):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=True)
            self.scrollback("SLR_page_header", env)

    # Read individual column data using locatorname and column name
    def get_testdata_filepath(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        data = (os.getcwd()+(df.loc[df['Testdata_name'] == locatorname]['Testdata_path'].dropna())).to_list()[0]
        return data           
    
    def get_template_file_details(self, filepath, locatorname, column_name):
        df = pd.read_excel(filepath)
        upload_sheet_path = os.getcwd()+(df.loc[df['Name'] == locatorname][column_name].to_list()[0])
        return upload_sheet_path    
    
    # Read individual column data using locatorname and column name
    def get_individual_col_data(self, filepath, locatorname, sheet, col1):
        df = pd.read_excel(filepath, sheet_name=sheet)
        data = df.loc[df['Name'] == locatorname][col1].dropna().to_list()
        return data    
    
    # Read 2 columns data at once using locatorname and column names
    def get_double_col_data(self, filepath, locatorname, sheet, col1, col2):
        df = pd.read_excel(filepath, sheet_name=sheet)
        data1 = df.loc[df['Name'] == locatorname][col1].dropna().to_list()
        data2 = df.loc[df['Name'] == locatorname][col2].dropna().to_list()
        result = [[data1[i], data2[i]] for i in range(0, len(data1))]
        return result

    # Reading 3 columns data at once using locatorname and column names
    def get_triple_col_data(self, filepath, locatorname, sheet, col1, col2, col3, nafilter=True):
        df = pd.read_excel(filepath, sheet_name=sheet, na_filter=nafilter)
        data1 = df.loc[df['Name'] == locatorname][col1].dropna().to_list()
        data2 = df.loc[df['Name'] == locatorname][col2].dropna().to_list()
        data3 = df.loc[df['Name'] == locatorname][col3].dropna().to_list()
        result = [[data1[i], data2[i], data3[i]] for i in range(0, len(data1))]
        return result

    # Reading 4 columns data at once using locatorname and column names
    def get_four_cols_data(self, filepath, locatorname, sheet, col1, col2, col3, col4):
        df = pd.read_excel(filepath, sheet_name=sheet)
        data1 = df.loc[df['Name'] == locatorname][col1].dropna().to_list()
        data2 = df.loc[df['Name'] == locatorname][col2].dropna().to_list()
        data3 = df.loc[df['Name'] == locatorname][col3].dropna().to_list()
        data4 = df.loc[df['Name'] == locatorname][col4].dropna().to_list()
        result = [[data1[i], data2[i], data3[i], data4[i]] for i in range(0, len(data1))]
        return result        

    # Read the individual column values
    def get_data_values(self, filepath, colname):
        file = pd.read_excel(filepath)
        value = list(file[colname].dropna())
        return value
    
    def list_comparison_between_reports_data(self, source_list, compex_list, webex_list=None, word=None):
        idx = 0
        res_index = []
        res_list = []
        if webex_list is not None and word is not None:
            for i in source_list:
                if i != compex_list[idx] and i != webex_list[idx] and i != word[idx]:
                    res_index.append(idx)
                idx = idx + 1
        elif webex_list is None and word is not None:
            for i in source_list:
                if i != compex_list[idx] and i != word[idx]:
                    res_index.append(idx)
                idx = idx + 1
        elif webex_list is not None and word is None:
            for i in source_list:
                if i != compex_list[idx] and i != webex_list[idx]:
                    res_index.append(idx)
                idx = idx + 1
        else:
            for i in source_list:
                if i != compex_list[idx]:
                    res_index.append(idx)
                idx = idx + 1

        if webex_list is not None and word is not None:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])
                res_list[index].append(webex_list[n])
                res_list[index].append(word[n])
        elif webex_list is None and word is not None:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])
                res_list[index].append(word[n])
        elif webex_list is not None and word is None:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])
                res_list[index].append(webex_list[n])
        else:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])

        return res_list
    
    @fWaitFor
    def get_latest_filename(self, UnivWaitFor=0):
        list_of_files = glob.glob('ActualOutputs//*')
        # Get the latest downloaded file name with full path based on the downloaded time
        latest_file_path = max(list_of_files, key=os.path.getctime)
        # Extracting the filename from the latest file path
        latest_filename = os.path.basename(latest_file_path)
        # self.LogScreenshot.fLogScreenshot(message=f"Latest Filename from Actual Outputs folder is : "
        #                                           f"{latest_filename}", pass_=True, log=True, screenshot=False)
        return latest_filename    
    
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

    # Read Required data for LIVESLR Page actions
    def get_slrtest_data(self, filepath, sheet, locatorname, option, option_radio_btn):
        df = pd.read_excel(filepath, sheet_name=sheet)
        opt = df.loc[df['Name'] == locatorname][option].dropna().to_list()
        opt_button = df.loc[df['Name'] == locatorname][option_radio_btn].dropna().to_list()
        result = [[opt[i], opt_button[i]] for i in range(0, len(opt))]
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
    
    def get_additional_criteria_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        criteria = df.loc[df['Name'] == locatorname]['AddtionalParam'].dropna().to_list()
        criteria_btn = df.loc[df['Name'] == locatorname]['AddtionalParam_button'].dropna().to_list()
        section_name = df.loc[df['Name'] == locatorname]['sectionname'].dropna().to_list()
        result = [[criteria[i], criteria_btn[i], section_name[i]] for i in range(0, len(criteria))]
        return result
    
    def upload_file(self, pop_data, env):
        expected_upload_status_text = "File(s) uploaded successfully"
        # Read population details from data sheet
        # pop_data = self.get_extraction_file_to_upload(filepath, 'prodfix', locatorname)

        for i in pop_data:
            # ele = self.select_element("select_update_dropdown", env)
            # time.sleep(2)
            # select = Select(ele)
            # select.select_by_visible_text(i[0])
            selected_update_val = self.base.selectbyvisibletext("select_update_dropdown", i[0], env)

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
                    # result = []
                    # td1 = self.select_elements('upload_table_row_1', env)
                    # for m in td1:
                    #     result.append(m.text)
                    result = self.get_texts('upload_table_row_1', env)
                    
                    if i[2] in result:
                        self.LogScreenshot.fLogScreenshot(message=f'Correct file with expected filename is being '
                                                                  f'uploaded: {i[2]}',
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        raise Exception("Wrong file is uploaded")
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Record count is not incremented after uploading "
                                                              f"the extraction file.",
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
            
    def delete_file(self, pop_data, msg_popup, tablerows, env):
        expected_delete_status_text = "Import status deleted successfully"

        # Fetching total rows count before deleting a file from top of the table
        table_rows_before = self.select_elements(tablerows, env)
        self.LogScreenshot.fLogScreenshot(message=f'Table length before deleting a file: {len(table_rows_before)}',
                                          pass_=True, log=True, screenshot=False)
        
        self.refreshpage()
        time.sleep(5)

        for i in pop_data:
            # result = []
            # td1 = self.select_elements('upload_table_row_1', env)
            # for m in td1:
            #     result.append(m.text)
            result = self.get_texts('upload_table_row_1', env)
            
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
                        self.LogScreenshot.fLogScreenshot(message=f"Record count is decremented after deleting "
                                                                  f"the extraction file.",
                                                          pass_=True, log=True, screenshot=False)
                except Exception:
                    self.LogScreenshot.fLogScreenshot(message=f"Record count is not decremented after deleting "
                                                              f"the extraction file.",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Record count is not decremented after deleting the extraction file.")
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find the uploaded Filename '{i[2]}' in first "
                                                          f"row of the table. Hence aborting the delete operation.",
                                                  pass_=False, log=True, screenshot=True)                
                raise Exception(f"Unable to find the uploaded Filename '{i[2]}' in first row of the table.")
