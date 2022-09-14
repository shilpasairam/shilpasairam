import os
from pathlib import Path
import time
import openpyxl
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ManagePopulationsPage import ManagePopulationsPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class ExcludedStudiesPage(Base):

    """Constructor of the ManageQAData Page class"""
    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    def go_to_excludedstudies(self, locator):
        self.click(locator, UnivWaitFor=10)
        time.sleep(5)

    def presence_of_elements(self, locator):
        self.scroll(locator)
        self.wait.until(ec.presence_of_element_located((getattr(By, self.locatortype(locator)), self.locatorpath(locator))))
        self.LogScreenshot.fLogScreenshot(message=f'Manage Excluded Studies option is present in Admin page.',
                                                pass_=True, log=True, screenshot=True)
    
    # Reading slr study type and Excluded study files data for Excluded Studies Page -> upload feature validation
    def get_study_file_details(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        study_type = df.loc[df['Name'] == locatorname]['Study_Types'].dropna().to_list()
        qa_file = df.loc[df['Name'] == locatorname]['ExcludedStudies_Excel_Files'].dropna().to_list()
        filenames = df.loc[df['Name'] == locatorname]['ExcludedStudies_Excel_File_names'].dropna().to_list()
        result = [[study_type[i], os.getcwd() + qa_file[i], filenames[i]] for i in range(0, len(study_type))]
        return result
    
    # Reading slr study type and Excluded study files data for Excluded Studies Page -> override feature validation
    def get_study_file_details_override(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        study_type = df.loc[df['Name'] == locatorname]['Study_Types'].dropna().to_list()
        qa_file = df.loc[df['Name'] == locatorname]['Override_ExcludedStudies_Excel_Files'].dropna().to_list()
        filenames = df.loc[df['Name'] == locatorname]['Override_ExcludedStudies_Excel_File_names'].dropna().to_list()
        result = [(study_type[i], os.getcwd() + qa_file[i], filenames[i]) for i in range(0, len(study_type))]
        return result
    
    # Reading Population data for Excluded Studies Page
    def get_pop_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop = df.loc[df['Name'] == locatorname]['population_name'].dropna().to_list()
        pop_button = df.loc[df['Name'] == locatorname]['pop_radio_button'].dropna().to_list()
        result = [(pop[i], pop_button[i]) for i in range(0, len(pop))]
        return result

    # Reading slr study type and Excluded study files data for complete workflow validation
    def get_file_details(self, filepath):
        file = pd.read_excel(filepath)
        study_type = list(file['Study_Types'].dropna())
        study_file = list(os.getcwd()+file['ExcludedStudies_Excel_Files'].dropna())
        study_filename = list(file['ExcludedStudies_Excel_File_names'].dropna())
        result = [[study_type[i], study_file[i], study_filename[i]] for i in range(0, len(study_type))]
        return result
    
    def add_multiple_excluded_study_data(self, locatorname, filepath):
        expected_upload_status_text = 'File(s) uploaded successfully'
        # Read the username
        username = self.get_text("get_user_name", UnivWaitFor=10)
        firstname = username.split()[0]

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for i in new_pop_data:
                for j in stdy_data:
                    expected_table_values = []
                    self.refreshpage()
                    time.sleep(4)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown")
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    expected_table_values.append(select1.first_selected_option.text)
                    time.sleep(1)
                    
                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                    select2 = Select(stdy_ele)
                    select2.select_by_visible_text(j[0])
                    expected_table_values.append(select2.first_selected_option.text)
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown")
                    select3 = Select(update_ele)
                    select3.select_by_index(1)
                    expected_table_values.append(select3.first_selected_option.text)
                    time.sleep(1)

                    self.input_text("ex_stdy_file_upload", j[1])
                    expected_table_values.append(j[2])
                    time.sleep(2)

                    self.click("ex_stdy_upload_button")
                    time.sleep(3)
                    actual_upload_status_text = self.get_text("ex_stdy_status_text", UnivWaitFor=10)
                    # time.sleep(1)

                    if actual_upload_status_text == expected_upload_status_text:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded Studies File upload is success for Population : {i[0]} -> SLR Type : {j[0]}.',
                                                pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading Excluded Studies File for Population : {i[0]} -> SLR Type: {j[0]}.',
                                                pass_=False, log=True, screenshot=True)
                        raise Exception("Unable to find status message during Excluded Studies file uploading")
                    
                    # Add the firstname to expected values list
                    expected_table_values.append(firstname)

                    # Read table data for specific population and slr study type
                    pop_ele = self.select_element("ex_stdy_pop_dropdown")
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    time.sleep(1)
                    
                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                    select2 = Select(stdy_ele)
                    select2.select_by_visible_text(j[0])
                    time.sleep(1)
                    
                    actual_table_values = []
                    td1 = self.select_elements('ex_stdy_table_data_row_1')
                    for m in td1:
                        actual_table_values.append(m.text)
                    
                    for n in expected_table_values:
                        if n in actual_table_values:
                            self.LogScreenshot.fLogScreenshot(message=f'Correct data is present in table.',
                                                pass_=True, log=True, screenshot=True)
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f'Mismatch is found in table data. Expected values are : {expected_table_values} and Actual values are : {actual_table_values}',
                                                pass_=False, log=True, screenshot=True) 
                            raise Exception(f'Mismatch is found in table data. Expected values are : {expected_table_values} and Actual values are : {actual_table_values}')                  
        except:
            raise Exception("Unable to upload the Excluded Studies data")

    def update_multiple_excluded_study_data(self, locatorname, filepath):
        expected_upload_status_text = 'File(s) uploaded successfully'
        # Read the username
        username = self.get_text("get_user_name", UnivWaitFor=10)
        firstname = username.split()[0]

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details_override(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for i in new_pop_data:
                for j in stdy_data:
                    expected_table_values = []
                    self.refreshpage()
                    time.sleep(3)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown")
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    expected_table_values.append(select1.first_selected_option.text)
                    time.sleep(1)
                    
                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                    select2 = Select(stdy_ele)
                    select2.select_by_visible_text(j[0])
                    expected_table_values.append(select2.first_selected_option.text)
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown")
                    select3 = Select(update_ele)
                    select3.select_by_index(1)
                    expected_table_values.append(select3.first_selected_option.text)
                    time.sleep(1)

                    self.input_text("ex_stdy_file_upload", j[1])
                    expected_table_values.append(j[2])
                    time.sleep(2)

                    self.click("ex_stdy_upload_button")
                    time.sleep(3)
                    self.jsclick("ex_stdy_popup_ok", message="Expected : popup reminder. Actual : popup is not shown")
                    time.sleep(3)
                    actual_upload_status_text = self.get_text("ex_stdy_status_text", UnivWaitFor=10)
                    # time.sleep(1)

                    if actual_upload_status_text == expected_upload_status_text:
                        self.LogScreenshot.fLogScreenshot(message=f'For Population : {i[0]} -> SLR Type : {j[0]}, updating the existing Excluded Studies File is success.',
                                                pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f'For Population : {i[0]} -> SLR Type : {j[0]}, Unable to find status message while updating the existing Excluded Studies File.',
                                                pass_=False, log=True, screenshot=True)
                        raise Exception("Unable to find status message while Updating the existing Excluded Studies file")
                    
                    # Add the firstname to expected values list
                    expected_table_values.append(firstname)

                    # Read table data for specific population and slr study type
                    pop_ele = self.select_element("ex_stdy_pop_dropdown")
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    time.sleep(1)
                    
                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                    select2 = Select(stdy_ele)
                    select2.select_by_visible_text(j[0])
                    time.sleep(1)
                    
                    actual_table_values = []
                    td1 = self.select_elements('ex_stdy_table_data_row_1')
                    for m in td1:
                        actual_table_values.append(m.text)
                    
                    self.LogScreenshot.fLogScreenshot(message=f'Expected values are : {expected_table_values} and Actual values are : {actual_table_values}',
                                                pass_=True, log=True, screenshot=False)
                    
                    for n in expected_table_values:
                        if n in actual_table_values:
                            self.LogScreenshot.fLogScreenshot(message=f'Updated data is present in table.',
                                                pass_=True, log=True, screenshot=True)
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f'Mismatch is found in table data. Expected values are : {expected_table_values} and Actual values are : {actual_table_values}',
                                                pass_=False, log=True, screenshot=True)
                            raise Exception(f'Mismatch is found in table data. Expected values are : {expected_table_values} and Actual values are : {actual_table_values}')                        
        except:
            raise Exception("Unable to update the existing Excluded Studies data")

    def del_multiple_excluded_study_data(self, locatorname, filepath):
        expected_delete_status_text = 'Excluded studies deleted successfully'

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details_override(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for i in new_pop_data:
                for j in stdy_data:
                    self.refreshpage()
                    time.sleep(3)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown")
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    time.sleep(1)
                    
                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                    select2 = Select(stdy_ele)
                    select2.select_by_visible_text(j[0])
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown")
                    select3 = Select(update_ele)
                    select3.select_by_index(1)
                    time.sleep(1)

                    self.LogScreenshot.fLogScreenshot(message=f'Data selected for deletion is : ',
                                                pass_=True, log=True, screenshot=True)

                    self.click("ex_stdy_delete")
                    time.sleep(2)
                    self.click("ex_stdy_popup_ok")
                    time.sleep(2)

                    actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                    # time.sleep(2)

                    if actual_delete_status_text == expected_delete_status_text:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded Studies File Deletion is success.',
                                                pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f'Error while deleting Excluded Studies File. Error Message is {actual_delete_status_text}',
                                                pass_=False, log=True, screenshot=True)
                        raise Exception("Error in Excluded Studies File Deletion")
        except:
            raise Exception("Unable to delete the existing Excluded Studies File")

    def compare_excludedstudy_file_with_report(self, filepath, pop_val, slrfilepath):
        expected_upload_status_text = 'File(s) uploaded successfully'
        # Read the username
        username = self.get_text("get_user_name", UnivWaitFor=10)
        firstname = username.split()[0]

        # Read study types and file paths to upload
        stdy_data = self.get_file_details(filepath)

        try:
            for i in stdy_data:
                expected_table_values = []
                self.refreshpage()
                time.sleep(4)
                self.go_to_excludedstudies("excluded_studies_link")
                pop_ele = self.select_element("ex_stdy_pop_dropdown")
                select1 = Select(pop_ele)
                select1.select_by_visible_text(pop_val)
                expected_table_values.append(select1.first_selected_option.text)
                time.sleep(1)
                
                stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                select2 = Select(stdy_ele)
                select2.select_by_visible_text(i[0])
                expected_table_values.append(select2.first_selected_option.text)
                time.sleep(1)

                update_ele = self.select_element("ex_stdy_update_dropdown")
                select3 = Select(update_ele)
                select3.select_by_index(1)
                expected_table_values.append(select3.first_selected_option.text)
                time.sleep(1)

                self.input_text("ex_stdy_file_upload", i[1])
                expected_table_values.append(i[2])
                time.sleep(2)

                self.click("ex_stdy_upload_button")
                time.sleep(3)
                actual_upload_status_text = self.get_text("ex_stdy_status_text", UnivWaitFor=10)
                # time.sleep(1)

                if actual_upload_status_text == expected_upload_status_text:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded Studies File upload is success for Population : {pop_val} -> SLR Type : {i[0]}.',
                                                pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading Excluded Studies File for Population : {pop_val} -> SLR Type: {i[0]}.',
                                            pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message during Excluded Studies file uploading")
                
                # Add the firstname to expected values list
                expected_table_values.append(firstname)

                # Read table data for specific population and slr study type
                pop_ele = self.select_element("ex_stdy_pop_dropdown")
                select1 = Select(pop_ele)
                select1.select_by_visible_text(pop_val)
                time.sleep(1)
                
                stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                select2 = Select(stdy_ele)
                select2.select_by_visible_text(i[0])
                time.sleep(1)
                
                actual_table_values = []
                td1 = self.select_elements('ex_stdy_table_data_row_1')
                for m in td1:
                    actual_table_values.append(m.text)
                
                for n in expected_table_values:
                    if n in actual_table_values:
                        self.LogScreenshot.fLogScreenshot(message=f'Correct data is present in table.',
                                            pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f'Mismatch is found in table data. Expected values are : {expected_table_values} and Actual values are : {actual_table_values}',
                                            pass_=False, log=True, screenshot=True)     
                        raise Exception(f'Mismatch is found in table data. Expected values are : {expected_table_values} and Actual values are : {actual_table_values}')

                # Go to live slr page
                self.liveslrpage.go_to_liveslr("SLR_Homepage")
                time.sleep(2)
                self.slrreport.select_data(f"{pop_val}", f"{pop_val}_radio_button")
                self.slrreport.select_data(i[0], f"{i[0]}_radio_button")
                self.slrreport.generate_download_report("excel_report")
                time.sleep(5)
                excel_filename = self.slrreport.getFilenameAndValidate(180)
                self.slrreport.validate_filename(excel_filename, slrfilepath)

                update_date_val = expected_table_values[2].translate({ord('/'): None})[-4:]+expected_table_values[2].translate({ord('/'): None})[:4]

                excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
                if f'Excluded studies {update_date_val}' in excel_data.sheetnames:
                    self.LogScreenshot.fLogScreenshot(message=f"'Excluded studies {update_date_val}' sheet is present in complete excel report",
                                                pass_=True, log=True, screenshot=False)

                    excel_sheet = excel_data[f'Excluded studies {update_date_val}']
                    if excel_sheet['H1'].value == 'Back To Toc':
                        self.LogScreenshot.fLogScreenshot(message=f"'Back To Toc' option is present in 'Excluded studies {update_date_val}' sheet",
                                                pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"'Back To Toc' option is not present in 'Excluded studies {update_date_val}' sheet",
                                                pass_=False, log=True, screenshot=False)
                        raise Exception(f"'Back To Toc' option is not present in 'Excluded studies {update_date_val}' sheet")
                    
                    toc_sheet = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name="TOC", skiprows=3)
                    col_data = list(toc_sheet.iloc[:, 1])
                    if f'Excluded studies {update_date_val}' in col_data:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded studies is present in TOC sheet.',
                                                pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded studies is not present in TOC sheet. Available Data from TOC sheet: {col_data}',
                                                pass_=False, log=True, screenshot=False)
                        raise Exception("'Excluded studies' is not present in TOC sheet.")
                    
                    studyfile = pd.read_excel(i[1])
                    excelfile = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=f'Excluded studies {update_date_val}', skiprows=1)

                    if studyfile.equals(excelfile):
                        self.LogScreenshot.fLogScreenshot(message=f"File contents between QA File '{Path(f'{i[1]}').stem}' and Complete Excel Report '{Path(f'ActualOutputs//{excel_filename}').stem}' are matching",
                                                pass_=True, log=True, screenshot=False)
                    else:
                        raise Exception(f"File contents between Study File '{Path(f'{i[1]}').stem}' and Complete Excel Report '{Path(f'ActualOutputs//{excel_filename}').stem}' are not matching")                        
                else:             
                    raise Exception("'Excluded studies' sheet is not present in complete excel report")
        except:
            raise Exception("Error in report comparision between Excluded study file and Complete Excel report")

    def del_after_studyfile_comparison(self, filepath, pop_val, slrfilepath):
        expected_delete_status_text = 'Excluded studies deleted successfully'

        # Read study types and file paths to upload
        stdy_data = self.get_file_details(filepath)

        try:
            for i in stdy_data:
                expected_table_values = []
                self.refreshpage()
                time.sleep(3)
                self.go_to_excludedstudies("excluded_studies_link")
                pop_ele = self.select_element("ex_stdy_pop_dropdown")
                select1 = Select(pop_ele)
                select1.select_by_visible_text(pop_val)
                expected_table_values.append(select1.first_selected_option.text)
                time.sleep(1)
                
                stdy_ele = self.select_element("ex_stdy_stdytype_dropdown")
                select2 = Select(stdy_ele)
                select2.select_by_visible_text(i[0])
                expected_table_values.append(select2.first_selected_option.text)
                time.sleep(1)

                update_ele = self.select_element("ex_stdy_update_dropdown")
                select3 = Select(update_ele)
                select3.select_by_index(1)
                expected_table_values.append(select3.first_selected_option.text)
                time.sleep(1)

                self.LogScreenshot.fLogScreenshot(message=f'Data selected for deletion is : ',
                                            pass_=True, log=True, screenshot=True)

                self.click("ex_stdy_delete")
                time.sleep(2)
                self.click("ex_stdy_popup_ok")
                time.sleep(2)

                actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'Excluded Studies File Deletion is success.',
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Error while deleting Excluded Studies File. Error Message is {actual_delete_status_text}',
                                            pass_=False, log=True, screenshot=True)
                    raise Exception("Error in Excluded Studies File Deletion")

                # Go to live slr page
                self.liveslrpage.go_to_liveslr("SLR_Homepage")
                time.sleep(2)
                self.slrreport.select_data(f"{pop_val}", f"{pop_val}_radio_button")
                self.slrreport.select_data(i[0], f"{i[0]}_radio_button")
                self.slrreport.generate_download_report("excel_report")
                time.sleep(5)
                excel_filename = self.slrreport.getFilenameAndValidate(180)
                self.slrreport.validate_filename(excel_filename, slrfilepath)

                update_date_val = expected_table_values[2].translate({ord('/'): None})[-4:]+expected_table_values[2].translate({ord('/'): None})[:4]
                self.LogScreenshot.fLogScreenshot(message=f"'Excluded studies' date value is 'Excluded studies {update_date_val}'",
                                                pass_=True, log=True, screenshot=False)

                excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
                if f'Excluded studies {update_date_val}' not in excel_data.sheetnames:
                    self.LogScreenshot.fLogScreenshot(message=f"'Excluded studies' sheet is not present in complete excel report as expected",
                                                pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"'Excluded studies' sheet is present in complete excel report which is not expected",
                                                pass_=False, log=True, screenshot=False)
                    raise Exception("'Excluded studies' sheet is present in complete excel report which is not expected")
                
                toc_sheet = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name="TOC", skiprows=3)
                col_data = list(toc_sheet.iloc[:, 1])
                if f'Excluded studies {update_date_val}' not in col_data:
                    self.LogScreenshot.fLogScreenshot(message=f'Excluded studies is not present in TOC sheet as expected.',
                                            pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Excluded studies is present in TOC sheet which is not expected. Available Data from TOC sheet: {col_data}',
                                            pass_=False, log=True, screenshot=False)
                    raise Exception("'Excluded studies' is not present in TOC sheet.")
        except:
            raise Exception("Unable to delete the existing Excluded Studies File")