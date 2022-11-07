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


class ManageQADataPage(Base):

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

    def go_to_manageqadata(self, locator):
        self.click(locator, UnivWaitFor=10)
        time.sleep(5)
    
    def presence_of_elements(self, locator):
        self.wait.until(ec.presence_of_element_located((getattr(By, self.locatortype(locator)),
                                                        self.locatorpath(locator))))
    
    # Reading Population data for ManageQAData Page
    def get_manageqa_pop_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop = df.loc[df['Name'] == locatorname]['QA_Population'].dropna().to_list()
        return pop

    def get_qa_file_details(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        study_type = df.loc[df['Name'] == locatorname]['Study_Types'].dropna().to_list()
        qa_file = df.loc[df['Name'] == locatorname]['QA_Excel_Files'].dropna().to_list()
        result = [[study_type[i], os.getcwd()+qa_file[i]] for i in range(0, len(study_type))]
        return result
    
    def get_qa_file_details_override(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        study_type = df.loc[df['Name'] == locatorname]['Study_Types'].dropna().to_list()
        qa_file = df.loc[df['Name'] == locatorname]['Override_QA_Excel_Files'].dropna().to_list()
        result = [[study_type[i], os.getcwd()+qa_file[i]] for i in range(0, len(study_type))]
        return result
    
    def get_invalid_qa_file_details(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        study_type = df.loc[df['Name'] == locatorname]['Study_Types'].dropna().to_list()
        qa_file = df.loc[df['Name'] == locatorname]['Invalid_Files'].dropna().to_list()
        result = [[study_type[i], os.getcwd()+qa_file[i]] for i in range(0, len(study_type))]
        return result

    def access_manageqadata_page_elements(self, locatorname, filepath):

        # Read population details from data sheet
        pop_name = self.get_manageqa_pop_data(filepath, locatorname)

        # Get StudyType and Files path to upload Managae QA Data
        study_data = self.get_qa_file_details(filepath, locatorname)

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                self.click("select_pop_dropdown")
                self.LogScreenshot.fLogScreenshot(message=f"Population dropdown is accessible. Listed elements are:",
                                                  pass_=True, log=True, screenshot=True)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)
                pop_value = select.first_selected_option.text

                self.click("select_stdy_type_dropdown")
                self.LogScreenshot.fLogScreenshot(message=f"SLR Type dropdown is accessible. Listed elements are:",
                                                  pass_=True, log=True, screenshot=True)
                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"QAName_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_citation", f"QACitation_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_reference", f"QAReference_{pop_value}_{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(2)                
                self.LogScreenshot.fLogScreenshot(message=f'User is able to enter the details for Population : '
                                                          f'{pop_value} -> SLR Type : {i[0]}',
                                                  pass_=True, log=True, screenshot=True)

                self.presence_of_elements("upload_save_button")
                self.presence_of_elements("delete_file_button")
                time.sleep(2)
                self.LogScreenshot.fLogScreenshot(message=f"ManageQAData Page elements are accessible for "
                                                          f"Population : {pop_value} -> SLR Type : {i[0]}.",
                                                  pass_=True, log=True, screenshot=True)
        except Exception:
            raise Exception("Unable to access the Manage QA Data page elements")

    def add_manage_qa_data_with_invalidfile(self, locatorname, filepath):
        expected_error_text = "Incorrect file extension for the quality assessment excel file, " \
                              "it should have the .xls or .xlsx extension"

        # Read population details from data sheet
        pop_name = self.get_manageqa_pop_data(filepath, locatorname)

        # Get StudyType and Files path to upload Managae QA Data
        study_data = self.get_invalid_qa_file_details(filepath, locatorname)

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)
                pop_value = select.first_selected_option.text

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"QAName_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_citation", f"QACitation_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_reference", f"QAReference_{pop_value}_{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(2)                
                self.LogScreenshot.fLogScreenshot(message=f'User is able to enter the details for Population : '
                                                          f'{pop_value} -> SLR Type : {i[0]}',
                                                  pass_=True, log=True, screenshot=True)

                self.click("upload_save_button")
                time.sleep(3)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_upload_status_text == expected_error_text:
                    self.LogScreenshot.fLogScreenshot(message=f"File with invalid format is not uploaded as expected. "
                                                              f"Invalid file is '{Path(f'{i[1]}').stem}'",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading File "
                                                              f"with invalid format for Population : {pop_value} -> "
                                                              f"SLR Type : {i[0]}. Invalid file is "
                                                              f"'{Path(f'{i[1]}').stem}'",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while uploading File with invalid format "
                                    f"for Population : {pop_value} -> SLR Type : {i[0]}. "
                                    f"Invalid file is '{Path(f'{i[1]}').stem}'")
        except Exception:
            raise Exception("Unable to upload the Manage QA Data")

    def add_multiple_manage_qa_data(self, locatorname, filepath):
        expected_upload_status_text = 'QA File successfully uploaded'

        # Read population details from data sheet
        pop_name = self.get_manageqa_pop_data(filepath, locatorname)

        # Get StudyType and Files path to upload Managae QA Data
        study_data = self.get_qa_file_details(filepath, locatorname)

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)
                pop_value = select.first_selected_option.text

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"QAName_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_citation", f"QACitation_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_reference", f"QAReference_{pop_value}_{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(2)                
                self.LogScreenshot.fLogScreenshot(message=f'User is able to enter the details for Population : '
                                                          f'{pop_value} -> SLR Type : {i[0]}',
                                                  pass_=True, log=True, screenshot=True)

                self.click("upload_save_button")
                time.sleep(2)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File upload is success for Population : '
                                                              f'{pop_value} -> SLR Type : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading QA File '
                                                              f'for Population : {pop_value} -> SLR Type : {i[0]}.',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while uploading QA File for Population : "
                                    f"{pop_value} -> SLR Type : {i[0]}.")
        except Exception:
            raise Exception("Unable to upload the Manage QA Data")

    def overwrite_multiple_manage_qa_data(self, locatorname, filepath):
        expected_upload_status_text = 'QA File successfully uploaded'

        # Read population details from data sheet
        pop_name = self.get_manageqa_pop_data(filepath, locatorname)

        # Get StudyType and Files path to Override the existing Managae QA Data
        study_data = self.get_qa_file_details_override(filepath, locatorname)
        
        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)
                pop_value = select.first_selected_option.text

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"QAName_{pop_value}_{i[0]}_Override")
                self.input_text("qa_checklist_citation", f"QACitation_{pop_value}_{i[0]}_Override")
                self.input_text("qa_checklist_reference", f"QAReference_{pop_value}_{i[0]}_Override")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(2)
                self.LogScreenshot.fLogScreenshot(message=f'User is able to override the details for Population : '
                                                          f'{pop_value} -> SLR Type : {i[0]}',
                                                  pass_=True, log=True, screenshot=True)

                self.click("upload_save_button")
                time.sleep(2)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'Updating the existing QA File is success for '
                                                              f'Population : {pop_value} -> SLR Type : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while Updating the '
                                                              f'existing QA Filefor Population : {pop_value} -> '
                                                              f'SLR Type : {i[0]}',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while Updating the existing QA Filefor "
                                    f"Population : {pop_value} -> SLR Type : {i[0]}")
        except Exception:
            raise Exception("Unable to overwrite the Manage QA Data")

    def del_multiple_manage_qa_data(self, locatorname, filepath):
        expected_delete_status_text = 'QA excel file successfully deleted'

        # Read population details from data sheet
        pop_name = self.get_manageqa_pop_data(filepath, locatorname)

        # Get StudyType and Files path to upload Managae QA Data
        study_data = self.get_qa_file_details(filepath, locatorname)

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)
                pop_value = select.first_selected_option.text

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)
                self.LogScreenshot.fLogScreenshot(message=f'Selected Data For Deletion :',
                                                  pass_=True, log=True, screenshot=True)

                self.click("delete_file_button")
                time.sleep(2)
                self.click("delete_file_popup")
                time.sleep(2)

                actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File Deletion is success for Population : '
                                                              f'{pop_value} -> SLR Type : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while deleting QA File '
                                                              f'for Population : {pop_value} -> SLR Type : {i[0]}.',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while deleting QA File for Population : "
                                    f"{pop_value} -> SLR Type : {i[0]}.")
        except Exception:
            raise Exception("Unable to delete the existing QA file")

    def compare_qa_file_with_report(self, locatorname, filepath):
        expected_upload_status_text = 'QA File successfully uploaded'

        # Read population details from data sheet
        pop_name = self.get_manageqa_pop_data(filepath, locatorname)

        # Get StudyType and Files path to upload Managae QA Data
        study_data = self.get_qa_file_details(filepath, locatorname)

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                self.go_to_manageqadata("manage_qa_data_button")
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"QAName_{pop_name[0]}_{i[0]}")
                self.input_text("qa_checklist_citation", f"QACitation_{pop_name[0]}_{i[0]}")
                self.input_text("qa_checklist_reference", f"QAReference_{pop_name[0]}_{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(3)
                self.LogScreenshot.fLogScreenshot(message=f'User is able to enter the details for Population : '
                                                          f'{pop_name[0]} -> SLR Type : {i[0]}',
                                                  pass_=True, log=True, screenshot=True)

                self.click("upload_save_button")
                time.sleep(3)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File upload is success for Population : '
                                                              f'{pop_name[0]} -> SLR Type : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading QA File '
                                                              f'for Population : {pop_name[0]} -> SLR Type : {i[0]}.',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while uploading QA File for Population : "
                                    f"{pop_name[0]} -> SLR Type : {i[0]}.")

                # Go to live slr page
                self.liveslrpage.go_to_liveslr("SLR_Homepage")
                if i[0] == 'Quality of life':
                    i[0] = 'Quality of Life'
                self.slrreport.select_data(f"{pop_name[0]}", f"{pop_name[0]}_radio_button")
                self.slrreport.select_data(i[0], f"{i[0]}_radio_button")
                self.slrreport.generate_download_report("excel_report")
                time.sleep(5)
                excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                self.slrreport.validate_filename(excel_filename1, filepath)

                excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename1}')
                if 'Quality Assessment' in excel_data.sheetnames:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' sheet is present in complete "
                                                              f"excel report",
                                                      pass_=True, log=True, screenshot=False)

                    excel_sheet = excel_data['Quality Assessment']
                    if excel_sheet['A1'].value == 'Back To Toc':
                        self.LogScreenshot.fLogScreenshot(message=f"'Back To Toc' option is present",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"'Back To Toc' option is not present",
                                                          pass_=False, log=True, screenshot=False)
                        raise Exception(f"'Back To Toc' option is not present")
                    
                    toc_sheet = pd.read_excel(f'ActualOutputs//{excel_filename1}', sheet_name="TOC", skiprows=3)
                    col_data = list(toc_sheet.iloc[:, 1])
                    if 'Quality Assessment' in col_data:
                        self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' is present in TOC sheet.",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' is not present in TOC sheet. "
                                                                  f"Available Data from TOC sheet: {col_data}",
                                                          pass_=False, log=True, screenshot=False)
                        raise Exception(f"'Quality Assessment' is present in TOC sheet which is not expected. "
                                        f"Available Data from TOC sheet: {col_data}")
                    
                    qafile = pd.read_excel(i[1])
                    excelfile = pd.read_excel(f'ActualOutputs//{excel_filename1}', sheet_name="Quality Assessment")

                    # Removing the 'Back To Toc' column to compare the exact data with uploaded file
                    excelfile = excelfile.iloc[:, 1:]
                    
                    if qafile.equals(excelfile):
                        self.LogScreenshot.fLogScreenshot(message=f"File contents between QA File "
                                                                  f"'{Path(f'{i[1]}').stem}' and Complete Excel Report "
                                                                  f"'{Path(f'ActualOutputs//{excel_filename1}').stem}' "
                                                                  f"are matching",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"File contents between QA File "
                                                                  f"'{Path({i[1]}).stem}' and Complete Excel Report "
                                                                  f"'{Path(f'ActualOutputs//{excel_filename1}').stem}' "
                                                                  f"are not matching",
                                                          pass_=False, log=True, screenshot=False)
                        raise Exception(f"File contents between QA File '{Path({i[1]}).stem}' "
                                        f"and Complete Excel Report '{Path(f'ActualOutputs//{excel_filename1}').stem}' "
                                        f"are not matching")
                else:
                    raise Exception("'Quality Assessment' sheet is not present in complete excel report")
        except Exception:
            raise Exception("Error in report comparision between QA data file and Complete Excel report")

    def del_data_after_qafile_comparison(self, locatorname, filepath):
        expected_delete_status_text = 'QA excel file successfully deleted'

        # Read population details from data sheet
        pop_name = self.get_manageqa_pop_data(filepath, locatorname)

        # Get StudyType and Files path to upload Managae QA Data
        study_data = self.get_qa_file_details(filepath, locatorname)

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                self.go_to_manageqadata("manage_qa_data_button")
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                if i[0] == 'Quality of Life':
                    i[0] = 'Quality of life'
                select.select_by_visible_text(i[0])
                time.sleep(1)
                self.LogScreenshot.fLogScreenshot(message=f'Selected Data For Deletion :',
                                                  pass_=True, log=True, screenshot=True)

                self.click("delete_file_button")
                time.sleep(2)
                self.click("delete_file_popup")
                time.sleep(2)

                actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File Deletion is success for Population : '
                                                              f'{pop_name[0]} -> SLR Type : {i[0]}.',
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while deleting QA File '
                                                              f'for Population : {pop_name[0]} -> SLR Type : {i[0]}.',
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while deleting QA File for Population : "
                                    f"{pop_name[0]} -> SLR Type : {i[0]}.")
                
                # Go to live slr page
                self.liveslrpage.go_to_liveslr("SLR_Homepage")
                if i[0] == 'Quality of life':
                    i[0] = 'Quality of Life'
                self.slrreport.select_data(f"{pop_name[0]}", f"{pop_name[0]}_radio_button")
                self.slrreport.select_data(i[0], f"{i[0]}_radio_button")
                self.slrreport.generate_download_report("excel_report")
                time.sleep(5)
                excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                self.slrreport.validate_filename(excel_filename1, filepath)

                excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename1}')
                if 'Quality Assessment' not in excel_data.sheetnames:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' sheet is not present in complete "
                                                              f"excel report as expected",
                                                      pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' sheet is present in complete "
                                                              f"excel report which is not expected",
                                                      pass_=True, log=True, screenshot=False)
                    raise Exception(f"'Quality Assessment' sheet is present in complete excel report which is not "
                                    f"expected")
                
                toc_sheet = pd.read_excel(f'ActualOutputs//{excel_filename1}', sheet_name="TOC", skiprows=3)
                col_data = list(toc_sheet.iloc[:, 1])
                if 'Quality Assessment' not in col_data:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' is not present in TOC sheet "
                                                              f"as expected",
                                                      pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' is present in TOC sheet which "
                                                              f"is not expected. Available Data from "
                                                              f"TOC sheet: {col_data}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"'Quality Assessment' is present in TOC sheet which is not expected. "
                                    f"Available Data from TOC sheet: {col_data}")
        except Exception:
            raise Exception("Unable to delete the existing QA file")
