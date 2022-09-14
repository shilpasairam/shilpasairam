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

    def get_qa_file_details(self, filepath):
        file = pd.read_excel(filepath)
        study_type = list(file['Study_Types'].dropna())
        qa_file = list(os.getcwd()+file['QA_Excel_Files'].dropna())
        result = [[study_type[i], qa_file[i]] for i in range(0, len(study_type))]
        return result
    
    def get_qa_file_details_override(self, filepath):
        file = pd.read_excel(filepath)
        study_type = list(file['Study_Types'].dropna())
        qa_file = list(os.getcwd()+file['Override_QA_Excel_Files'].dropna())
        result = [(study_type[i], qa_file[i]) for i in range(0, len(study_type))]
        return result

    # def add_manage_qa_data(self, manage_qa_page, study_data, filepath):
    #     expected_upload_status_text = 'QA File successfully uploaded'
        
    #     self.click(manage_qa_page)
    #     # Read population details from data sheet
    #     new_pop_data, new_pop_val = self.mngpoppage.get_pop_data(filepath)
    #     try:
    #         for i in study_data:
    #             self.refreshpage()
    #             pop_ele = self.select_element("select_pop_dropdown")
    #             select = Select(pop_ele)
    #             select.select_by_visible_text(new_pop_val[0])

    #             stdy_ele = self.select_element("select_stdy_type_dropdown")
    #             select = Select(stdy_ele)
    #             select.select_by_visible_text(i[0])

    #             self.input_text("qa_checklist_name", f"QAName_{i[0]}")
    #             self.input_text("qa_checklist_citation", f"QACitation_{i[0]}")
    #             self.input_text("qa_checklist_reference", f"QAReference_{i[0]}")
    #             self.input_text("qa_excel_file_upload", i[1])
    #             time.sleep(1)

    #             self.click("upload_save_button")
    #             time.sleep(1)
    #             actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
    #             time.sleep(2)

    #             if actual_upload_status_text == expected_upload_status_text:
    #                 self.LogScreenshot.fLogScreenshot(message=f'QA File upload is success.',
    #                                         pass_=True, log=True, screenshot=True)
    #             else:
    #                 self.LogScreenshot.fLogScreenshot(message=f'Error while uploading QA File. Error Message is {actual_upload_status_text}',
    #                                         pass_=False, log=True, screenshot=True)
    #                 raise Exception("Error in QA file uploading")
    #     except:
    #         raise Exception("Unable to upload the Manage QA Data")

    # def del_manage_qa_data(self, manage_qa_page, study_data, del_locator, del_locator_popup, filepath):
    #     expected_delete_status_text = 'QA excel file successfully deleted'

    #     self.click(manage_qa_page)
    #     # Read population details from data sheet
    #     new_pop_data, new_pop_val = self.mngpoppage.get_pop_data(filepath)
    #     try:
    #         for i in study_data:
    #             self.refreshpage()
    #             pop_ele = self.select_element("select_pop_dropdown")
    #             select = Select(pop_ele)
    #             select.select_by_visible_text(new_pop_val[0])

    #             stdy_ele = self.select_element("select_stdy_type_dropdown")
    #             select = Select(stdy_ele)
    #             select.select_by_visible_text(i[0])

    #             self.click(del_locator)
    #             time.sleep(1)
    #             self.click(del_locator_popup)
    #             time.sleep(1)

    #             actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)
    #             time.sleep(2)

    #             if actual_delete_status_text == expected_delete_status_text:
    #                 self.LogScreenshot.fLogScreenshot(message=f'QA File Deletion is success.',
    #                                         pass_=True, log=True, screenshot=True)
    #             else:
    #                 self.LogScreenshot.fLogScreenshot(message=f'Error while deleting QA File. Error Message is {actual_delete_status_text}',
    #                                         pass_=False, log=True, screenshot=True)
    #                 raise Exception("Error in QA file Deletion")
        
    #     except:
    #         raise Exception("Unable to delete the existing QA file")

    def add_multiple_manage_qa_data(self, study_data, pop_index):
        expected_upload_status_text = 'QA File successfully uploaded'

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_index(pop_index)
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
                self.LogScreenshot.fLogScreenshot(message=f'User is able to enter the details for Population : {pop_value} -> SLR Type : {i[0]}',
                                            pass_=True, log=True, screenshot=True)

                self.click("upload_save_button")
                time.sleep(2)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File upload is success for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading QA File for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while uploading QA File for Population : {pop_value} -> SLR Type : {i[0]}.")
        except:
            raise Exception("Unable to upload the Manage QA Data")

    def overwrite_multiple_manage_qa_data(self, study_data, pop_index):
        expected_upload_status_text = 'QA File successfully uploaded'
        
        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_index(pop_index)
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
                self.LogScreenshot.fLogScreenshot(message=f'User is able to override the details for Population : {pop_value} -> SLR Type : {i[0]}',
                                            pass_=True, log=True, screenshot=True)

                self.click("upload_save_button")
                time.sleep(2)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'Updating the existing QA File is success for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while Updating the existing QA Filefor Population : {pop_value} -> SLR Type : {i[0]}',
                                            pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while Updating the existing QA Filefor Population : {pop_value} -> SLR Type : {i[0]}")
        except:
            raise Exception("Unable to overwrite the Manage QA Data")

    def del_multiple_manage_qa_data(self, study_data, del_locator, del_locator_popup, pop_index):
        expected_delete_status_text = 'QA excel file successfully deleted'

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_index(pop_index)
                time.sleep(1)
                pop_value = select.first_selected_option.text

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)
                self.LogScreenshot.fLogScreenshot(message=f'Selected Data For Deletion :',
                                            pass_=True, log=True, screenshot=True)

                self.click(del_locator)
                time.sleep(2)
                self.click(del_locator_popup)
                time.sleep(2)

                actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File Deletion is success for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while deleting QA File for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while deleting QA File for Population : {pop_value} -> SLR Type : {i[0]}.")
        except:
            raise Exception("Unable to delete the existing QA file")

    def compare_qa_file_with_report(self, study_data, pop_value, slrfilepath):
        expected_upload_status_text = 'QA File successfully uploaded'

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                self.go_to_manageqadata("manage_qa_data_button")
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_value)
                time.sleep(1)

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"QAName_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_citation", f"QACitation_{pop_value}_{i[0]}")
                self.input_text("qa_checklist_reference", f"QAReference_{pop_value}_{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(3)
                self.LogScreenshot.fLogScreenshot(message=f'User is able to enter the details for Population : {pop_value} -> SLR Type : {i[0]}',
                                            pass_=True, log=True, screenshot=True)

                self.click("upload_save_button")
                time.sleep(3)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)
                # time.sleep(2)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File upload is success for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading QA File for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while uploading QA File for Population : {pop_value} -> SLR Type : {i[0]}.")

                # Go to live slr page
                self.liveslrpage.go_to_liveslr("SLR_Homepage")
                if i[0] == 'Quality of life':
                    i[0] = 'Quality of Life'
                self.slrreport.select_data(f"{pop_value}", f"{pop_value}_radio_button")
                self.slrreport.select_data(i[0], f"{i[0]}_radio_button")
                self.slrreport.generate_download_report("excel_report")
                time.sleep(5)
                excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                self.slrreport.validate_filename(excel_filename1, slrfilepath)

                excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename1}')
                if 'Quality Assessment' in excel_data.sheetnames:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' sheet is present in complete excel report",
                                                pass_=True, log=True, screenshot=False)

                    excel_sheet = excel_data['Quality Assessment']
                    if excel_sheet['E1'].value == 'Back To Toc':
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
                        self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' is not present in TOC sheet. Available Data from TOC sheet: {col_data}",
                                                pass_=False, log=True, screenshot=False)
                        raise Exception(f"'Quality Assessment' is present in TOC sheet which is not expected. Available Data from TOC sheet: {col_data}")
                    
                    qafile = pd.read_excel(i[1])
                    excelfile = pd.read_excel(f'ActualOutputs//{excel_filename1}', sheet_name="Quality Assessment", skiprows=1)

                    if qafile.equals(excelfile):
                        self.LogScreenshot.fLogScreenshot(message=f"File contents between QA File '{Path(f'{i[1]}').stem}' and Complete Excel Report '{Path(f'ActualOutputs//{excel_filename1}').stem}' are matching",
                                                pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"File contents between QA File '{Path({i[1]}).stem}' and Complete Excel Report '{Path(f'ActualOutputs//{excel_filename1}').stem}' are not matching",
                                                pass_=False, log=True, screenshot=False)
                        raise Exception(f"File contents between QA File '{Path({i[1]}).stem}' and Complete Excel Report '{Path(f'ActualOutputs//{excel_filename1}').stem}' are not matching")
                else:
                    raise Exception("'Quality Assessment' sheet is not present in complete excel report")
        except:
            raise Exception("Error in report comparision between QA data file and Complete Excel report")

    def del_data_after_qafile_comparison(self, study_data, pop_value, slrfilepath):
        expected_delete_status_text = 'QA excel file successfully deleted'

        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                self.go_to_manageqadata("manage_qa_data_button")
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_value)
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
                    self.LogScreenshot.fLogScreenshot(message=f'QA File Deletion is success for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while deleting QA File for Population : {pop_value} -> SLR Type : {i[0]}.',
                                            pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while deleting QA File for Population : {pop_value} -> SLR Type : {i[0]}.")
                
                # Go to live slr page
                self.liveslrpage.go_to_liveslr("SLR_Homepage")
                if i[0] == 'Quality of life':
                    i[0] = 'Quality of Life'
                self.slrreport.select_data(f"{pop_value}", f"{pop_value}_radio_button")
                self.slrreport.select_data(i[0], f"{i[0]}_radio_button")
                self.slrreport.generate_download_report("excel_report")
                time.sleep(5)
                excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                self.slrreport.validate_filename(excel_filename1, slrfilepath)

                excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename1}')
                if 'Quality Assessment' not in excel_data.sheetnames:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' sheet is not present in complete excel report as expected",
                                                pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' sheet is present in complete excel report which is not expected",
                                                pass_=True, log=True, screenshot=False)
                    raise Exception(f"'Quality Assessment' sheet is present in complete excel report which is not expected")
                
                toc_sheet = pd.read_excel(f'ActualOutputs//{excel_filename1}', sheet_name="TOC", skiprows=3)
                col_data = list(toc_sheet.iloc[:, 1])
                if 'Quality Assessment' not in col_data:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' is not present in TOC sheet as expected",
                                            pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"'Quality Assessment' is present in TOC sheet which is not expected. Available Data from TOC sheet: {col_data}",
                                            pass_=False, log=True, screenshot=False)
                    raise Exception(f"'Quality Assessment' is present in TOC sheet which is not expected. Available Data from TOC sheet: {col_data}")
        except:
            raise Exception("Unable to delete the existing QA file")
