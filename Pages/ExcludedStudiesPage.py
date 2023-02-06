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

    """Constructor of the ExcludedStudies Page class"""
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
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    # def go_to_excludedstudies(self, locator, env):
    #     self.click(locator, env, UnivWaitFor=10)
    #     time.sleep(5)

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
        study_file = list(os.getcwd() + file['ExcludedStudies_Excel_Files'].dropna())
        study_filename = list(file['ExcludedStudies_Excel_File_names'].dropna())
        result = [[study_type[i], study_file[i], study_filename[i]] for i in range(0, len(study_type))]
        return result

    # Check the presence of Manage Excluded Studies option in Admin page
    def presence_of_elements(self, locator, env):
        self.scroll(locator, env)
        self.wait.until(ec.presence_of_element_located((getattr(By, self.locatortype(locator, env)),
                                                        self.locatorpath(locator, env))))
        self.LogScreenshot.fLogScreenshot(message=f'Manage Excluded Studies option is present in Admin page.',
                                          pass_=True, log=True, screenshot=True)

    # Check Manage Excluded Studies page elements are accessible or not
    def access_excludedstudy_page_elements(self, locatorname, filepath, env):
        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        for i in new_pop_data:
            for j in stdy_data:
                time.sleep(2)
                self.click("ex_stdy_pop_dropdown", env)
                self.LogScreenshot.fLogScreenshot(message=f"Population dropdown is accessible. Listed elements are:",
                                                  pass_=True, log=True, screenshot=True)
                pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                select1 = Select(pop_ele)
                select1.select_by_visible_text(i[0])
                time.sleep(1)

                self.click("ex_stdy_stdytype_dropdown", env)
                self.LogScreenshot.fLogScreenshot(message=f"SLR Type dropdown is accessible. Listed elements are:",
                                                  pass_=True, log=True, screenshot=True)
                stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env)
                select2 = Select(stdy_ele)
                select2.select_by_visible_text(j[0])
                time.sleep(1)

                self.click("ex_stdy_update_dropdown", env)
                self.LogScreenshot.fLogScreenshot(message=f"Updated dropdown is accessible. Listed elements are:",
                                                  pass_=True, log=True, screenshot=True)
                update_ele = self.select_element("ex_stdy_update_dropdown", env)
                select3 = Select(update_ele)
                select3.select_by_index(1)
                time.sleep(1)

                self.input_text("ex_stdy_file_upload", j[1], env)
                time.sleep(2)

                if self.clickable("ex_stdy_upload_button", env):
                    self.LogScreenshot.fLogScreenshot(message=f"Upload button is clickable after selecting the file",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Upload button is not clickable after selecting the "
                                                              f"file",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Upload button is not clickable after selecting the file")

    def add_multiple_excluded_study_data(self, locatorname, filepath, env):
        expected_upload_status_text = 'File(s) uploaded successfully'
        # Read the username
        username = self.get_text("get_user_name", env, UnivWaitFor=10)
        firstname = username.split()[0]

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for i in new_pop_data:
                for j in stdy_data:
                    expected_table_values = []
                    time.sleep(2)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    expected_table_values.append(select1.first_selected_option.text)
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env, UnivWaitFor=10)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(j[0])
                    expected_table_values.append(select2.first_selected_option.text)
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown", env, UnivWaitFor=10)
                    select3 = Select(update_ele)
                    time.sleep(1)
                    select3.select_by_index(1)
                    expected_table_values.append(select3.first_selected_option.text)
                    time.sleep(1)

                    self.input_text("ex_stdy_file_upload", j[1], env)
                    expected_table_values.append(j[2])
                    time.sleep(2)

                    self.click("ex_stdy_upload_button", env)
                    time.sleep(4)
                    actual_upload_status_text = self.get_text("ex_stdy_status_text", env, UnivWaitFor=10)
                    # time.sleep(1)

                    if actual_upload_status_text == expected_upload_status_text:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded Studies File upload is success for '
                                                                  f'Population : {i[0]} -> SLR Type : {j[0]}.',
                                                          pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading '
                                                                  f'Excluded Studies File for Population : {i[0]} -> '
                                                                  f'SLR Type: {j[0]}.',
                                                          pass_=False, log=True, screenshot=True)
                        raise Exception("Unable to find status message during Excluded Studies file uploading")

                    # Add the firstname to expected values list
                    expected_table_values.append(firstname)

                    # Read table data for specific population and slr study type
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    time.sleep(1)
                    select1.select_by_visible_text(i[0])
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(j[0])
                    time.sleep(1)

                    actual_table_values = []
                    td1 = self.select_elements('ex_stdy_table_data_row_1', env)
                    for m in td1:
                        actual_table_values.append(m.text)

                    for n in expected_table_values:
                        if n in actual_table_values:
                            self.LogScreenshot.fLogScreenshot(message=f'Correct data is present in table.',
                                                              pass_=True, log=True, screenshot=True)
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f'Mismatch is found in table data. '
                                                                      f'Expected values are : {expected_table_values} '
                                                                      f'and Actual values are : {actual_table_values}',
                                                              pass_=False, log=True, screenshot=True)
                            raise Exception(f'Mismatch is found in table data. Expected values are : '
                                            f'{expected_table_values} and Actual values are : {actual_table_values}')
        except Exception:
            raise Exception("Unable to upload the Excluded Studies data")

    def update_multiple_excluded_study_data(self, locatorname, filepath, env):
        expected_upload_status_text = 'File(s) uploaded successfully'
        # Read the username
        username = self.get_text("get_user_name", env, UnivWaitFor=10)
        firstname = username.split()[0]

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details_override(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for i in new_pop_data:
                for j in stdy_data:
                    expected_table_values = []
                    time.sleep(2)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    expected_table_values.append(select1.first_selected_option.text)
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env, UnivWaitFor=10)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(j[0])
                    expected_table_values.append(select2.first_selected_option.text)
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown", env, UnivWaitFor=10)
                    select3 = Select(update_ele)
                    time.sleep(1)
                    select3.select_by_index(1)
                    expected_table_values.append(select3.first_selected_option.text)
                    time.sleep(1)

                    self.input_text("ex_stdy_file_upload", j[1], env)
                    expected_table_values.append(j[2])
                    time.sleep(2)

                    self.click("ex_stdy_upload_button", env)
                    time.sleep(3)
                    self.jsclick("ex_stdy_popup_ok", env, message="Expected : popup reminder. Actual : popup is "
                                                                  "not shown")
                    time.sleep(3)
                    actual_upload_status_text = self.get_text("ex_stdy_status_text", env, UnivWaitFor=10)
                    # time.sleep(1)

                    if actual_upload_status_text == expected_upload_status_text:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'For Population : {i[0]} -> SLR Type : {j[0]}, updating the existing Excluded '
                                    f'Studies File is success.',
                            pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'For Population : {i[0]} -> SLR Type : {j[0]}, Unable to find status message '
                                    f'while updating the existing Excluded Studies File.',
                            pass_=False, log=True, screenshot=True)
                        raise Exception("Unable to find status message while Updating the existing Excluded Studies "
                                        "file")

                    # Add the firstname to expected values list
                    expected_table_values.append(firstname)

                    # Read table data for specific population and slr study type
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    time.sleep(1)
                    select1.select_by_visible_text(i[0])
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(j[0])
                    time.sleep(1)

                    actual_table_values = []
                    td1 = self.select_elements('ex_stdy_table_data_row_1', env)
                    for m in td1:
                        actual_table_values.append(m.text)

                    for n in expected_table_values:
                        if n in actual_table_values:
                            self.LogScreenshot.fLogScreenshot(message=f'Updated data is present in table.',
                                                              pass_=True, log=True, screenshot=True)
                        else:
                            self.LogScreenshot.fLogScreenshot(
                                message=f'Mismatch is found in table data. Expected values are : '
                                        f'{expected_table_values} and Actual values are : {actual_table_values}',
                                pass_=False, log=True, screenshot=True)
                            raise Exception(
                                f'Mismatch is found in table data. Expected values are : {expected_table_values} '
                                f'and Actual values are : {actual_table_values}')
        except Exception:
            raise Exception("Unable to update the existing Excluded Studies data")

    def del_multiple_excluded_study_data(self, locatorname, filepath, env):
        expected_delete_status_text = 'Excluded studies deleted successfully'

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details_override(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for i in new_pop_data:
                for j in stdy_data:
                    time.sleep(2)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(i[0])
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env, UnivWaitFor=10)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(j[0])
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown", env, UnivWaitFor=10)
                    select3 = Select(update_ele)
                    time.sleep(1)
                    select3.select_by_index(1)
                    time.sleep(1)

                    self.LogScreenshot.fLogScreenshot(message=f'Data selected for deletion is : ',
                                                      pass_=True, log=True, screenshot=True)

                    self.click("ex_stdy_delete", env)
                    time.sleep(2)
                    self.click("ex_stdy_popup_ok", env)
                    time.sleep(2)

                    actual_delete_status_text = self.get_text("ex_stdy_status_text", env, UnivWaitFor=10)
                    # time.sleep(2)

                    if actual_delete_status_text == expected_delete_status_text:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded Studies File Deletion is success.',
                                                          pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'Unable to find status message while deleting Excluded Studies File.',
                            pass_=False, log=True, screenshot=True)
                        raise Exception("Error in Excluded Studies File Deletion")
        except Exception:
            raise Exception("Unable to delete the existing Excluded Studies File")

    def compare_excludedstudy_file_with_report(self, filepath, locatorname, env):
        expected_upload_status_text = 'File(s) uploaded successfully'
        # Read the username
        username = self.get_text("get_user_name", env, UnivWaitFor=10)
        firstname = username.split()[0]

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for pop_val in new_pop_data:
                for i in stdy_data:
                    expected_table_values = []
                    self.go_to_page("excluded_studies_link", env)
                    time.sleep(2)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(pop_val[0])
                    expected_table_values.append(select1.first_selected_option.text)
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env, UnivWaitFor=10)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(i[0])
                    expected_table_values.append(select2.first_selected_option.text)
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown", env, UnivWaitFor=10)
                    select3 = Select(update_ele)
                    time.sleep(1)
                    select3.select_by_index(1)
                    expected_table_values.append(select3.first_selected_option.text)
                    time.sleep(1)

                    self.input_text("ex_stdy_file_upload", i[1], env)
                    expected_table_values.append(i[2])
                    time.sleep(2)

                    self.click("ex_stdy_upload_button", env)
                    time.sleep(2)

                    if self.isdisplayed("ex_stdy_status_text", env):
                        actual_upload_status_text = self.get_text("ex_stdy_status_text", env, UnivWaitFor=30)
                    else:
                        time.sleep(2)
                        actual_upload_status_text = self.get_text("ex_stdy_status_text", env, UnivWaitFor=30)

                    if actual_upload_status_text == expected_upload_status_text:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'Excluded Studies File upload is success for Population : {pop_val[0]} -> '
                                    f'SLR Type : {i[0]}.',
                            pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'Unable to find status message while uploading Excluded Studies File for '
                                    f'Population : {pop_val[0]} -> SLR Type: {i[0]}.',
                            pass_=False, log=True, screenshot=True)
                        raise Exception("Unable to find status message during Excluded Studies file uploading")

                    # Add the firstname to expected values list
                    expected_table_values.append(firstname)

                    # Read table data for specific population and slr study type
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    time.sleep(1)
                    select1.select_by_visible_text(pop_val[0])
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(i[0])
                    time.sleep(2)

                    actual_table_values = []
                    td1 = self.select_elements('ex_stdy_table_data_row_1', env)
                    time.sleep(1)
                    for m in td1:
                        actual_table_values.append(m.text)

                    for n in expected_table_values:
                        if n in actual_table_values:
                            self.LogScreenshot.fLogScreenshot(message=f'Correct data is present in table.',
                                                              pass_=True, log=True, screenshot=True)
                        else:
                            self.LogScreenshot.fLogScreenshot(
                                message=f'Mismatch is found in table data. Expected values are : '
                                        f'{expected_table_values} and Actual values are : {actual_table_values}',
                                pass_=False, log=True, screenshot=True)
                            raise Exception(
                                f'Mismatch is found in table data. Expected values are : {expected_table_values} and '
                                f'Actual values are : {actual_table_values}')

                    # Go to live slr page
                    self.go_to_page("SLR_Homepage", env)
                    self.slrreport.select_data(f"{pop_val[0]}", f"{pop_val[0]}_radio_button", env)
                    self.slrreport.select_data(i[0], f"{i[0]}_radio_button", env)
                    self.slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename = self.slrreport.getFilenameAndValidate(180)
                    # excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    excel_filename = self.slrreport.get_and_validate_filename(filepath)

                    update_date_val = expected_table_values[2].translate({ord('/'): None})[-4:] + expected_table_values[
                                                                                                    2].translate(
                        {ord('/'): None})[:4]

                    excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
                    if f'Excluded studies {update_date_val}' in excel_data.sheetnames:
                        self.LogScreenshot.fLogScreenshot(
                            message=f"'Excluded studies {update_date_val}' sheet is present in complete excel report",
                            pass_=True, log=True, screenshot=False)

                        excel_sheet = excel_data[f'Excluded studies {update_date_val}']
                        if excel_sheet['A1'].value == 'Back To Toc':
                            self.LogScreenshot.fLogScreenshot(
                                message=f"'Back To Toc' option is present in 'Excluded studies {update_date_val}' "
                                        f"sheet",
                                pass_=True, log=True, screenshot=False)
                        else:
                            self.LogScreenshot.fLogScreenshot(
                                message=f"'Back To Toc' option is not present in 'Excluded studies {update_date_val}' "
                                        f"sheet",
                                pass_=False, log=True, screenshot=False)
                            raise Exception(
                                f"'Back To Toc' option is not present in 'Excluded studies {update_date_val}' sheet")

                        toc_sheet = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name="TOC", skiprows=3)
                        col_data = list(toc_sheet.iloc[:, 1])
                        if f'Excluded studies {update_date_val}' in col_data:
                            self.LogScreenshot.fLogScreenshot(message=f'Excluded studies is present in TOC sheet.',
                                                              pass_=True, log=True, screenshot=False)
                        else:
                            self.LogScreenshot.fLogScreenshot(
                                message=f'Excluded studies is not present in TOC sheet. Available Data from TOC sheet: '
                                        f'{col_data}',
                                pass_=False, log=True, screenshot=False)
                            raise Exception("'Excluded studies' is not present in TOC sheet.")

                        studyfile = pd.read_excel(i[1])
                        excelfile = pd.read_excel(f'ActualOutputs//{excel_filename}',
                                                  sheet_name=f'Excluded studies {update_date_val}')

                        # Removing the 'Back To Toc' column to compare the exact data with uploaded file
                        excelfile = excelfile.iloc[:, 1:]

                        if studyfile.equals(excelfile):
                            self.LogScreenshot.fLogScreenshot(
                                message=f"File contents between QA File '{Path(f'{i[1]}').stem}' and Complete Excel "
                                        f"Report '{Path(f'ActualOutputs//{excel_filename}').stem}' are matching",
                                pass_=True, log=True, screenshot=False)
                        else:
                            raise Exception(
                                f"File contents between Study File '{Path(f'{i[1]}').stem}' and Complete Excel Report "
                                f"'{Path(f'ActualOutputs//{excel_filename}').stem}' are not matching")
                    else:
                        raise Exception("'Excluded studies' sheet is not present in complete excel report")
        except Exception:
            raise Exception("Error in report comparision between Excluded study file and Complete Excel report")

    def del_after_studyfile_comparison(self, filepath, locatorname, env):
        expected_delete_status_text = 'Excluded studies deleted successfully'

        # Read study types and file paths to upload
        stdy_data = self.get_study_file_details(filepath, locatorname)
        # Read population details from data sheet
        new_pop_data = self.get_pop_data(filepath, locatorname)

        try:
            for pop_val in new_pop_data:
                for i in stdy_data:
                    expected_table_values = []
                    self.go_to_page("excluded_studies_link", env)
                    time.sleep(2)
                    pop_ele = self.select_element("ex_stdy_pop_dropdown", env)
                    select1 = Select(pop_ele)
                    select1.select_by_visible_text(pop_val[0])
                    expected_table_values.append(select1.first_selected_option.text)
                    time.sleep(1)

                    stdy_ele = self.select_element("ex_stdy_stdytype_dropdown", env, UnivWaitFor=10)
                    select2 = Select(stdy_ele)
                    time.sleep(1)
                    select2.select_by_visible_text(i[0])
                    expected_table_values.append(select2.first_selected_option.text)
                    time.sleep(1)

                    update_ele = self.select_element("ex_stdy_update_dropdown", env, UnivWaitFor=10)
                    select3 = Select(update_ele)
                    time.sleep(1)
                    select3.select_by_index(1)
                    expected_table_values.append(select3.first_selected_option.text)
                    time.sleep(1)

                    self.LogScreenshot.fLogScreenshot(message=f'Data selected for deletion is : ',
                                                      pass_=True, log=True, screenshot=True)

                    self.click("ex_stdy_delete", env)
                    time.sleep(2)
                    self.click("ex_stdy_popup_ok", env)
                    time.sleep(2)

                    if self.isdisplayed("get_status_text", env):
                        actual_delete_status_text = self.get_text("get_status_text", env, UnivWaitFor=30)
                    else:
                        time.sleep(2)
                        actual_delete_status_text = self.get_text("get_status_text", env, UnivWaitFor=30)                    

                    if actual_delete_status_text == expected_delete_status_text:
                        self.LogScreenshot.fLogScreenshot(message=f'Excluded Studies File Deletion is success.',
                                                          pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'Error while deleting Excluded Studies File. Error Message is '
                                    f'{actual_delete_status_text}',
                            pass_=False, log=True, screenshot=True)
                        raise Exception("Error in Excluded Studies File Deletion")

                    # Go to live slr page
                    self.go_to_page("SLR_Homepage", env)
                    self.slrreport.select_data(f"{pop_val[0]}", f"{pop_val[0]}_radio_button", env)
                    self.slrreport.select_data(i[0], f"{i[0]}_radio_button", env)
                    self.slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename = self.slrreport.getFilenameAndValidate(180)
                    # excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    excel_filename = self.slrreport.get_and_validate_filename(filepath)

                    update_date_val = expected_table_values[2].translate({ord('/'): None})[-4:] + expected_table_values[
                                                                                                    2].translate(
                        {ord('/'): None})[:4]

                    excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
                    if f'Excluded studies {update_date_val}' not in excel_data.sheetnames:
                        self.LogScreenshot.fLogScreenshot(
                            message=f"'Excluded studies' sheet is not present in complete excel report as expected",
                            pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f"'Excluded studies' sheet is present in complete excel report which is not "
                                    f"expected", pass_=False, log=True, screenshot=False)
                        raise Exception(
                            "'Excluded studies' sheet is present in complete excel report which is not expected")

                    toc_sheet = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name="TOC", skiprows=3)
                    col_data = list(toc_sheet.iloc[:, 1])
                    if f'Excluded studies {update_date_val}' not in col_data:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'Excluded studies is not present in TOC sheet as expected.',
                            pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f'Excluded studies is present in TOC sheet which is not expected. '
                                    f'Available Data from TOC sheet: {col_data}',
                            pass_=False, log=True, screenshot=False)
                        raise Exception("'Excluded studies' is present in TOC sheet.")
        except Exception:
            raise Exception("Unable to delete the existing Excluded Studies File")
