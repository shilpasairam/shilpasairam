from datetime import date, datetime, timedelta
import os
import time
import pandas as pd
import numpy as np
from selenium.webdriver.support.wait import WebDriverWait

from pathlib import Path
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.SLRReportPage import SLRReport
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support.ui import Select

from utilities.readProperties import ReadConfig


class ProtocolPage(Base):

    """Constructor of the Protocol Page class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of ExtendedBase class
        self.exbase = ExtendedBase(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)                 
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    # Reading Population data for PRISMAs Page
    def get_prisma_pop_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop = df.loc[df['Name'] == locatorname]['Prisma_Population'].dropna().to_list()
        return pop
    
    def get_prisma_excelfile_details(self, filepath, locatorname, column_name):
        df = pd.read_excel(filepath)
        sheet_name = os.getcwd()+(df.loc[df['Name'] == locatorname][column_name].to_list()[0])
        return sheet_name
    
    def get_prisma_data_details(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        sheet_name = df.loc[df['Name'] == locatorname]['Study_Types'].to_list()
        sheet1 = df.loc[df['Name'] == locatorname]['OriginalStudiesNumbers'].to_list()
        sheet2 = df.loc[df['Name'] == locatorname]['RecordsNumber'].to_list()
        sheet3 = df.loc[df['Name'] == locatorname]['fullTextReviewRecordsNumber'].to_list()
        sheet4 = df.loc[df['Name'] == locatorname]['totalRecordsNumber'].to_list()
        sheet5 = df.loc[df['Name'] == locatorname]['Prisma_Image'].to_list()
        prisma_data = [(sheet_name[i], int(sheet1[i]), int(sheet2[i]), int(sheet3[i]), int(sheet4[i]),
                        (os.getcwd()+sheet5[i])) for i in range(0, len(sheet_name))]
        return prisma_data

    def add_prisma_excel_file(self, locatorname, filepath, env):
        expected_excel_upload_status_text = "PRISMA successfully uploaded"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)
        
        # self.refreshpage()
        # time.sleep(3)
        try:
            # pop_ele = self.select_element("prisma_pop_dropdown", env)
            # select = Select(pop_ele)
            # select.select_by_visible_text(pop_name[0])
            selected_pop_val = self.base.selectbyvisibletext("prisma_pop_dropdown", pop_name[0], env)

            # Read the PRISMA file path required to upload
            prisma_excel = self.get_prisma_excelfile_details(filepath, locatorname, 'Prisma_Excel_File')

            self.input_text("prisma_excel_file", prisma_excel, env, UnivWaitFor=10)
            self.click("prisma_excel_upload_btn", env)
            time.sleep(2)
            # actual_excel_upload_status_text = self.get_text("prisma_excel_status_text", env, UnivWaitFor=30)
            actual_excel_upload_status_text = self.get_status_text("prisma_excel_status_text", env)
            # time.sleep(2)

            if actual_excel_upload_status_text == expected_excel_upload_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"PRISMA Excel File Upload is success.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading "
                                                          f"PRISMA Excel File. Actual status message is "
                                                          f"{actual_excel_upload_status_text} and Expected status "
                                                          f"message is {expected_excel_upload_status_text}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message while uploading PRISMA Excel File.")
        except Exception:
            raise Exception("Unable to upload PRISMA Excel file")
    
    def del_prisma_excel_file(self, locatorname, excel_del_locator, excel_del_popup, filepath, env):
        expected_excel_del_status_text = "PRISMA excel file successfully deleted"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)
        
        self.refreshpage()
        time.sleep(3)
        try:
            # pop_ele = self.select_element("prisma_pop_dropdown", env)
            # select = Select(pop_ele)
            # select.select_by_visible_text(pop_name[0])
            selected_pop_val = self.base.selectbyvisibletext("prisma_pop_dropdown", pop_name[0], env)

            self.del_prisma_image(locatorname, filepath, "prisma_image_delete_btn", "prisma_image_delete_popup", env)
            time.sleep(2)

            self.click(excel_del_locator, env)
            time.sleep(2)
            self.click(excel_del_popup, env)
            time.sleep(2)

            # actual_excel_del_status_text = self.get_text("prisma_excel_status_text", env, UnivWaitFor=30)
            actual_excel_del_status_text = self.get_status_text("prisma_excel_status_text", env)
            # time.sleep(2)

            if actual_excel_del_status_text == expected_excel_del_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"PRISMA Excel File Delete is success.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting "
                                                          f"PRISMA Excel File. Actual status message is "
                                                          f"{actual_excel_del_status_text} and Expected status "
                                                          f"message is {expected_excel_del_status_text}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message while deleting PRISMA Excel File.")
            time.sleep(5)
        except Exception:
            raise Exception("Unable to delete PRISMA Excel file")

    def upload_prisma_image(self, locatorname, filepath, index, env):
        expected_image_upload_status_text = "PRISMA successfully updated"

        # Read the Data required to upload for study types
        stdy_data = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'Study_Types', 'Prisma_Image')
        prisma_numbers = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'stdy_type_locators', 'stdy_type_values')

        try:
            for i in stdy_data:
                # stdy_ele = self.select_element("prisma_study_type_dropdown", env)
                # select = Select(stdy_ele)
                # select.select_by_visible_text(i[0])
                selected_slr_val = self.base.selectbyvisibletext("prisma_study_type_dropdown", i[0], env)
                time.sleep(2)

                # self.input_text("original_studies", i[1]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("records_number", i[2]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("fulltext_review", i[3]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("total_record_number", i[4]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("prisma_image", i[5], env, UnivWaitFor=5)
                # time.sleep(1)
                for j in prisma_numbers:
                    self.input_text(j[0], j[1]+index, env, UnivWaitFor=5)

                self.input_text("prisma_image", os.getcwd()+i[1], env, UnivWaitFor=5)
                self.click("prisma_image_save_btn", env)
                time.sleep(2)
                
                # actual_image_upload_status_text = self.get_text("prisma_image_status_text", env, UnivWaitFor=30)
                actual_image_upload_status_text = self.get_status_text("prisma_image_status_text", env)
                # time.sleep(2)

                if actual_image_upload_status_text == expected_image_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Upload is success.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading "
                                                              f"PRISMA Image File. Actual status message is "
                                                              f"{actual_image_upload_status_text} and Expected "
                                                              f"status message is {expected_image_upload_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message while uploading PRISMA Image File.")
        except Exception:
            raise Exception("Unable to upload PRISMA image")
    
    def del_prisma_image(self, locatorname, filepath, del_locator, del_popup, env):
        expected_image_del_status_text = "PRISMA successfully deleted"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)

        # stdy_data = self.get_prisma_data_details(filepath, locatorname)

        stdy_data = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'Study_Types', 'Prisma_Image')

        try:
            for i in stdy_data:
                self.refreshpage()
                time.sleep(3)
                # pop_ele = self.select_element("prisma_pop_dropdown", env)
                # select = Select(pop_ele)
                # select.select_by_visible_text(pop_name[0])
                selected_pop_val = self.base.selectbyvisibletext("prisma_pop_dropdown", pop_name[0], env)
                time.sleep(1)

                # stdy_ele = self.select_element("prisma_study_type_dropdown", env)
                # select = Select(stdy_ele)
                # select.select_by_visible_text(i[0])
                selected_slr_val = self.base.selectbyvisibletext("prisma_study_type_dropdown", i[0], env)
                time.sleep(1)

                self.click(del_locator, env)
                time.sleep(2)
                self.click(del_popup, env)
                time.sleep(2)
                
                # actual_image_del_status_text = self.get_text("prisma_image_status_text", env, UnivWaitFor=30)
                actual_image_del_status_text = self.get_status_text("prisma_image_status_text", env)
                # time.sleep(2)

                if actual_image_del_status_text == expected_image_del_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Delete is success.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting "
                                                              f"PRISMA Image File. Actual status message is "
                                                              f"{actual_image_del_status_text} and Expected status "
                                                              f"message is {expected_image_del_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message while deleting PRISMA Image File.")
                time.sleep(5)
        except Exception:
            raise Exception("Unable to delete PRISMA image")

    def override_prisma_details(self, locatorname, filepath, index, env):
        expected_excel_upload_status_text = "There is an existing PRISMA Excel file for this population. " \
                                            "Please delete the existing file before uploading a new one"

        expected_image_upload_status_text = "There is an existing PRISMA Image file for this population. " \
                                            "Please delete the existing file before uploading a new one"
        
        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)

        # Read the Data required to upload for study types
        # stdy_data = self.get_prisma_data_details(filepath, locatorname)
        stdy_data = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'Study_Types', 'Prisma_Image')
        prisma_numbers = self.exbase.get_double_col_data(filepath, locatorname, 'Sheet1', 'stdy_type_locators', 'stdy_type_values')

        self.refreshpage()
        
        try:
            # pop_ele = self.select_element("prisma_pop_dropdown", env)
            # select = Select(pop_ele)
            # select.select_by_visible_text(pop_name[0])
            selected_pop_val = self.base.selectbyvisibletext("prisma_pop_dropdown", pop_name[0], env)

            # Read the PRISMA file path required to upload
            prisma_excel = self.get_prisma_excelfile_details(filepath, locatorname, 'Prisma_Excel_File')

            self.input_text("prisma_excel_file", prisma_excel, env, UnivWaitFor=10)
            self.click("prisma_excel_upload_btn", env)
            time.sleep(2)
            
            actual_excel_upload_status_text = self.get_status_text("prisma_excel_status_text", env)

            if actual_excel_upload_status_text == expected_excel_upload_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"There is an existing PRISMA Excel file for Population : "
                                                          f"'{pop_name[0]}'", pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading "
                                                          f"PRISMA Excel File. Actual status message is "
                                                          f"{actual_excel_upload_status_text} and Expected status "
                                                          f"message is {expected_excel_upload_status_text}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message while uploading PRISMA Excel File.")

            for i in stdy_data:
                # stdy_ele = self.select_element("prisma_study_type_dropdown", env)
                # select = Select(stdy_ele)
                # select.select_by_visible_text(i[0])
                selected_slr_val = self.base.selectbyvisibletext("prisma_study_type_dropdown", i[0], env)
                time.sleep(2)

                # self.input_text("original_studies", i[1]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("records_number", i[2]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("fulltext_review", i[3]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("total_record_number", i[4]+index, env, UnivWaitFor=5)
                # time.sleep(1)
                # self.input_text("prisma_image", i[5], env, UnivWaitFor=5)
                # time.sleep(1)
                for j in prisma_numbers:
                    self.input_text(j[0], j[1]+index, env, UnivWaitFor=5)

                self.input_text("prisma_image", os.getcwd()+i[1], env, UnivWaitFor=5)
                self.click("prisma_image_save_btn", env)
                time.sleep(2)

                actual_image_upload_status_text = self.get_status_text("prisma_image_status_text", env)

                if actual_image_upload_status_text == expected_image_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"There is an existing PRISMA Image file for Population "
                                                              f": '{pop_name[0]}' -> SLR Type : '{i[0]}'",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading "
                                                              f"PRISMA Image File. Actual status message is "
                                                              f"{actual_image_upload_status_text} and Expected "
                                                              f"status message is {expected_image_upload_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message while uploading PRISMA Image File.")
        except Exception:
            raise Exception("Unable to upload PRISMA Excel file")

    def add_picos_details(self, locatorname, filepath, env):
        expected_status_text = "Saved successfully"

        # Read population details from data sheet
        # pop_name = self.get_prisma_pop_data(filepath, locatorname)
        pop_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Population')
        stdy_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Study_Types')
        expected_row_headers = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Row_headers')
        expected_col_headers = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Col_headers')

        time.sleep(2)
        try:
            for j in stdy_data:
                # pop_ele = self.select_element("picos_pop_dropdown", env)
                # select = Select(pop_ele)
                # select.select_by_visible_text(pop_data[0])
                selected_pop_val = self.base.selectbyvisibletext("picos_pop_dropdown", pop_data[0], env)
                time.sleep(1)

                # stdy_ele = self.select_element("picos_study_type_dropdown", env)
                # select = Select(stdy_ele)
                # select.select_by_visible_text(j)
                selected_slr_val = self.base.selectbyvisibletext("picos_study_type_dropdown", j, env)
                time.sleep(1)

                actual_row_headers = self.get_texts("row_header_values", env)

                actual_col_headers = self.get_texts("col_header_values", env)
                # Removing the empty column name
                actual_col_headers.pop(0)

                row_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_row_headers,
                                                                                            actual_row_headers)

                if len(row_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page row headers are displayed as expected for"
                                                              f" Population -> {pop_data[0]}, Study Type -> {j}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page row headers for "
                                                              f"Population -> {pop_data[0]}, Study Type -> {j}. "
                                                              f"Mismatch values are arranged in following order -> "
                                                              f"Expected row headers, Actual row headers. "
                                                              f"{row_header_comparison}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {pop_data[0]}, "
                                    f"Study Type -> {j}.")

                col_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_col_headers,
                                                                                            actual_col_headers)

                if len(col_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page col headers are displayed as expected for "
                                                              f"Population -> {pop_data[0]}, Study Type -> {j}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page col headers for "
                                                              f"Population -> {pop_data[0]}, Study Type -> {j}. "
                                                              f"Mismatch values are arranged in following order -> "
                                                              f"Expected col headers, Actual col headers. "
                                                              f"{col_header_comparison}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {pop_data[0]}, "
                                    f"Study Type -> {j}.")

                # Enter values in PICOS page
                data_eles = self.select_elements('row_data', env)
                for index, locator in enumerate(data_eles):
                    # self.input_text(locator, f"Test_Automation_{index}", env)
                    locator.clear()
                    locator.send_keys(f"Test_Automation_{index}")

                self.scrollback('picos_page_heading', env)
                self.jsclick("picos_save_btn", env)
                time.sleep(1)
                # actual_excel_upload_status_text = self.get_text("prisma_excel_status_text", env, UnivWaitFor=30)
                actual_status_text = self.get_status_text("prisma_excel_status_text", env)
                # time.sleep(2)

                if actual_status_text == expected_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"Addition of PICOS data is success for Population -> "
                                                              f"{pop_data[0]}, Study Type -> {j}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding PICOS data "
                                                              f"for Population -> {pop_data[0]}, Study Type -> {j}. "
                                                              f"Actual status message is {actual_status_text} and "
                                                              f"Expected status message is {expected_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding PICOS data for Population -> "
                                    f"{pop_data[0]}, Study Type -> {j}.")
                
                self.refreshpage()
                time.sleep(2)
        except Exception:
            raise Exception("Unable to add PICOS data")

    def validate_view_picos_oncology(self, locatorname, filepath, env):
        expected_status_text = "Saved successfully"

        # Read population data values
        pop_list = self.exbase.get_population_data(filepath, 'Sheet1', locatorname)
        # Read slrtype data values
        slrtype = self.exbase.get_slrtype_data(filepath, 'Sheet1', locatorname)

        expected_row_headers = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Row_headers')
        expected_col_headers = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Col_headers')
        data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'data')

        for i in pop_list:
            for index, j in enumerate(slrtype):
                self.base.presence_of_admin_page_option("protocol_link", env)
                self.base.go_to_nested_page("protocol_link", "picos", env)

                selected_pop_val = self.base.selectbyvisibletext("picos_pop_dropdown", i[0], env)
                time.sleep(1)

                if j[0] == 'Quality of Life':
                    j[0] = 'Quality of life'
                selected_slr_val = self.base.selectbyvisibletext("picos_study_type_dropdown", j[0], env)
                time.sleep(1)

                actual_row_headers = self.get_texts("row_header_values", env)

                actual_col_headers = self.get_texts("col_header_values", env)
                # Removing the empty column name
                actual_col_headers.pop(0)

                row_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_row_headers,
                                                                                            actual_row_headers)

                if len(row_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page row headers are displayed as expected for"
                                                              f" Population -> {i[0]}, Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page row headers for "
                                                              f"Population -> {i[0]}, Study Type -> {j[0]}. "
                                                              f"Mismatch values are arranged in following order -> "
                                                              f"Expected row headers, Actual row headers. "
                                                              f"{row_header_comparison}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {i[0]}, "
                                    f"Study Type -> {j[0]}.")

                col_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_col_headers,
                                                                                            actual_col_headers)

                if len(col_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page col headers are displayed as expected for "
                                                              f"Population -> {i[0]}, Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page col headers for "
                                                              f"Population -> {i[0]}, Study Type -> {j[0]}. "
                                                              f"Mismatch values are arranged in following order -> "
                                                              f"Expected col headers, Actual col headers. "
                                                              f"{col_header_comparison}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {i[0]}, "
                                    f"Study Type -> {j[0]}.")

                # Enter values in PICOS page
                picos_data = []
                data_eles = self.select_elements('row_data', env)
                idx = 0
                for locator in data_eles:
                    if idx > 2:
                        idx = 0
                    locator.clear()
                    locator.send_keys(f"{data[idx]}")
                    picos_data.append(data[idx].split('\n'))
                    idx += 1
                
                # This will give one single list with all the picos data added
                protocol_picos_data = list(np.concatenate(picos_data))

                self.scrollback('picos_page_heading', env)
                self.jsclick("picos_save_btn", env)
                time.sleep(1)

                actual_status_text = self.get_status_text("prisma_excel_status_text", env)

                if actual_status_text == expected_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"Addition of PICOS data is success for Population -> "
                                                              f"{i[0]}, Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding PICOS data "
                                                              f"for Population -> {i[0]}, Study Type -> {j[0]}. "
                                                              f"Actual status message is {actual_status_text} and "
                                                              f"Expected status message is {expected_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding PICOS data for Population -> "
                                    f"{i[0]}, Study Type -> {j[0]}.")                
                self.refreshpage()
                time.sleep(2)                

                self.base.go_to_page("SLR_Homepage", env)
                self.slrreport.select_data(i[0], i[1], env)

                if j[0] == 'Quality of life':
                    j[0] = 'Quality of Life'
                self.slrreport.select_data(j[0], j[1], env)

                self.click("view_picos_button", env)
                time.sleep(1)

                view_picos_data = self.get_texts("view_picos_data", env)

                picos_data_comparison = self.slrreport.list_comparison_between_reports_data(protocol_picos_data,
                                                                                            view_picos_data)

                if len(picos_data_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS data under Search LiveSLR -> View PICOS section "
                                                              f"is matching with PICOS data under Protocol -> PICOS "
                                                              f"and Inc-Exc criteria page for Population -> {i[0]}, "
                                                              f"Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS data under Search LiveSLR -> "
                                                              f"View PICOS section with PICOS data under Protocol -> "
                                                              f"PICOS and Inc-Exc criteria page for Population -> "
                                                              f"{i[0]}, Study Type -> {j[0]}. Mismatch values are "
                                                              f"arranged in following order -> Protocol PICOS Data, "
                                                              f"View PICOS Data. {col_header_comparison}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Mismatch found in PICOS data under Search LiveSLR -> View PICOS section with "
                                    f"PICOS data under Protocol -> PICOS and Inc-Exc criteria page for Population -> "
                                    f"{i[0]}, Study Type -> {j[0]}.")
                self.refreshpage()

    def validate_view_picos_nononcology(self, locatorname, filepath, env):
        expected_status_text = "Saved successfully"

        # Read population data values
        pop_list = self.exbase.get_population_data(filepath, 'Sheet1', locatorname)
        # Read slrtype data values
        slrtype = self.exbase.get_slrtype_data(filepath, 'Sheet1', locatorname)

        expected_row_headers = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Row_headers')
        expected_col_headers = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Col_headers')
        # data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'data')

        for i in pop_list:
            for index, j in enumerate(slrtype):
                if j[0] == 'Clinical-Interventional':
                    data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'data_intervention')
                elif j[0] == 'Clinical-RWE':
                    data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'data_rwe')                
                self.base.presence_of_admin_page_option("protocol_link", env)
                self.base.go_to_nested_page("protocol_link", "picos", env)

                selected_pop_val = self.base.selectbyvisibletext("picos_pop_dropdown", i[0], env)
                time.sleep(1)

                if j[0] == 'Quality of Life':
                    j[0] = 'Quality of life'
                selected_slr_val = self.base.selectbyvisibletext("picos_study_type_dropdown", j[0], env)
                time.sleep(1)

                actual_row_headers = self.get_texts("row_header_values", env)

                actual_col_headers = self.get_texts("col_header_values", env)
                # Removing the empty column name
                actual_col_headers.pop(0)

                row_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_row_headers,
                                                                                            actual_row_headers)

                if len(row_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page row headers are displayed as expected for"
                                                              f" Population -> {i[0]}, Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page row headers for "
                                                              f"Population -> {i[0]}, Study Type -> {j[0]}. "
                                                              f"Mismatch values are arranged in following order -> "
                                                              f"Expected row headers, Actual row headers. "
                                                              f"{row_header_comparison}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {i[0]}, "
                                    f"Study Type -> {j[0]}.")

                col_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_col_headers,
                                                                                            actual_col_headers)

                if len(col_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page col headers are displayed as expected for "
                                                              f"Population -> {i[0]}, Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page col headers for "
                                                              f"Population -> {i[0]}, Study Type -> {j[0]}. "
                                                              f"Mismatch values are arranged in following order -> "
                                                              f"Expected col headers, Actual col headers. "
                                                              f"{col_header_comparison}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {i[0]}, "
                                    f"Study Type -> {j[0]}.")

                # Enter values in PICOS page
                picos_data = []
                data_eles = self.select_elements('row_data', env)
                idx = 0
                for locator in data_eles:
                    if idx > 3:
                        idx = 0
                    locator.clear()
                    locator.send_keys(f"{data[idx]}")
                    picos_data.append(data[idx].split('\n'))
                    idx += 1
                
                # This will give one single list with all the picos data added
                protocol_picos_data = list(np.concatenate(picos_data))

                self.scrollback('picos_page_heading', env)
                self.jsclick("picos_save_btn", env)
                time.sleep(1)

                actual_status_text = self.get_status_text("prisma_excel_status_text", env)

                if actual_status_text == expected_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"Addition of PICOS data is success for Population -> "
                                                              f"{i[0]}, Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding PICOS data "
                                                              f"for Population -> {i[0]}, Study Type -> {j[0]}. "
                                                              f"Actual status message is {actual_status_text} and "
                                                              f"Expected status message is {expected_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding PICOS data for Population -> "
                                    f"{i[0]}, Study Type -> {j[0]}.")                
                self.refreshpage()
                time.sleep(2)                

                self.base.go_to_page("SLR_Homepage", env)
                self.slrreport.select_data(i[0], i[1], env)

                temp = j[0]
                if j[0] == 'Clinical-Interventional' or j[0] == 'Clinical-RWE':
                    j[0] = 'Clinical'
                if j[0] == 'Quality of life':
                    j[0] = 'Quality of Life'
                self.slrreport.select_data(j[0], j[1], env)

                self.click("view_picos_button", env)
                time.sleep(1)

                if temp == 'Clinical-Interventional':
                    self.click("view_picos_clinical_intervention_tab", env)                   
                if temp == 'Clinical-RWE':
                    self.click("view_picos_clinical_rwe_tab", env)

                view_picos_data = self.get_texts("view_picos_active_tab_data", env)

                picos_data_comparison = self.slrreport.list_comparison_between_reports_data(protocol_picos_data,
                                                                                            view_picos_data)

                if len(picos_data_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS data under Search LiveSLR -> View PICOS section "
                                                              f"is matching with PICOS data under Protocol -> PICOS "
                                                              f"and Inc-Exc criteria page for Population -> {i[0]}, "
                                                              f"Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS data under Search LiveSLR -> "
                                                              f"View PICOS section with PICOS data under Protocol -> "
                                                              f"PICOS and Inc-Exc criteria page for Population -> "
                                                              f"{i[0]}, Study Type -> {j[0]}. Mismatch values are "
                                                              f"arranged in following order -> Protocol PICOS Data, "
                                                              f"View PICOS Data. {col_header_comparison}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Mismatch found in PICOS data under Search LiveSLR -> View PICOS section with "
                                    f"PICOS data under Protocol -> PICOS and Inc-Exc criteria page for Population -> "
                                    f"{i[0]}, Study Type -> {j[0]}.")
                self.refreshpage()

    def add_valid_search_strategy_details(self, locatorname, filepath, env, project):
        expected_excel_upload_status_text = "Search strategy updated successfully"

        today = date.today()
        # Manipulating the date values when day point to month end
        if today.day in [30, 31]:
            day_val = (today - timedelta(10)).strftime("%d")
        else:
            day_val = today.day

        # Read population details from data sheet
        pop_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Population')
        stdy_data = self.exbase.get_four_cols_data(filepath, locatorname, 'Sheet1', 'Study_Types', 'Files_to_upload',
                                                   'db_search_val', 'Template_name')
        template_name = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Template_name')

        self.refreshpage()
        time.sleep(2)
        try:
            for j in stdy_data:
                # pop_ele = self.select_element("searchstrategy_pop_dropdown", env)
                # select = Select(pop_ele)
                # select.select_by_visible_text(pop_data[0])
                selected_pop_val = self.base.selectbyvisibletext("searchstrategy_pop_dropdown", pop_data[0], env)
                time.sleep(1)

                # stdy_ele = self.select_element("searchstrategy_study_type_dropdown", env)
                # select = Select(stdy_ele)
                # select.select_by_visible_text(j[0])
                selected_slr_val = self.base.selectbyvisibletext("searchstrategy_study_type_dropdown", j[0], env)
                time.sleep(1)

                self.slrreport.generate_download_report("searchstrategy_template_download_btn", env)
                # Renaming the filename because there is an issue in downloading filenames with same name multiple
                # times in headless mode. So as an alternative renaming the file after downloading in each iteration
                self.file_rename(self.slrreport.get_latest_filename(UnivWaitFor=180),
                                 f"{j[0]}_search-strategy-template_{project}.xlsx")
                downloaded_template_name = self.slrreport.get_latest_filename(UnivWaitFor=180)                
                if downloaded_template_name == j[3]:
                    self.LogScreenshot.fLogScreenshot(message=f"Correct Template is downloaded. Template name is "
                                                              f"{downloaded_template_name}",
                                                      pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch in search strategy template name. Expected "
                                                              f"Template name is {template_name[0]} and Actual "
                                                              f"Template name is {downloaded_template_name}",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch in search strategy template name.")

                self.click("searchstrategy_date", env)
                self.select_calendar_date(day_val)

                self.input_text("searchstrategy_dbsearch", j[2], env)

                jscmd = ReadConfig.get_remove_att_JScommand(17, 'hidden')
                self.jsclick_hide(jscmd)
                self.input_text("searchstrategy_upload_file", os.getcwd()+"\\"+j[1], env)  # f"{os.getcwd()} + {j[1]}"

                self.jsclick("searchstrategy_upload_btn", env)
                time.sleep(2)

                actual_excel_upload_status_text = self.get_status_text("searchstrategy_status_text", env)
                # time.sleep(2)

                if actual_excel_upload_status_text == expected_excel_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"Addition of Search strategy data is success for "
                                                              f"Population -> {pop_data[0]}, Study Type -> {j[0]}.",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding "
                                                              f"Search strategy data for Population -> {pop_data[0]}, "
                                                              f"Study Type -> {j[0]}. Actual message is "
                                                              f"{actual_excel_upload_status_text} and Expected "
                                                              f"message is {expected_excel_upload_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding Search strategy data for "
                                    f"Population -> {pop_data[0]}, Study Type -> {j[0]}.")
                
                self.refreshpage()
                time.sleep(2)
        except Exception:
            raise Exception("Unable to add Search strategy data")

    def add_invalid_search_strategy_details(self, locatorname, filepath, env):
        expected_error_text = "Error while updating search strategy: The file extension should belong to this " \
                              "list: [.xls, .xlsx]"

        today = date.today()
        # Manipulating the date values when day point to month end
        if today.day in [30, 31]:
            day_val = (today - timedelta(10)).strftime("%d")
        else:
            day_val = today.day

        # Read population details from data sheet
        pop_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Population')
        stdy_data = self.exbase.get_four_cols_data(filepath, locatorname, 'Sheet1', 'Study_Types', 'Invalid_Files',
                                                   'db_search_val', 'Template_name')
        template_name = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Template_name')

        time.sleep(2)
        try:
            for j in stdy_data:
                # pop_ele = self.select_element("searchstrategy_pop_dropdown", env)
                # select = Select(pop_ele)
                # select.select_by_visible_text(pop_data[0])
                selected_pop_val = self.base.selectbyvisibletext("searchstrategy_pop_dropdown", pop_data[0], env)
                time.sleep(1)

                # stdy_ele = self.select_element("searchstrategy_study_type_dropdown", env)
                # select = Select(stdy_ele)
                # select.select_by_visible_text(j[0])
                selected_slr_val = self.base.selectbyvisibletext("searchstrategy_study_type_dropdown", j[0], env)
                self.LogScreenshot.fLogScreenshot(message=f"Selected Population and SLR Type Details: ",
                                                  pass_=True, log=True, screenshot=True)
                time.sleep(1)

                self.click("searchstrategy_date", env)
                self.select_calendar_date(day_val)

                self.input_text("searchstrategy_dbsearch", j[2], env)

                jscmd = ReadConfig.get_remove_att_JScommand(17, 'hidden')
                self.jsclick_hide(jscmd)
                self.input_text("searchstrategy_upload_file", os.getcwd()+"\\"+j[1], env)

                self.jsclick("searchstrategy_upload_btn", env)
                time.sleep(2)

                actual_error_text = self.get_status_text("searchstrategy_status_text", env)
                # time.sleep(2)

                if actual_error_text == expected_error_text:
                    self.LogScreenshot.fLogScreenshot(message=f"File with invalid format is not uploaded as expected. "
                                                              f"Invalid file is '{os.path.basename(j[1])}'",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding "
                                                              f"Search strategy data for Population -> {pop_data[0]}, "
                                                              f"Study Type -> {j[0]}. Actual status message is "
                                                              f"{actual_error_text} and Expected status message is "
                                                              f"{expected_error_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding Search strategy data for Population "
                                    f"-> {pop_data[0]}, Study Type -> {j[0]}.")
                
                self.refreshpage()
                time.sleep(2)
        except Exception:
            raise Exception("Unable to add Search strategy data")

    def validate_view_search_strategy(self, locatorname, filepath, env, prjname):
        expected_excel_upload_status_text = "Search strategy updated successfully"

        today = date.today()
        # Manipulating the date values when day point to month end
        if today.day in [30, 31]:
            day_val = (today - timedelta(10)).strftime("%d")
        else:
            day_val = today.day        

        # Read population details from data sheet
        pop_data = self.exbase.get_population_data(filepath, 'Sheet1', locatorname)
        stdy_data = self.exbase.get_four_cols_data(filepath, locatorname, 'Sheet1', 'Study_Types',
                                                   'slrtype_Radio_button', 'Files_to_upload', 'db_search_val')

        time.sleep(2)
        try:
            for i in pop_data:
                for j in stdy_data:
                    self.base.presence_of_admin_page_option("protocol_link", env)
                    self.base.go_to_nested_page("protocol_link", "searchstrategy", env)

                    selected_pop_val = self.base.selectbyvisibletext("searchstrategy_pop_dropdown", i[0], env)
                    time.sleep(1)

                    selected_slr_val = self.base.selectbyvisibletext("searchstrategy_study_type_dropdown", j[0], env)
                    time.sleep(1)

                    self.click("searchstrategy_date", env)
                    self.select_calendar_date(day_val)

                    self.input_text("searchstrategy_dbsearch", j[3], env)

                    jscmd = ReadConfig.get_remove_att_JScommand(17, 'hidden')
                    self.jsclick_hide(jscmd)
                    self.input_text("searchstrategy_upload_file", os.getcwd()+"\\"+j[2], env)

                    self.jsclick("searchstrategy_upload_btn", env)
                    time.sleep(2)

                    actual_excel_upload_status_text = self.get_status_text("searchstrategy_status_text", env)
                    # time.sleep(2)

                    if actual_excel_upload_status_text == expected_excel_upload_status_text:
                        self.LogScreenshot.fLogScreenshot(
                            message=f"Addition of Search strategy data is success for Population -> {i[0]}, "
                                    f"Study Type -> {j[0]}.", pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f"Unable to find status message while adding Search strategy data for "
                                    f"Population -> {i[0]}, Study Type -> {j[0]}. Actual message is "
                                    f"{actual_excel_upload_status_text} and Expected message is "
                                    f"{expected_excel_upload_status_text}", pass_=False, log=True, screenshot=True)
                        raise Exception(f"Unable to find status message while adding Search strategy data for "
                                        f"Population -> {i[0]}, Study Type -> {j[0]}.")
                    
                    uploaded_filename = self.export_web_table(
                        "table table-bordered table-striped table-sm search-terms",
                        f"uploadedsearchstrategydata_{j[0]}_{prjname}")

                    self.refreshpage()
                    time.sleep(2)

                    self.base.go_to_page("SLR_Homepage", env)
                    self.slrreport.select_data(i[0], i[1], env)

                    temp = j[0]
                    if j[0] == 'Clinical-Interventional' or j[0] == 'Clinical-RWE':
                        j[0] = 'Clinical'
                    if j[0] == 'Quality of life':
                        j[0] = 'Quality of Life'
                    self.slrreport.select_data(j[0], j[1], env)

                    self.click("view_searchstrategy_button", env)
                    time.sleep(1)

                    if temp == 'Clinical-Interventional':
                        self.click("view_search_clinical_intervention_tab", env)                   
                    if temp == 'Clinical-RWE':
                        self.click("view_search_clinical_rwe_tab", env)

                    if temp == 'Clinical-Interventional' or temp == 'Clinical-RWE':
                        j[0] = temp
                        view_date_val = self.get_text("view_search_date_active_tab_data", env)
                        database_search_val = self.get_text("view_search_database_active_tab_data", env)
                    else:
                        view_date_val = self.get_text("view_search_date", env)
                        database_search_val = self.get_text("view_search_database", env)

                    view_strategy_filename = self.export_web_table(
                        "table table-bordered table-striped table-sm search-term",
                        f"viewsearchstrategydata_{j[0]}_{prjname}")

                    searchstrategy_data = pd.read_excel(
                        f'ActualOutputs//web_table_exports//{uploaded_filename}', usecols='B:D')
                    view_searchstrategy_data = pd.read_excel(
                        f'ActualOutputs//web_table_exports//{view_strategy_filename}', usecols='B:D')

                    view_searchstrategy_data.rename(columns={'Unnamed: 0': 'Line'}, inplace=True)

                    if searchstrategy_data.equals(view_searchstrategy_data) and j[3] == database_search_val:
                        self.LogScreenshot.fLogScreenshot(
                            message=f"File contents between Uploaded Search Strategy File "
                                    f"'{Path(f'ActualOutputs//web_table_exports//{uploaded_filename}').name}' and "
                                    f"View Search Strategy Data "
                                    f"'{Path(f'ActualOutputs//web_table_exports//{view_strategy_filename}').name}' "
                                    f"are matching", pass_=True, log=True, screenshot=True)
                    else:
                        self.LogScreenshot.fLogScreenshot(
                            message=f"Strategy DB value : {j[3]} and View DB value : {database_search_val}",
                            pass_=False, log=True, screenshot=True)
                        self.LogScreenshot.fLogScreenshot(
                            message=f"File contents between Uploaded Search Strategy File "
                                    f"'{Path(f'ActualOutputs//web_table_exports//{uploaded_filename}').name}' and "
                                    f"View Search Strategy Data "
                                    f"'{Path(f'ActualOutputs//web_table_exports//{view_strategy_filename}').name}' "
                                    f"are not matching", pass_=False, log=True, screenshot=True)
                        raise Exception(f"File contents between Uploaded Search Strategy File "
                                        f"'{Path(f'ActualOutputs//web_table_exports//{uploaded_filename}').name}' "
                                        f"and View Search Strategy Data "
                                        f"'{Path(f'ActualOutputs//web_table_exports//{view_strategy_filename}').name}' "
                                        f"are not matching")
                    
                    self.refreshpage()
        except Exception:
            raise Exception("Unable to add Search strategy data")
