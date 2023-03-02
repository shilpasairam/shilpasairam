from datetime import date, timedelta
import os
import time
import pandas as pd
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

    # def go_to_prisma(self, locator, button):
    #     self.click(locator, UnivWaitFor=10)
    #     self.jsclick(button)
    #     time.sleep(5)

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
            pop_ele = self.select_element("prisma_pop_dropdown", env)
            select = Select(pop_ele)
            select.select_by_visible_text(pop_name[0])

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
                                                          f"PRISMA Excel File. Actual status message is {actual_excel_upload_status_text} and Expected status message is {expected_excel_upload_status_text}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message while uploading PRISMA Excel File.")
        except Exception:
            raise Exception("Unable to upload PRISMA Excel file")
    
    def del_prisma_excel_file(self, locatorname, excel_del_locator, excel_del_popup, filepath, env):
        expected_excel_del_status_text = "PRISMA excel file successfully deleted"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)
        
        # self.refreshpage()
        # time.sleep(3)
        try:
            pop_ele = self.select_element("prisma_pop_dropdown", env)
            select = Select(pop_ele)
            select.select_by_visible_text(pop_name[0])

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
                                                          f"PRISMA Excel File. Actual status message is {actual_excel_del_status_text} and Expected status message is {expected_excel_del_status_text}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message while deleting PRISMA Excel File.")
            time.sleep(5)
        except Exception:
            raise Exception("Unable to delete PRISMA Excel file")

    def upload_prisma_image(self, locatorname, filepath, index, env):
        expected_image_upload_status_text = "PRISMA successfully updated"

        # Read the Data required to upload for study types
        stdy_data = self.get_prisma_data_details(filepath, locatorname)

        try:
            for i in stdy_data:
                stdy_ele = self.select_element("prisma_study_type_dropdown", env)
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(2)

                self.input_text("original_studies", i[1]+index, env, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("records_number", i[2]+index, env, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("fulltext_review", i[3]+index, env, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("total_record_number", i[4]+index, env, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("prisma_image", i[5], env, UnivWaitFor=5)
                time.sleep(1)

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
                                                              f"PRISMA Image File. Actual status message is {actual_image_upload_status_text} and Expected status message is {expected_image_upload_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message while uploading PRISMA Image File.")
        except Exception:
            raise Exception("Unable to upload PRISMA image")
    
    def del_prisma_image(self, locatorname, filepath, del_locator, del_popup, env):
        expected_image_del_status_text = "PRISMA successfully deleted"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)

        stdy_data = self.get_prisma_data_details(filepath, locatorname)

        try:
            for i in stdy_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("prisma_pop_dropdown", env)
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)

                stdy_ele = self.select_element("prisma_study_type_dropdown", env)
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
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
                                                              f"PRISMA Image File. Actual status message is {actual_image_del_status_text} and Expected status message is {expected_image_del_status_text}",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message while deleting PRISMA Image File.")
                time.sleep(5)
        except Exception:
            raise Exception("Unable to delete PRISMA image")

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
                pop_ele = self.select_element("picos_pop_dropdown", env)
                select = Select(pop_ele)
                select.select_by_visible_text(pop_data[0])

                stdy_ele = self.select_element("picos_study_type_dropdown", env)
                select = Select(stdy_ele)
                select.select_by_visible_text(j)

                actual_row_headers = []
                actual_row_eles = self.select_elements('row_header_values', env)
                for k in actual_row_eles:
                    actual_row_headers.append(k.text)

                actual_col_headers = []
                actual_col_eles = self.select_elements('col_header_values', env)
                for v in actual_col_eles:
                    actual_col_headers.append(v.text)
                # Removing the empty column name
                actual_col_headers.pop(0)

                row_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_row_headers, actual_row_headers)

                if len(row_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page row headers are displayed as expected for Population -> {pop_data[0]}, Study Type -> {j}.",
                                                        pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page row headers for Population -> {pop_data[0]}, Study Type -> {j}. Mismatch values are arranged in following order -> Expected row headers, Actual row headers. {row_header_comparison}",
                                                        pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {pop_data[0]}, Study Type -> {j}.")

                col_header_comparison = self.slrreport.list_comparison_between_reports_data(expected_col_headers, actual_col_headers)

                if len(col_header_comparison) == 0:
                    self.LogScreenshot.fLogScreenshot(message=f"PICOS page col headers are displayed as expected for Population -> {pop_data[0]}, Study Type -> {j}.",
                                                        pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in PICOS page col headers for Population -> {pop_data[0]}, Study Type -> {j}. Mismatch values are arranged in following order -> Expected col headers, Actual col headers. {col_header_comparison}",
                                                        pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch found in PICOS page row headers for Population -> {pop_data[0]}, Study Type -> {j}.")

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
                    self.LogScreenshot.fLogScreenshot(message=f"Addition of PICOS data is success for Population -> {pop_data[0]}, Study Type -> {j}.",
                                                    pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding "
                                                            f"PICOS data for Population -> {pop_data[0]}, Study Type -> {j}. Actual status message is {actual_status_text} and Expected status message is {expected_status_text}",
                                                    pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding PICOS data for Population -> {pop_data[0]}, Study Type -> {j}.")
                
                self.refreshpage()
                time.sleep(2)
        except Exception:
            raise Exception("Unable to add PICOS data")

    def add_valid_search_strategy_details(self, locatorname, filepath, env):
        expected_excel_upload_status_text = "Search strategy updated successfully"

        today = date.today()
        # Manipulating the date values when day point to month end
        if today.day in [30, 31]:
            day_val = (today - timedelta(10)).strftime("%d")
        else:
            day_val = today.day

        # Read population details from data sheet
        pop_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Population')
        stdy_data = self.exbase.get_four_cols_data(filepath, locatorname, 'Sheet1', 'Study_Types', 'Files_to_upload', 'db_search_val', 'Template_name')
        template_name = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Template_name')

        time.sleep(2)
        try:
            for j in stdy_data:
                pop_ele = self.select_element("searchstrategy_pop_dropdown", env)
                select = Select(pop_ele)
                select.select_by_visible_text(pop_data[0])

                stdy_ele = self.select_element("searchstrategy_study_type_dropdown", env)
                select = Select(stdy_ele)
                select.select_by_visible_text(j[0])

                self.slrreport.generate_download_report("searchstrategy_template_download_btn", env)
                # Renaming the filename because there is an issue in downloading filenames with same name multiple times in headless mode.
                # So as an alternative renaming the file after downloading in each iteration
                self.file_rename(self.slrreport.get_latest_filename(UnivWaitFor=180), f"{j[0]}_search-strategy-template.xlsx")
                downloaded_template_name = self.slrreport.get_latest_filename(UnivWaitFor=180)                
                if downloaded_template_name == j[3]:
                    self.LogScreenshot.fLogScreenshot(message=f"Correct Template is downloaded. Template name is {downloaded_template_name}",
                                                        pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Mismatch in search strategy template name. Expected Template name is {template_name[0]} and Actual Template name is {downloaded_template_name}",
                                                        pass_=False, log=True, screenshot=False)
                    raise Exception(f"Mismatch in search strategy template name.")

                self.click("searchstrategy_date", env)
                self.select_calendar_date(day_val)

                self.input_text("searchstrategy_dbsearch", j[2], env)

                jscmd = ReadConfig.get_remove_att_JScommand(17, 'hidden')
                self.jsclick_hide(jscmd)
                self.input_text("searchstrategy_upload_file", os.getcwd()+"\\"+j[1], env) # f"{os.getcwd()} + {j[1]}"

                self.jsclick("searchstrategy_upload_btn", env)
                time.sleep(2)

                actual_excel_upload_status_text = self.get_status_text("searchstrategy_status_text", env)
                # time.sleep(2)

                if actual_excel_upload_status_text == expected_excel_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"Addition of Search strategy data is success for Population -> {pop_data[0]}, Study Type -> {j[0]}.",
                                                    pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding "
                                                            f"Search strategy data for Population -> {pop_data[0]}, Study Type -> {j[0]}. Actual message is {actual_excel_upload_status_text} and Expected message is {expected_excel_upload_status_text}",
                                                    pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding Search strategy data for Population -> {pop_data[0]}, Study Type -> {j[0]}.")
                
                self.refreshpage()
                time.sleep(2)
        except Exception:
            raise Exception("Unable to add Search strategy data")

    def add_invalid_search_strategy_details(self, locatorname, filepath, env):
        expected_error_text = "Error while updating search strategy: The file extension should belong to this list: [.xls, .xlsx]"

        today = date.today()
        # Manipulating the date values when day point to month end
        if today.day in [30, 31]:
            day_val = (today - timedelta(10)).strftime("%d")
        else:
            day_val = today.day

        # Read population details from data sheet
        pop_data = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Population')
        stdy_data = self.exbase.get_four_cols_data(filepath, locatorname, 'Sheet1', 'Study_Types', 'Invalid_Files', 'db_search_val', 'Template_name')
        template_name = self.exbase.get_individual_col_data(filepath, locatorname, 'Sheet1', 'Template_name')

        time.sleep(2)
        try:
            for j in stdy_data:
                pop_ele = self.select_element("searchstrategy_pop_dropdown", env)
                select = Select(pop_ele)
                select.select_by_visible_text(pop_data[0])

                stdy_ele = self.select_element("searchstrategy_study_type_dropdown", env)
                select = Select(stdy_ele)
                select.select_by_visible_text(j[0])
                self.LogScreenshot.fLogScreenshot(message=f"Selected Population and SLR Type Details: ",
                                                    pass_=True, log=True, screenshot=True)

                self.click("searchstrategy_date", env)
                self.select_calendar_date(day_val)

                self.input_text("searchstrategy_dbsearch", j[2], env)

                jscmd = ReadConfig.get_remove_att_JScommand(17, 'hidden')
                self.jsclick_hide(jscmd)
                self.input_text("searchstrategy_upload_file", os.getcwd()+"\\"+j[1], env)

                self.jsclick("searchstrategy_upload_btn", env)
                # time.sleep(2)

                actual_error_text = self.get_status_text("searchstrategy_status_text", env)
                # time.sleep(2)

                if actual_error_text == expected_error_text:
                    self.LogScreenshot.fLogScreenshot(message=f"File with invalid format is not uploaded as expected. "
                                                              f"Invalid file is '{os.path.basename(j[1])}'",
                                                      pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while adding "
                                                            f"Search strategy data for Population -> {pop_data[0]}, Study Type -> {j[0]}. Actual status message is {actual_error_text} and Expected status message is {expected_error_text}",
                                                    pass_=False, log=True, screenshot=True)
                    raise Exception(f"Unable to find status message while adding Search strategy data for Population -> {pop_data[0]}, Study Type -> {j[0]}.")
                
                self.refreshpage()
                time.sleep(2)
        except Exception:
            raise Exception("Unable to add Search strategy data")
