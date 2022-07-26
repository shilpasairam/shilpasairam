import os
import time

import pytest

from Pages.Base import Base
from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_SLR_Custom_Report:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getslrtestdata()

    def test_prisma_elements(self, extra):
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        self.pop_list = self.liveslrpage.get_population_data(self.filepath)
        # Read slrtype data values
        self.slrtype = self.liveslrpage.get_slrtype_data(self.filepath)
        # Read reportedvariables locator details
        self.rpt_data, self.rpt_data_chkbox = self.liveslrpage.get_reported_variables(self.filepath)
        # Read StudyDesign locator details
        self.study_data, self.study_data_chkbox = self.liveslrpage.get_study_design(self.filepath)
        # Read reportedvariables and studydesign expected data values
        self.design_val, self.var_val = self.liveslrpage.get_data_values(self.filepath)

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for index, j in enumerate(self.slrtype):
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.study_data[1], self.study_data_chkbox[1],
                                                      "study_design_section")
                    self.slrreport.select_sub_section(self.study_data[3], self.study_data_chkbox[3],
                                                      "study_design_section")
                    self.slrreport.select_sub_section(self.rpt_data[0], self.rpt_data_chkbox[0],
                                                      "reported_variable_section")
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    self.slrreport.validate_additional_criteria_val(self.filepath, "study_design_value",
                                                                    "reported_variable_value")

                    self.base.scroll("New_total_selected")
                    prism = self.base.get_text("New_total_selected")
                    
                    self.slrreport.generate_download_report("excel_report")
                    time.sleep(5)
                    excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(excel_filename1, self.filepath)

                    self.slrreport.generate_download_report("word_report")
                    time.sleep(5)
                    word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(word_filename1, self.filepath)

                    self.slrreport.preview_result("preview_results")
                    self.slrreport.table_display_check("Table")
                    self.slrreport.generate_download_report("Export_as_excel")
                    time.sleep(5)
                    webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(webexcel_filename1, self.filepath)
                    self.slrreport.back_to_report_page("Back_to_search_page")

                    self.slrreport.prism_value_validation(prism, excel_filename1, webexcel_filename1, word_filename1)
                    self.slrreport.validate_selected_area(i[0], j[0])
            except Exception:
                raise Exception("Unable to select element")