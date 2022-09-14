import os
import time

import pytest

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from Pages.UtilityOutcome import UtilityOutcome
from utilities.readProperties import ReadConfig
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_UtilityOutcome_ECON:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    util_filepath = ReadConfig.getutilityoutcome_ECON_data()

    def test_econ_utility_outcome_results_comparison(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        self.utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        self.pop_list = self.utiloutcome.get_population_data_specific_sheet(self.util_filepath, 'NewImportLogic')
        # Read slrtype data values
        self.slrtype = self.utiloutcome.get_slrtype_data_specific_sheet(self.util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[0], self.rpt_data_chkbox[0],
                                                          "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    time.sleep(5)
                    excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(excel_filename1, self.util_filepath)

                    self.slrreport.generate_download_report("word_report")
                    time.sleep(5)
                    word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(word_filename1, self.util_filepath)

                    self.slrreport.preview_result("preview_results")
                    self.slrreport.table_display_check("Table")
                    self.slrreport.generate_download_report("Export_as_excel")
                    time.sleep(5)
                    webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(webexcel_filename1, self.util_filepath)
                    self.slrreport.back_to_report_page("Back_to_search_page")

                    self.slrreport.excel_content_validation(webexcel_filename1, excel_filename1)

                    self.slrreport.excel_to_word_content_validation(webexcel_filename1, excel_filename1, word_filename1)

                    self.utiloutcome.econ_utility_summary_validation(webexcel_filename1, excel_filename1, self.util_filepath, word_filename1)

            except Exception:
                raise Exception("Unable to select element")

    def test_econ_utility_outcome_results_comparison_oldimport_logic(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        self.utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        self.pop_list = self.utiloutcome.get_population_data_specific_sheet(self.util_filepath, 'OldImportLogic')
        # Read slrtype data values
        self.slrtype = self.utiloutcome.get_slrtype_data_specific_sheet(self.util_filepath, 'OldImportLogic')
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.get_reported_variables_specific_sheet(self.util_filepath, 'OldImportLogic')

        # # Clearing the logs before test runs
        # open(".\\Logs\\testlog.log", "w").close()
        #
        # # Removing the screenshots before the test runs
        # if os.path.exists(f'Reports/screenshots'):
        #     for root, dirs, files in os.walk(f'Reports/screenshots'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        # # Removing the files before the test runs
        # if os.path.exists(f'ActualOutputs'):
        #     for root, dirs, files in os.walk(f'ActualOutputs'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[0], self.rpt_data_chkbox[0],
                                                        "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    time.sleep(5)
                    excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(excel_filename1, self.util_filepath)

                    self.slrreport.generate_download_report("word_report")
                    time.sleep(5)
                    word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(word_filename1, self.util_filepath)

                    self.slrreport.preview_result("preview_results")
                    self.slrreport.table_display_check("Table")
                    self.slrreport.generate_download_report("Export_as_excel")
                    time.sleep(5)
                    webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    self.slrreport.validate_filename(webexcel_filename1, self.util_filepath)
                    self.slrreport.back_to_report_page("Back_to_search_page")

                    self.slrreport.excel_content_validation(webexcel_filename1, excel_filename1)

                    self.slrreport.excel_to_word_content_validation(webexcel_filename1, excel_filename1, word_filename1)

                    self.utiloutcome.econ_utility_summary_validation_old_imports(webexcel_filename1, excel_filename1, self.util_filepath, word_filename1)

            except Exception:
                raise Exception("Unable to select element")