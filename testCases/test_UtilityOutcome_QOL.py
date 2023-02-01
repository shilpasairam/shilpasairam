from csv import excel
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
class Test_UtilityOutcome_QOL:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    util_filepath = ReadConfig.getutilityoutcome_QOL_data()

    @pytest.mark.C26937
    def test_qol_presenceof_utilitysummary_into_excelreport(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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

        self.LogScreenshot.fLogScreenshot(message=f"*****Presence of Utility Summary Sheet in Complete Excel "
                                                  f"Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(excel_filename, self.util_filepath, 'NewImportLogic')
                    
                    self.utiloutcome.qol_presenceof_utilitysummary_into_excelreport(excel_filename, self.util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C26956
    def test_qol_verify_source_to_target_row_counts_in_excelreport(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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

        self.LogScreenshot.fLogScreenshot(message=f"*****Utility Summary Sheet Row count validation between "
                                                  f"Source Utility File and Complete Excel Report*****",
                                          pass_=True, log=True, screenshot=False)

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(excel_filename, self.util_filepath, 'NewImportLogic')
                    
                    self.utiloutcome.qol_verify_source_to_target_row_counts_excelreport(excel_filename,
                                                                                        self.util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C27567
    def test_qol_presenceof_utilitysummary_into_wordreport(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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

        self.LogScreenshot.fLogScreenshot(message=f"*****Presence of Utility Summary Sheet in Word Report "
                                                  f"validation*****",
                                          pass_=True, log=True, screenshot=False)

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")

                    self.slrreport.generate_download_report("word_report")
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(word_filename, self.util_filepath, 'NewImportLogic')
                    
                    self.utiloutcome.qol_presenceof_utilitysummary_into_wordreport(word_filename, self.util_filepath)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C27569
    def test_qol_verify_source_to_target_row_counts_in_wordreport(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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

        self.LogScreenshot.fLogScreenshot(message=f"*****Utility Summary Sheet Row count validation between "
                                                  f"Source Utility File and Word Report*****",
                                          pass_=True, log=True, screenshot=False)

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")

                    self.slrreport.generate_download_report("word_report")
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(word_filename, self.util_filepath, 'NewImportLogic')

                    self.utiloutcome.qol_verify_source_to_target_row_counts_wordreport(word_filename,
                                                                                       self.util_filepath)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C26938
    def test_qol_verify_excelreport_utility_summary_sorting_order(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(excel_filename, self.util_filepath, 'NewImportLogic')

                    self.utiloutcome.qol_verify_excelreport_utility_summary_sorting_order(excel_filename,
                                                                                          self.util_filepath)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C27568
    def test_qol_verify_wordreport_utility_table_sorting_order(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    
                    self.slrreport.generate_download_report("word_report")
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(word_filename, self.util_filepath, 'NewImportLogic')

                    self.utiloutcome.qol_verify_wordreport_utility_table_sorting_order(word_filename,
                                                                                       self.util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C30281
    def test_qol_utility_outcome_results_comparison_newimport_logic(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

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

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(excel_filename, self.util_filepath, 'NewImportLogic')

                    self.slrreport.generate_download_report("word_report")
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(word_filename, self.util_filepath, 'NewImportLogic')

                    self.slrreport.preview_result("preview_results")
                    self.slrreport.table_display_check("Table")
                    self.slrreport.generate_download_report("Export_as_excel")
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    webexcel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(webexcel_filename, self.util_filepath, 'NewImportLogic')
                    self.slrreport.back_to_report_page("Back_to_search_page")

                    self.utiloutcome.qol_utility_summary_validation(webexcel_filename, excel_filename,
                                                                    self.util_filepath, word_filename)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C30280
    @pytest.mark.C31399
    def test_qol_utility_outcome_results_comparison_oldimport_logic(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'OldImportLogic')

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

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(excel_filename, self.util_filepath, 'OldImportLogic')

                    self.slrreport.generate_download_report("word_report")
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(word_filename, self.util_filepath, 'OldImportLogic')

                    self.slrreport.preview_result("preview_results")
                    self.slrreport.table_display_check("Table")
                    self.slrreport.generate_download_report("Export_as_excel")
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    webexcel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(webexcel_filename, self.util_filepath, 'OldImportLogic')
                    self.slrreport.back_to_report_page("Back_to_search_page")

                    self.utiloutcome.qol_utility_summary_validation_old_imports(webexcel_filename, excel_filename,
                                                                                self.util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C31498
    def test_qol_validate_utilitysummarytab_and_contents_into_excelreport(self, extra):
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Presence of Utility Summary Sheet in Complete Excel "
                                                  f"Report validation*****",
                                          pass_=True, log=True, screenshot=False)
        
        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for index, i in enumerate(scenarios):
            try:
                self.utiloutcome.qol_validate_utilitysummarytab_and_contents_into_excelreport(i, self.util_filepath,
                                                                                              index)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C26906
    def test_qol_presenceof_newlyadded_utilitycolumn_names(self, extra):
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
        self.rpt_data, self.rpt_data_chkbox = self.utiloutcome.\
            get_reported_variables_specific_sheet(self.util_filepath, 'NewImportLogic')

        expected_dict = {"FH-1": "Utility/Disutility Summary (Excluding point estimates)",
                         "FH-2": "Utility Point Estimate Reported with Health States",
                         "FH-3": "Disutility Point Estimate Reported with Health States",
                         "FH-4": "Utility Elicitation Method and Source"}

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

        self.LogScreenshot.fLogScreenshot(message=f"*****Column names validation started*****",
                                          pass_=True, log=True, screenshot=False)

        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for j in self.slrtype:
                    self.slrreport.select_data(j[0], j[1])
                    self.slrreport.select_sub_section(self.rpt_data[1], self.rpt_data_chkbox[1],
                                                      "reported_variable_section")
                    
                    self.slrreport.generate_download_report("excel_report")
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(excel_filename, self.util_filepath, 'NewImportLogic')

                    self.slrreport.preview_result("preview_results")
                    self.slrreport.table_display_check("Table")
                    self.utiloutcome.check_column_names_in_previewresults(expected_dict)
                    self.slrreport.generate_download_report("Export_as_excel")
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    webexcel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    self.utiloutcome.validate_filename(webexcel_filename, self.util_filepath, 'NewImportLogic')
                    self.slrreport.back_to_report_page("Back_to_search_page")

                    self.utiloutcome.presenceof_utilitycolumn_names(webexcel_filename, excel_filename, expected_dict)
                
                self.LogScreenshot.fLogScreenshot(message=f"*****Column names validation Completed*****",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                raise Exception("Unable to select element")
