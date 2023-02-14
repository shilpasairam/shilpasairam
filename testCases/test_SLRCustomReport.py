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
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getslrtestdata()
    testdata_723 = ReadConfig.getTestdata("livehta_723_data")
    testdata_931 = ReadConfig.getTestdata("livehta_931_data")

    @pytest.mark.C26790
    @pytest.mark.C26859
    @pytest.mark.C26860
    def test_reports_comparison(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getslrtestdata(env)        
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
        self.pop_list = self.liveslrpage.get_population_data(filepath)
        # Read slrtype data values
        self.slrtype = self.liveslrpage.get_slrtype_data(filepath)
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.liveslrpage.get_reported_variables(filepath)

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

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("SLR_Homepage", env)
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1], env)
                for index, j in enumerate(self.slrtype):
                    self.slrreport.select_data(j[0], j[1], env)
                    if j[0] == "Clinical":
                        self.slrreport.select_sub_section(self.rpt_data[3], self.rpt_data_chkbox[3], env,
                                                          "reported_variable_section")

                    self.slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # excel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    excel_filename = self.slrreport.get_and_validate_filename(filepath)

                    self.slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # word_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    word_filename = self.slrreport.get_and_validate_filename(filepath)

                    self.slrreport.preview_result("preview_results", env)
                    self.slrreport.table_display_check("Table", env)
                    self.slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # webexcel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    webexcel_filename = self.slrreport.get_and_validate_filename(filepath)
                    self.slrreport.back_to_report_page("Back_to_search_page", env)

                    self.slrreport.check_sorting_order_in_excel_report(webexcel_filename, excel_filename)

                    self.slrreport.excel_content_validation(filepath, index, webexcel_filename, excel_filename)

                    # self.slrreport.word_content_validation(filepath, index, word_filename)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C28728
    def test_presencof_publicationtype_col_in_wordreport(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getslrtestdata(env)        
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
        self.pop_list = self.liveslrpage.get_population_data(filepath)
        # Read slrtype data values
        self.slrtype = self.liveslrpage.get_slrtype_data(filepath)
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.liveslrpage.get_reported_variables(filepath)

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

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("SLR_Homepage", env)
        try:
            self.slrreport.select_data(self.pop_list[0][0], self.pop_list[0][1], env)
            self.slrreport.select_data(self.slrtype[0][0], self.slrtype[0][1], env)
            self.slrreport.select_sub_section(self.rpt_data[3], self.rpt_data_chkbox[3], env,
                                              "reported_variable_section")

            self.slrreport.generate_download_report("excel_report", env)
            # time.sleep(5)
            # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
            # excel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
            excel_filename = self.slrreport.get_and_validate_filename(filepath)

            self.slrreport.generate_download_report("word_report", env)
            # time.sleep(5)
            # word_filename1 = self.slrreport.getFilenameAndValidate(180)
            # word_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
            word_filename = self.slrreport.get_and_validate_filename(filepath)

            self.slrreport.preview_result("preview_results", env)
            self.slrreport.table_display_check("Table", env)
            self.slrreport.generate_download_report("Export_as_excel", env)
            # time.sleep(5)
            # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
            # webexcel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
            webexcel_filename = self.slrreport.get_and_validate_filename(filepath)
            self.slrreport.back_to_report_page("Back_to_search_page", env)

            self.slrreport.presencof_publicationtype_col_in_wordreport(webexcel_filename, excel_filename,
                                                                       word_filename)
        except Exception:
            raise Exception("Unable to select element")

    @pytest.mark.C31466
    def test_interventional_to_clinical_changes(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getslrtestdata(env)        
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
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

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        try:
            self.slrreport.test_interventional_to_clinical_changes(filepath, env)
        except Exception:
            raise Exception("Unable to select element")

    @pytest.mark.C37419
    def test_validate_population_col_in_wordreport(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
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

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        scenarios = ['scenario1', 'scenario2']
        for i in scenarios:
            try:
                self.slrreport.validate_population_col_in_wordreport(self.testdata_723, i, env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C31565
    def test_validate_control_chars_in_wordreport(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
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

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        scenarios = ['scenario1']
        for i in scenarios:
            try:
                self.slrreport.validate_control_chars_in_wordreport(self.testdata_931, i, env)
            except Exception:
                raise Exception("Unable to select element")
