import os
import time

import pytest

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

    @pytest.mark.C26790
    @pytest.mark.C26859
    @pytest.mark.C26860
    def test_reports_comparison(self, extra):
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
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.liveslrpage.get_reported_variables(self.filepath)

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
        self.loginPage.complete_login(self.username, self.password, self.baseURL)
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        for i in self.pop_list:
            try:
                self.slrreport.select_data(i[0], i[1])
                for index, j in enumerate(self.slrtype):
                    self.slrreport.select_data(j[0], j[1])
                    if j[0] == "Clinical":
                        self.slrreport.select_sub_section(self.rpt_data[3], self.rpt_data_chkbox[3],
                                                          "reported_variable_section")

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

                    self.slrreport.check_sorting_order_in_excel_report(webexcel_filename1, excel_filename1)

                    self.slrreport.excel_content_validation(self.filepath, index, webexcel_filename1, excel_filename1)

                    self.slrreport.excel_to_word_content_validation(webexcel_filename1, excel_filename1, word_filename1)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C28728
    def test_presencof_publicationtype_col_in_wordreport(self, extra):
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
        # Read reportedvariables data values
        self.rpt_data, self.rpt_data_chkbox = self.liveslrpage.get_reported_variables(self.filepath)

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
        self.loginPage.complete_login(self.username, self.password, self.baseURL)
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        try:
            self.slrreport.select_data(self.pop_list[0][0], self.pop_list[0][1])
            self.slrreport.select_data(self.slrtype[0][0], self.slrtype[0][1])
            self.slrreport.select_sub_section(self.rpt_data[3], self.rpt_data_chkbox[3], "reported_variable_section")

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

            self.slrreport.presencof_publicationtype_col_in_wordreport(webexcel_filename1, excel_filename1,
                                                                       word_filename1)
        except Exception:
            raise Exception("Unable to select element")
