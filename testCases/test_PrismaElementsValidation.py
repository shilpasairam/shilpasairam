import os
import time

import pytest

from Pages.Base import Base
from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.ExcludedStudies_liveSLR import ExcludedStudies_liveSLR
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_SLR_Custom_Report:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getslrtestdata()
    prisma_path = ReadConfig.getexcludedstudiesliveslrpath()

    @pytest.mark.C26957
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
        self.loginPage.complete_login(self.username, self.password, self.baseURL)
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

    @pytest.mark.C31633
    def test_prisma_ele_comparison_between_Excel_and_Word_Report(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Prisma Elements Comparison between Complete Excel "
                                                  f"and Word Report*****",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = self.exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = self.exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = self.exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                self.slrreport.test_prisma_ele_comparison_between_Excel_and_Word_Report(pop_data, slr_type,
                                                                                        add_criteria, self.prisma_path)
            except Exception:
                raise Exception("Unable to select element")
                
    @pytest.mark.C31634
    def test_prisma_ele_comparison_between_Excel_and_UI(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Prisma Elements Comparison between Complete Excel "
                                                  f"and UI*****",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = self.exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = self.exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = self.exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                self.slrreport.test_prisma_ele_comparison_between_Excel_and_UI(pop_data, slr_type, add_criteria,
                                                                               self.prisma_path)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C32354
    def test_prisma_count_comparison_between_prismatab_and_excludedstudiesliveslr(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Prisma Counts Comparison between 'Updated PRISMA' sheet and "
                                                  f"'Excluded studies - LiveSLR' sheet in Complete Excel Report*****",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = self.exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = self.exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = self.exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                self.slrreport.test_prisma_count_comparison_between_prismatab_and_excludedstudiesliveslr(pop_data,
                                                                                                         slr_type,
                                                                                                         add_criteria,
                                                                                                         self.
                                                                                                         prisma_path)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C31632
    def test_prisma_tab_format_changes(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"*****Updated PRISMA tab format changes in Complete Excel "
                                                  f"Report*****",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = self.exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = self.exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = self.exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                self.slrreport.test_prisma_tab_format_changes(pop_data, slr_type, add_criteria, self.prisma_path)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C32349
    def test_publication_identifier_count_in_updated_prisma_tab(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        self.exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
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

        self.LogScreenshot.fLogScreenshot(message=f"***Publications Count Comparison between 'Updated PRISMA' sheet "
                                                  f"and 'Excluded studies - LiveSLR' sheet in Complete Excel Report***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = self.exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = self.exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = self.exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                self.slrreport.test_publication_identifier_count_in_updated_prisma_tab(pop_data, slr_type,
                                                                                       add_criteria, self.prisma_path)
            except Exception:
                raise Exception("Unable to select element")
