import os
import time

import pytest

from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.ImportPublicationsPage import ImportPublicationPage
from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.ExcludedStudies_liveSLR import ExcludedStudies_liveSLR
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_PRISMA_Elements:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getslrtestdata()
    prisma_path = ReadConfig.getexcludedstudiesliveslrpath()

    @pytest.mark.C26957
    def test_prisma_elements(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getslrtestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Read population data values
        pop_list = liveslrpage.get_population_data(filepath)
        # Read slrtype data values
        slrtype = liveslrpage.get_slrtype_data(filepath)
        # Read reportedvariables locator details
        rpt_data, rpt_data_chkbox = liveslrpage.get_reported_variables(filepath)
        # Read StudyDesign locator details
        study_data, study_data_chkbox = liveslrpage.get_study_design(filepath)
        # Read reportedvariables and studydesign expected data values
        design_val, var_val = liveslrpage.get_data_values(filepath)

        request.node._tcid = caseid
        request.node._title = "Validate PRISMA Count between UI, WebExcel, Complete Excel and Word Report in Search " \
                              "LIVESLR page "

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for index, j in enumerate(slrtype):
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.validate_selected_area(i[0], j[0], env)
                    
                    slrreport.select_sub_section(study_data[1], study_data_chkbox[1], env, "study_design_section")
                    slrreport.select_sub_section(study_data[3], study_data_chkbox[3], env, "study_design_section")
                    slrreport.select_sub_section(rpt_data[0], rpt_data_chkbox[0], env, "reported_variable_section")
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    slrreport.validate_additional_criteria_val(filepath, "study_design_value",
                                                               "reported_variable_value", env)

                    base.scroll("New_total_selected_Onco", env)
                    prism = base.get_text("New_total_selected_Onco", env)
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # excel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    excel_filename = slrreport.get_and_validate_filename(filepath)

                    slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # word_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    word_filename = slrreport.get_and_validate_filename(filepath)

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # webexcel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    webexcel_filename = slrreport.get_and_validate_filename(filepath)
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    slrreport.prism_value_validation(prism, excel_filename, webexcel_filename, word_filename, env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C31633
    def test_prisma_ele_comparison_between_Excel_and_Word_Report(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate PRISMA elements comparison between Complete Excel and Word Report"

        LogScreenshot.fLogScreenshot(message=f"*****Prisma Elements Comparison between Complete Excel and Word "
                                             f"Report*****", pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                slrreport.test_prisma_ele_comparison_between_Excel_and_Word_Report(pop_data, slr_type,
                                                                                   add_criteria, self.prisma_path,
                                                                                   env)
            except Exception:
                raise Exception("Unable to select element")
                
    @pytest.mark.C31634
    def test_prisma_ele_comparison_between_Excel_and_UI(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate PRISMA elements comparison between Complete Excel and UI"

        LogScreenshot.fLogScreenshot(message=f"*****Prisma Elements Comparison between Complete Excel and UI*****",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                slrreport.prisma_ele_comparison_between_Excel_and_UI(i, pop_data, slr_type, add_criteria,
                                                                     'Updated PRISMA', self.prisma_path, env,
                                                                     'Oncology')
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C32354
    def test_prisma_count_comparison_between_prismatab_and_excludedstudiesliveslr(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate PRISMA count between PRISMA Tab and ExcludedStudies_LiveSLR Tab in Complete " \
                              "Excel Report"

        LogScreenshot.fLogScreenshot(message=f"*****Prisma Counts Comparison between 'Updated PRISMA' sheet and "
                                             f"'Excluded studies - LiveSLR' sheet in Complete Excel Report*****",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                slrreport.test_prisma_count_comparison_between_prismatab_and_excludedstudiesliveslr(pop_data,
                                                                                                    slr_type,
                                                                                                    add_criteria,
                                                                                                    self.prisma_path,
                                                                                                    env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C31632
    def test_prisma_tab_format_changes(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate PRISMA Tab format changes in Complete Excel Report"

        LogScreenshot.fLogScreenshot(message=f"*****Updated PRISMA tab format changes in Complete Excel Report*****",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                slrreport.test_prisma_tab_format_changes(pop_data, slr_type, add_criteria, self.prisma_path, env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C32349
    def test_publication_identifier_count_in_updated_prisma_tab(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Compare Publications Count between 'Updated PRISMA' sheet and 'Excluded studies - " \
                              "LiveSLR' sheet in Complete Excel Report "

        LogScreenshot.fLogScreenshot(message=f"***Publications Count Comparison between 'Updated PRISMA' sheet "
                                             f"and 'Excluded studies - LiveSLR' sheet in Complete Excel Report***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for i in scenarios:
            try:
                pop_data = exstdy_liveslr.get_population_data(self.prisma_path, i)
                slr_type = exstdy_liveslr.get_slrtype_data(self.prisma_path, i)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(self.prisma_path, i)

                slrreport.test_publication_identifier_count_in_updated_prisma_tab(pop_data, slr_type,
                                                                                  add_criteria, self.prisma_path, env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C39058
    def test_nononcology_prisma_ele_comparison_between_Excel_and_UI(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Validate PRISMA elements comparison between Complete Excel, UI and Expected Count"

        LogScreenshot.fLogScreenshot(message=f"*****Prisma Elements Comparison between Complete Excel, UI and Expected Count for Non-Oncology Population*****",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        filepath = exbase.get_testdata_filepath(basefile, "livehta_2301_data")

        scenarios = ['scenario1', 'scenario2']

        for i in scenarios:
            try:
                pop_data = exstdy_liveslr.get_population_data(filepath, i)
                slr_type = exstdy_liveslr.get_slrtype_data(filepath, i)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(filepath, i)

                slrreport.prisma_ele_comparison_between_Excel_and_UI(i, pop_data, slr_type, add_criteria,
                                                                     'Updated PRISMAs', filepath, env, 'Non-Oncology')
            except Exception:
                raise Exception("Unable to select element")
