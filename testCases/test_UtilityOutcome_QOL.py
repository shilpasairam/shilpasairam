from csv import excel
import os
import time

import pytest
from Pages.Base import Base

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
    # util_filepath = ReadConfig.getutilityoutcome_QOL_data()

    @pytest.mark.C26937
    @pytest.mark.C31110
    def test_qol_presenceof_utilitysummary_into_excelreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env) 
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Presence of Utility Summary Sheet in Complete Excel Report"

        LogScreenshot.fLogScreenshot(message=f"*****Presence of Utility Summary Sheet in Complete Excel Report "
                                             f"validation*****", pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'NewImportLogic')
                    
                    utiloutcome.qol_presenceof_utilitysummary_into_excelreport(excel_filename, util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C26956
    @pytest.mark.C31110
    def test_qol_verify_source_to_target_row_counts_in_excelreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Utility Summary Sheet Row count validation in Complete Excel Report"

        LogScreenshot.fLogScreenshot(message=f"*****Utility Summary Sheet Row count validation between Source "
                                             f"Utility File and Complete Excel Report*****",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'NewImportLogic')
                    
                    utiloutcome.qol_verify_source_to_target_row_counts_excelreport(excel_filename, util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C27567
    @pytest.mark.C31114
    def test_qol_presenceof_utilitysummary_into_wordreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Presence of Utility Summary Sheet in Word Report"

        LogScreenshot.fLogScreenshot(message=f"*****Presence of Utility Summary Sheet in Word Report validation*****",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")

                    slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(word_filename, util_filepath, 'NewImportLogic')
                    
                    utiloutcome.qol_presenceof_utilitysummary_into_wordreport(word_filename, util_filepath)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C27569
    @pytest.mark.C31114
    def test_qol_verify_source_to_target_row_counts_in_wordreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)         
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Utility Summary Sheet Row count validation in Complete Word Report"

        LogScreenshot.fLogScreenshot(message=f"*****Utility Summary Sheet Row count validation between Source "
                                             f"Utility File and Word Report*****",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")

                    slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(word_filename, util_filepath, 'NewImportLogic')

                    utiloutcome.qol_verify_source_to_target_row_counts_wordreport(word_filename, util_filepath)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C26938
    @pytest.mark.C31110
    def test_qol_verify_excelreport_utility_summary_sorting_order(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Validate sorting order of Utility Summary sheet in Complete Excel Report"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'NewImportLogic')

                    utiloutcome.qol_verify_excelreport_utility_summary_sorting_order(excel_filename, util_filepath)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C27568
    @pytest.mark.C31114
    def test_qol_verify_wordreport_utility_table_sorting_order(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Validate sorting order of Utility Summary sheet in Complete Word Report"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    
                    slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(word_filename, util_filepath, 'NewImportLogic')

                    utiloutcome.qol_verify_wordreport_utility_table_sorting_order(word_filename, util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C30281
    def test_qol_utility_outcome_results_comparison_newimport_logic(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Utility Outcome comparison with New Import Logic"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'NewImportLogic')

                    slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(word_filename, util_filepath, 'NewImportLogic')

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    webexcel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(webexcel_filename, util_filepath, 'NewImportLogic')
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    utiloutcome.qol_utility_summary_validation(webexcel_filename, excel_filename,
                                                               util_filepath, word_filename)

            except Exception:
                raise Exception("Unable to select element")
    
    @pytest.mark.C30280
    @pytest.mark.C31399
    def test_qol_utility_outcome_results_comparison_oldimport_logic(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'OldImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'OldImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'OldImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Utility Outcome comparison with Old Import Logic"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'OldImportLogic')

                    slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    word_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(word_filename, util_filepath, 'OldImportLogic')

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    webexcel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(webexcel_filename, util_filepath, 'OldImportLogic')
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    utiloutcome.qol_utility_summary_validation_old_imports(webexcel_filename, excel_filename,
                                                                           util_filepath)

            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C31498
    def test_qol_validate_utilitysummarytab_and_contents_into_excelreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "QOL -> Workflow of Utility Outcome with different data"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        scenarios = ['scenario1', 'scenario2', 'scenario3', 'scenario4']

        for index, i in enumerate(scenarios):
            try:
                utiloutcome.qol_validate_utilitysummarytab_and_contents_into_excelreport(i, util_filepath, index, env)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C26906
    @pytest.mark.C31109
    def test_qol_presenceof_newlyadded_utilitycolumn_names(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_QOL_data(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)         
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of utilityoutcome class
        utiloutcome = UtilityOutcome(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Read population data values
        pop_list = utiloutcome.get_population_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read slrtype data values
        slrtype = utiloutcome.get_slrtype_data_specific_sheet(util_filepath, 'NewImportLogic')
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = utiloutcome.get_reported_variables_specific_sheet(util_filepath, 'NewImportLogic')

        request.node._tcid = caseid
        request.node._title = "QOL -> Validate presence of newly added Utility Column names"
        
        expected_dict = {"FH-1": "Utility/Disutility Summary (Excluding point estimates)",
                         "FH-2": "Utility Point Estimate Reported with Health States",
                         "FH-3": "Disutility Point Estimate Reported with Health States",
                         "FH-4": "Utility Elicitation Method and Source"}

        LogScreenshot.fLogScreenshot(message=f"*****Column names validation started*****",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[1], rpt_data_chkbox[1], env, "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'NewImportLogic')

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    utiloutcome.check_column_names_in_previewresults(expected_dict, env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    webexcel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(webexcel_filename, util_filepath, 'NewImportLogic')
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    utiloutcome.presenceof_utilitycolumn_names(webexcel_filename, excel_filename, expected_dict)
                
                LogScreenshot.fLogScreenshot(message=f"*****Column names validation Completed*****",
                                             pass_=True, log=True, screenshot=False)
            except Exception:
                raise Exception("Unable to select element")
