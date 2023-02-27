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
class Test_UtilityOutcome_ECON:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # util_filepath = ReadConfig.getutilityoutcome_ECON_data()

    @pytest.mark.C29953
    def test_econ_utility_outcome_results_comparison_newimport_logic(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_ECON_data(env)
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
        request.node._title = "ECON -> Utility Outcome comparison with New Import Logic"

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

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[0], rpt_data_chkbox[0], env,
                                                      "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # excel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'NewImportLogic')                    

                    slrreport.generate_download_report("word_report", env)
                    # time.sleep(5)
                    # word_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # word_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    word_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(word_filename, util_filepath, 'NewImportLogic')                    

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # webexcel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    webexcel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(webexcel_filename, util_filepath, 'NewImportLogic')                    
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    utiloutcome.econ_utility_summary_validation(webexcel_filename, excel_filename,
                                                                     util_filepath, word_filename)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C29986
    def test_econ_utility_outcome_results_comparison_oldimport_logic(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_ECON_data(env)        
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
        request.node._title = "ECON -> Utility Outcome comparison with Old Import Logic"

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

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        for i in pop_list:
            try:
                slrreport.select_data(i[0], i[1], env)
                for j in slrtype:
                    slrreport.select_data(j[0], j[1], env)
                    slrreport.select_sub_section(rpt_data[0], rpt_data_chkbox[0], env,
                                                      "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # excel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'OldImportLogic')

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # webexcel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    webexcel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(webexcel_filename, util_filepath, 'OldImportLogic')
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    utiloutcome.econ_utility_summary_validation_old_imports(webexcel_filename, excel_filename,
                                                                                 util_filepath)
            except Exception:
                raise Exception("Unable to select element")

    @pytest.mark.C26907
    def test_econ_presenceof_newlyadded_utilitycolumn_names(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        util_filepath = ReadConfig.getutilityoutcome_ECON_data(env)        
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
        request.node._title = "ECON -> Validate presence of Newly added Utility Column Names"

        expected_dict = {"FK-14": "Other Utility Data (Excluding point estimates)",
                         "FK-127": "Utility point estimates reported with health states",
                         "FK-128": "Disutility point estimates reported with health states",
                         "FK-129": "Utility Elicitation Method and Source"}

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
                    slrreport.select_sub_section(rpt_data[0], rpt_data_chkbox[0], env,
                                                      "reported_variable_section")
                    
                    slrreport.generate_download_report("excel_report", env)
                    # time.sleep(5)
                    # excel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # excel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    excel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(excel_filename, util_filepath, 'NewImportLogic')

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    utiloutcome.check_column_names_in_previewresults(expected_dict, env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    # time.sleep(5)
                    # webexcel_filename1 = self.slrreport.getFilenameAndValidate(180)
                    # webexcel_filename1 = self.slrreport.get_latest_filename(UnivWaitFor=180)
                    webexcel_filename = slrreport.get_latest_filename(UnivWaitFor=180)
                    utiloutcome.validate_filename(webexcel_filename, util_filepath, 'NewImportLogic')
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    utiloutcome.presenceof_utilitycolumn_names(webexcel_filename, excel_filename, expected_dict)
                
                LogScreenshot.fLogScreenshot(message=f"*****Column names validation Completed*****",
                                                  pass_=True, log=True, screenshot=False)
            except Exception:
                raise Exception("Unable to select element")
