import os
import time

import pytest
from Pages.Base import Base
from Pages.ExcludedStudies_liveSLR import ExcludedStudies_liveSLR
from Pages.ExtendedBasePage import ExtendedBase
from Pages.ImportPublicationsPage import ImportPublicationPage

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig
from utilities.logScreenshot import cLogScreenshot


@pytest.mark.usefixtures("init_driver")
class Test_Search_LiveSLR:
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    testdata_723 = ReadConfig.getTestdata("livehta_723_data")
    testdata_931 = ReadConfig.getTestdata("livehta_931_data")

    @pytest.mark.C40493
    def test_Smoketest_SLRreports_Download(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getsmoketestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
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
        request.node._title = "Smoke Test -> Validate Search LIVESLR Page Navigation, download the reports and " \
                              "verify the downloaded filename for Oncology and Non-Oncology Population"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        scenarios = ['scenario1', 'scenario2']

        for scenario in scenarios:
            try:
                # Read population data values
                pop_list = exbase.get_population_data(filepath, 'Sheet1', scenario)
                # Read slrtype data values
                slrtype = exbase.get_slrtype_data(filepath, 'Sheet1', scenario)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(filepath, scenario)
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, scenario, 'Sheet1', 'Project')

                for i in pop_list:
                    slrreport.select_data(i[0], i[1], env)
                    for index, j in enumerate(slrtype):
                        slrreport.select_data(j[0], j[1], env)
                        if len(add_criteria) != 0:
                            for k in add_criteria:
                                slrreport.select_sub_section(f"{k[0]}", f"{k[1]}", env, f"{k[2]}")

                        slrreport.generate_download_report("excel_report", env)
                        excel_filename = slrreport.get_and_validate_filename(filepath)

                        if project_name[0] == 'Oncology':
                            slrreport.generate_download_report("word_report", env)
                            word_filename = slrreport.get_and_validate_filename(filepath)

                        slrreport.preview_result("preview_results", env)
                        slrreport.table_display_check("Table", env)
                        slrreport.generate_download_report("Export_as_excel", env)
                        webexcel_filename = slrreport.get_and_validate_filename(filepath)
                        slrreport.back_to_report_page("Back_to_search_page", env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C26860
    def test_reports_comparison(self, extra, env, request, caseid):
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
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = liveslrpage.get_reported_variables(filepath)
        # Read Expected Excel file
        source_template = slrreport.get_source_template(filepath, 'ExpectedSourceTemplateFile_Excel')

        request.node._tcid = caseid
        request.node._title = "Validate Downloaded Reports content comparison from Search LIVESLR page"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        try:
            for i in pop_list:
                slrreport.select_data(i[0], i[1], env)
                for index, j in enumerate(slrtype):
                    slrreport.select_data(j[0], j[1], env)
                    if j[0] == "Clinical":
                        slrreport.select_sub_section(rpt_data[3], rpt_data_chkbox[3], env, "reported_variable_section")

                    slrreport.generate_download_report("excel_report", env)
                    excel_filename = slrreport.get_and_validate_filename(filepath)

                    slrreport.generate_download_report("word_report", env)
                    word_filename = slrreport.get_and_validate_filename(filepath)

                    slrreport.preview_result("preview_results", env)
                    slrreport.table_display_check("Table", env)
                    slrreport.generate_download_report("Export_as_excel", env)
                    webexcel_filename = slrreport.get_and_validate_filename(filepath)
                    slrreport.back_to_report_page("Back_to_search_page", env)

                    slrreport.check_sorting_order_in_excel_report(webexcel_filename, excel_filename)

                    slrreport.excel_content_validation(source_template, index, webexcel_filename, excel_filename,
                                                       "Study Identifier")

                    slrreport.word_content_validation(filepath, index, word_filename)
        except Exception:
            raise Exception("Unable to select element")

    @pytest.mark.C28728
    def test_presencof_publicationtype_col_in_wordreport(self, extra, env, request, caseid):
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
        # Read reportedvariables data values
        rpt_data, rpt_data_chkbox = liveslrpage.get_reported_variables(filepath)

        request.node._tcid = caseid
        request.node._title = "Validate presence of PublicationType column in Complete Word Report"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)
        try:
            slrreport.select_data(pop_list[0][0], pop_list[0][1], env)
            slrreport.select_data(slrtype[0][0], slrtype[0][1], env)
            slrreport.select_sub_section(rpt_data[3], rpt_data_chkbox[3], env, "reported_variable_section")

            slrreport.generate_download_report("excel_report", env)
            excel_filename = slrreport.get_and_validate_filename(filepath)

            slrreport.generate_download_report("word_report", env)
            word_filename = slrreport.get_and_validate_filename(filepath)

            slrreport.preview_result("preview_results", env)
            slrreport.table_display_check("Table", env)
            slrreport.generate_download_report("Export_as_excel", env)
            webexcel_filename = slrreport.get_and_validate_filename(filepath)
            slrreport.back_to_report_page("Back_to_search_page", env)

            slrreport.presencof_publicationtype_col_in_wordreport(webexcel_filename, excel_filename, word_filename)
        except Exception:
            raise Exception("Unable to select element")

    @pytest.mark.C31466
    def test_interventional_to_clinical_changes(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getslrtestdata(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of liveslrpage class
        liveslrpage = LiveSLRPage(self.driver, extra)
        # Read population data values
        pop_list = liveslrpage.get_population_data(filepath)

        request.node._tcid = caseid
        request.node._title = "Validate label format changes -> From Interventional to Clinical"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        try:
            slrreport.test_interventional_to_clinical_changes(filepath, pop_list, env)
        except Exception:
            raise Exception("Unable to select element")

    @pytest.mark.C37419
    def test_validate_population_col_in_wordreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Population Column in Complete Word Report"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        scenarios = ['scenario1', 'scenario2']

        try:
            for i in scenarios:
                slrreport.validate_population_col_in_wordreport(self.testdata_723, i, env)
        except Exception:
            raise Exception("Unable to select element")

    @pytest.mark.C31565
    def test_validate_control_chars_in_wordreport(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Control Chars in Complete Word Report"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        scenarios = ['scenario1']

        try:
            for i in scenarios:
                slrreport.validate_control_chars_in_wordreport(self.testdata_931, i, env)
        except Exception:
            raise Exception("Unable to select element")

    @pytest.mark.C37775
    def test_nononcology_validate_ep_details_liveslr_page(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology Import Tool - Validate presence of Endpoint Details in LiveSLR -> Select " \
                              "Category(ies) to View section "

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_liveslr_data")

        scenarios = ['scenario1']

        try:
            for i in scenarios:
                slrreport.validate_presence_of_ep_details_in_liveslr_page(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C35127
    # @pytest.mark.C35150
    def test_nononcology_validate_uniquestudies_liveslr_page(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology Import Tool - Validate presence of Unique Studies for Project Level and " \
                              "SLR Type Level in Search LiveSLR page "

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_liveslr_data")

        scenarios = ['scenario1']

        try:
            for i in scenarios:
                slrreport.validate_presence_of_uniquestudies_in_liveslr_page(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C38487
    # @pytest.mark.C38489
    def test_nononcology_validate_uniquestudies_reporting_outcomes(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Verify Endpoints(EPs) present under Select Studies reporting " \
                              "Outcome(s) with all 3 EPs and with NR"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "livehta_886_data")

        scenarios = ['scenario1', 'scenario2']

        try:
            for i in scenarios:
                slrreport.validate_uniquestudies_reporting_outcomes(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C39312
    def test_nononcology_slrreport_comparison(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Validate Downloaded Reports content comparison from Search LIVESLR page"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_liveslr_report_data")

        scenarios = ['scenario1', 'scenario2']

        try:
            for scenario in scenarios:
                # Read population data values
                pop_list = exbase.get_population_data(filepath, 'Sheet1', scenario)
                # Read slrtype data values
                slrtype = exbase.get_slrtype_data(filepath, 'Sheet1', scenario)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(filepath, scenario)
                source_template = exbase.get_source_template(filepath, 'Sheet1', scenario)

                base.go_to_page("SLR_Homepage", env)
                for i in pop_list:
                    slrreport.select_data(i[0], i[1], env)
                    for index, j in enumerate(slrtype):
                        slrreport.select_data(j[0], j[1], env)
                        if len(add_criteria) != 0:
                            for k in add_criteria:
                                slrreport.select_sub_section(f"{k[0]}", f"{k[1]}", env, f"{k[2]}")

                        slrreport.generate_download_report("excel_report", env)
                        excel_filename = slrreport.get_and_validate_filename(filepath)

                        # slrreport.generate_download_report("word_report", env)
                        # word_filename = slrreport.get_and_validate_filename(filepath)

                        slrreport.preview_result("preview_results", env)
                        slrreport.table_display_check("Table", env)
                        slrreport.generate_download_report("Export_as_excel", env)
                        webexcel_filename = slrreport.get_and_validate_filename(filepath)
                        slrreport.back_to_report_page("Back_to_search_page", env)

                        slrreport.excel_content_validation(source_template, index, webexcel_filename, excel_filename,
                                                           "LiveSLR Study ID")

        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C39090
    def test_nononcology_rankorder_of_extractions(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of ExcludedStudies_liveSLR class
        exstdy_liveslr = ExcludedStudies_liveSLR(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Validate Rank order of extractions in Complete Excel and Standard " \
                              "Excel Reports"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_liveslr_report_data")

        scenarios = ['scenario1', 'scenario2']

        try:
            for scenario in scenarios:
                # Read population data values
                pop_list = exbase.get_population_data(filepath, 'Sheet1', scenario)
                # Read slrtype data values
                slrtype = exbase.get_slrtype_data(filepath, 'Sheet1', scenario)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(filepath, scenario)
                source_template = exbase.get_source_template(filepath, 'Sheet1', scenario)

                base.go_to_page("SLR_Homepage", env)
                for i in pop_list:
                    slrreport.select_data(i[0], i[1], env)
                    for index, j in enumerate(slrtype):
                        slrreport.select_data(j[0], j[1], env)
                        if len(add_criteria) != 0:
                            for k in add_criteria:
                                slrreport.select_sub_section(f"{k[0]}", f"{k[1]}", env, f"{k[2]}")

                        slrreport.generate_download_report("excel_report", env)
                        excel_filename = slrreport.get_and_validate_filename(filepath)

                        # slrreport.generate_download_report("word_report", env)
                        # word_filename = slrreport.get_and_validate_filename(filepath)

                        slrreport.preview_result("preview_results", env)
                        slrreport.table_display_check("Table", env)
                        slrreport.generate_download_report("Export_as_excel", env)
                        webexcel_filename = slrreport.get_and_validate_filename(filepath)
                        slrreport.back_to_report_page("Back_to_search_page", env)

                        slrreport.non_oncology_check_sorting_order_in_excel_report(webexcel_filename, excel_filename)

        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C40194
    def test_nononcology_validate_study_design_section(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Verify presence of Design values under Select Study Design section"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononco_studydesign_section_validation")

        scenarios = ['scenario1', 'scenario2']

        try:
            for i in scenarios:
                slrreport.validate_study_design_section(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C39828
    def test_validate_excelreportcontent_with_multiple_extractions_upload(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Verify Admin user can upload multiple files under one project/population for " \
                              "Oncology or Non-Oncology, Download the reports and validate the contents"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "livehta_1904_data")

        scenarios = ['scenario1', 'scenario2']

        try:
            for scenario in scenarios:
                # Read population data values
                pop_list = exbase.get_population_data(filepath, 'Sheet1', scenario)
                # Read slrtype data values
                slrtype = exbase.get_slrtype_data(filepath, 'Sheet1', scenario)
                # Sorting the SLR Type data to execute orderwise
                slrtype_ = sorted(list(set(tuple(sorted(sub)) for sub in slrtype)), key=lambda x: x[1])
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, scenario, 'Sheet1', 'Project')

                base.go_to_page("SLR_Dashboard", env)
                LogScreenshot.fLogScreenshot(message=f"***For '{project_name[0]}' project -> First Upload is "
                                                     f"started***", pass_=True, log=True, screenshot=False)
                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_success(scenario, filepath, env)

                slrreport.validate_content_with_multiple_extractions_upload(scenario, filepath, pop_list, slrtype_,
                                                                            env, project_name[0])

                imppubpage.delete_file(scenario, filepath, "file_status_popup_text", "upload_table_rows", env)
                LogScreenshot.fLogScreenshot(message=f"***For '{project_name[0]}' project -> First Uploaded "
                                                     f"Extraction file is removed***",
                                             pass_=True, log=True, screenshot=False)

        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C39873
    def test_upload_extraction_file_and_excel_content_validation(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology - Validate Update date under Study Characteristics"

        LogScreenshot.fLogScreenshot(message=f"***Validation of Update date for Oncology projects is started***",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        scenarios = ['pop1']

        for scenario in scenarios:
            try:
                # Read population data values
                pop_list = exbase.get_population_data(filepath, 'Sheet1', scenario)
                # Read slrtype data values
                slrtype = exbase.get_slrtype_data(filepath, 'Sheet1', scenario)
                # Sorting the SLR Type data to execute orderwise
                slrtype_ = sorted(list(set(tuple(sorted(sub)) for sub in slrtype)), key=lambda x: x[1])

                slrreport.upload_extraction_file_and_excel_content_validation(scenario, filepath, pop_list,
                                                                              slrtype_, env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error while validating Update date for Oncology projects",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error while validating Update date for Oncology projects")

        LogScreenshot.fLogScreenshot(message=f"***Validation of Update date for Oncology projects is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C41160
    def test_Smoketest_client_user(self, extra, env, request, caseid):
        clientusername = ReadConfig.getClientUserName()
        clientpassword = ReadConfig.getClientPassword()
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getsmoketestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
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
        request.node._title = "Client User -> Validate access for Oncology and Non-Oncology Population"

        loginPage.driver.get(baseURL)
        loginPage.complete_login(clientusername, clientpassword, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.go_to_page("SLR_Homepage", env)

        scenarios = ['scenario1', 'scenario2']

        for scenario in scenarios:
            try:
                # Read population data values
                pop_list = exbase.get_population_data(filepath, 'Sheet1', scenario)
                # Read slrtype data values
                slrtype = exbase.get_slrtype_data(filepath, 'Sheet1', scenario)
                add_criteria = exstdy_liveslr.get_additional_criteria_data(filepath, scenario)
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, scenario, 'Sheet1', 'Project')

                for i in pop_list:
                    slrreport.select_data(i[0], i[1], env)
                    for j in slrtype:
                        slrreport.select_data(j[0], j[1], env)
                        if len(add_criteria) != 0:
                            for k in add_criteria:
                                slrreport.select_sub_section(f"{k[0]}", f"{k[1]}", env, f"{k[2]}")

                        slrreport.generate_download_report("excel_report", env)
                        excel_filename = slrreport.get_and_validate_filename(filepath)

                        if project_name[0] == 'Oncology':
                            slrreport.generate_download_report("word_report", env)
                            word_filename = slrreport.get_and_validate_filename(filepath)

                        slrreport.preview_result("preview_results", env)
                        slrreport.table_display_check("Table", env)
                        slrreport.generate_download_report("Export_as_excel", env)
                        webexcel_filename = slrreport.get_and_validate_filename(filepath)
                        slrreport.back_to_report_page("Back_to_search_page", env)

            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page with Client User",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing LiveSLR Page with Client User")

    @pytest.mark.C41789
    def test_nononcology_validate_geographic_region_section(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Verify presence of Design values under Select Geographical " \
                              "Regions section"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononco_geographic_section_validation")

        scenarios = ['scenario1', 'scenario2']

        try:
            for i in scenarios:
                base.go_to_page("SLR_Dashboard", env)
                LogScreenshot.fLogScreenshot(
                    message=f"***Validation of presence of Geographical Regions in LiveSLR -> Select "
                            f"Geographical Regions section is started***",
                    pass_=True, log=True, screenshot=False)

                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_success(i, filepath, env)

                slrreport.validate_geographic_region_section(i, filepath, env)

                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)

                LogScreenshot.fLogScreenshot(
                    message=f"***Validation of presence of Geographical Regions in LiveSLR -> Select "
                            f"Geographical Regions section is completed***",
                    pass_=True, log=True, screenshot=False)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Element Not Found")

    @pytest.mark.C41788
    def test_nononcology_validate_geographic_region_section_clientuser(self, extra, env, request, caseid):
        clientusername = ReadConfig.getClientUserName()
        clientpassword = ReadConfig.getClientPassword()
        baseURL = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Verify presence of Design values under Select Geographical " \
                              "Regions section for Client User"

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(clientusername, clientpassword, "launch_live_slr", "Cytel LiveSLR", baseURL,
                                        env)

        filepath = exbase.get_testdata_filepath(basefile, "nononco_geographic_section_clientuser_validation")

        scenarios = ['scenario1']

        try:
            for i in scenarios:
                base.go_to_page("SLR_Dashboard", env)
                LogScreenshot.fLogScreenshot(
                    message=f"***Validation of presence of Geographical Regions in LiveSLR -> Select Geographical "
                            f"Regions section for Client user is started***",
                    pass_=True, log=True, screenshot=False)

                slrreport.validate_geographic_region_section(i, filepath, env)

                LogScreenshot.fLogScreenshot(
                    message=f"***Validation of presence of Geographical Regions in LiveSLR -> Select Geographical "
                            f"Regions section for Client user is completed***",
                    pass_=True, log=True, screenshot=False)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing LiveSLR Page with Client User",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing LiveSLR Page with Client User")
