"""
Test will validate the Import publications page
"""

import os
import pytest
import openpyxl
import pandas as pd
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.ImportPublicationsPage import ImportPublicationPage

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ImportPublicationPage:
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C30246
    def test_upload_extraction_template_validdata(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology Import Tool - Validate Upload Extraction Template with Success Icon"

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_success(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C27547
    def test_upload_extraction_template_with_header_mismatch(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology Import Tool - Validate Upload Extraction Template with Column Header Mismatch"

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Header Mismatch "
                                             f"validation is started***", pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop2']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Header Mismatch validation "
                                             f"is completed***", pass_=True, log=True, screenshot=False)

    @pytest.mark.C27379
    def test_upload_extraction_template_with_letters_in_publication_identifier(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology Import Tool - Validate Upload Extraction Template with Letters in " \
                              "Publication Identifier column "

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with letters in Publication Identifier "
                                             f"validation is started***", pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop3']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with letters in Publication Identifier "
                                             f"validation is completed***", pass_=True, log=True, screenshot=False)

    @pytest.mark.C27380
    def test_upload_extraction_template_with_empty_value_in_publication_identifier(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology Import Tool - Validate Upload Extraction Template with Empty value in " \
                              "Publication Identifier column "

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                             f"Identifier validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop4']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Empty value in Publication "
                                             f"Identifier validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C28986
    def test_upload_extraction_template_with_duplicate_value_in_FA19_column(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology Import Tool - Validate Upload Extraction Template with Duplicate value in " \
                              "Interventions(per arm) column "

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Duplicate value in Interventions("
                                             f"per arm) validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop5']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Duplicate value in Interventions("
                                             f"per arm) validation is completed***",
                                     pass_=True, log=True, screenshot=False)
    '''
    Commenting this TC because of LIVEHTA-1904 implementation
    @pytest.mark.C37454
    def test_upload_extraction_template_for_same_update(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology Import Tool - Validate No duplicate uploads have been made for the same " \
                              "update in the same Oncology population "
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop1']

        for index, i in enumerate(pop_list):
            try:
                imppubpage.upload_file_with_success(i, filepath, env)
                imppubpage.upload_file_for_same_population(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
    '''

    @pytest.mark.C39864
    def test_upload_extraction_template_with_invalid_date_value(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Oncology Import Tool - Validate Upload Extraction Template with Invalid data for Update Date column."

        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Invalid data for Update Date"
                                             f" validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop6']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Extraction Template with Invalid data for Update Date"
                                             f" validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C39873
    def test_upload_template_and_excel_content_validation(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getimportpublicationsdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        imppubpage = ImportPublicationPage(self.driver, extra)

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
                # Source File which is needed for comparison
                source_template = slrreport.get_source_template(filepath, 'ExpectedSourceTemplateFile')

                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_success(scenario, filepath, env)

                base.go_to_page("SLR_Homepage", env)
                for i in pop_list:
                    slrreport.select_data(i[0], i[1], env)
                    for index, j in enumerate(slrtype_):
                        slrreport.select_data(j[0], j[1], env)

                        slrreport.generate_download_report("excel_report", env)
                        excel_filename = slrreport.get_and_validate_filename(filepath)

                        slrreport.preview_result("preview_results", env)
                        slrreport.table_display_check("Table", env)
                        slrreport.generate_download_report("Export_as_excel", env)
                        webexcel_filename = slrreport.get_and_validate_filename(filepath)
                        slrreport.back_to_report_page("Back_to_search_page", env)

                        slrreport.excel_content_validation(source_template, index, webexcel_filename, excel_filename, "Study Identifier")

                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.delete_file(scenario, filepath, "file_status_popup_text", "upload_table_rows", env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error while validating Update date for Oncology projects",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error while validating Update date for Oncology projects")
        
        LogScreenshot.fLogScreenshot(message=f"***Validation of Update date for Oncology projects is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C38840
    def test_nononcology_upload_and_del_extraction_template_success(self, extra, env, request, caseid):
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

        request.node._tcid = caseid
        request.node._title = "Non-Oncology Import Tool - Validate Upload Extraction File with Success Icon"

        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_importtool")
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop1']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_success(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C38857
    def test_nononcology_upload_template_failure_for_invalid_col_id_and_colname(self, extra, env, request, caseid):
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

        request.node._tcid = caseid
        request.node._title = "Non-Oncology Import Tool - Validate Upload Extraction File with Failure Icon for " \
                              "Invalid Column ID and Column Name "

        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_importtool")
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop2']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C38858
    def test_nononcology_upload_template_failure_for_invalid_col_mapping(self, extra, env, request, caseid):
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

        request.node._tcid = caseid
        request.node._title = "Non-Oncology Import Tool - Validate Upload Extraction File with Failure Icon for " \
                              "Invalid Column Mapping "

        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_importtool")
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop3']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C39016
    def test_nononcology_upload_template_failure_for_invalid_data(self, extra, env, request, caseid):
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

        request.node._tcid = caseid
        request.node._title = "Non-Oncology Import Tool - Validate Upload Extraction File with Failure Icon for " \
                              "Invalid Data "

        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_importtool")
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop4']

        for i in pop_list:
            try:
                imppubpage.upload_file_with_errors(i, filepath, env)
                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C39841
    def test_upload_extraction_template_for_same_update(self, extra, env, request, caseid):
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

        request.node._tcid = caseid
        request.node._title = "Import Tool - Validate duplicate uploads have been made for the same update in the same Oncology or Non-Oncology population "
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_importtool")
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop5', 'pop6']

        for i in pop_list:
            try:
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, i, 'Sheet1', 'Project')

                LogScreenshot.fLogScreenshot(message=f"***For '{project_name[0]}' project -> First Upload is started***", pass_=True, log=True, screenshot=False)
                imppubpage.upload_file_with_success(i, filepath, env)

                imppubpage.upload_file_for_same_population(i, filepath, env, project_name[0])

                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
                LogScreenshot.fLogScreenshot(message=f"***For '{project_name[0]}' project -> First Uploaded Extraction file is removed***", pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
