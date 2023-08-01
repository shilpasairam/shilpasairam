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
from Pages.ManagePopulationsPage import ManagePopulationsPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.SLRReportPage import SLRReport
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ImportPublicationPage:
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C30246
    def test_upload_extraction_file_validdata(self, extra, env, request, caseid):
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
    def test_upload_extraction_file_with_header_mismatch(self, extra, env, request, caseid):
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
    def test_upload_extraction_file_with_letters_in_publication_identifier(self, extra, env, request, caseid):
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
    def test_upload_extraction_file_with_empty_value_in_publication_identifier(self, extra, env, request, caseid):
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
    def test_upload_extraction_file_with_duplicate_value_in_FA19_column(self, extra, env, request, caseid):
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
    def test_upload_extraction_file_with_invalid_date_value(self, extra, env, request, caseid):
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
        request.node._title = "Oncology Import Tool - Validate Upload Extraction Template with Invalid data for " \
                              "Update Date column."

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

    @pytest.mark.C38840
    def test_nononcology_upload_and_del_extraction_file_success(self, extra, env, request, caseid):
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
    def test_nononcology_upload_file_failure_for_invalid_col_id_and_colname(self, extra, env, request, caseid):
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
    def test_nononcology_upload_file_failure_for_invalid_col_mapping(self, extra, env, request, caseid):
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
    def test_nononcology_upload_file_failure_for_invalid_data(self, extra, env, request, caseid):
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

    @pytest.mark.C41185
    def test_nononcology_upload_file_with_same_publications(self, extra, env, request, caseid):
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
        request.node._title = "Non-Oncology Import Tool - Validate Upload Extraction File with Failure Icon " \
                              "passing same Publication for Primary and Related Publication Column"

        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_importtool")
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
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C39841
    def test_upload_extraction_file_for_same_update(self, extra, env, request, caseid):
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
        request.node._title = "Import Tool - Validate duplicate uploads have been made for the same update in the " \
                              "same Oncology or Non-Oncology population "
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_importtool")
        base.presence_of_admin_page_option("importpublications_button", env)
        base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)

        pop_list = ['pop6', 'pop7']

        for i in pop_list:
            try:
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, i, 'Sheet1', 'Project')

                LogScreenshot.fLogScreenshot(message=f"***For '{project_name[0]}' project -> First Upload is "
                                                     f"started***",
                                             pass_=True, log=True, screenshot=False)
                imppubpage.upload_file_with_success(i, filepath, env)

                imppubpage.upload_file_for_same_population(i, filepath, env, project_name[0])

                imppubpage.delete_file(i, filepath, "file_status_popup_text", "upload_table_rows", env)
                LogScreenshot.fLogScreenshot(message=f"***For '{project_name[0]}' project -> First Uploaded "
                                                     f"Extraction file is removed***",
                                             pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.C41860
    def test_nononcology_validate_C1toC5_column_in_extractionfile(self, extra, env, request, caseid):
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
        # Creating object of ManagePopulationsPage class
        mngpoppage = ManagePopulationsPage(self.driver, extra)        

        request.node._tcid = caseid
        request.node._title = "Non-Oncology - Validate Editing the existing Non-Oncology Population and " \
                              "Validate Upload Extraction File with Failure Icon for Invalid Data"

        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        livehta_2532_template_data = exbase.get_testdata_filepath(basefile, "livehta_2532_template_validation")
        livehta_2532_importool_data = exbase.get_testdata_filepath(basefile, "livehta_2532_importool_validation")        

        scenarios = ['scenario1', 'scenario2']

        for i in scenarios:
            try:
                base.refreshpage()
                base.go_to_page("managepopulations_button", env)
                mngpoppage.\
                    non_onocolgy_edit_population_by_uploading_invalid_template(i,
                                                                               'Test_NonOncology_Automation_1',
                                                                               'edit_population',
                                                                               livehta_2532_template_data,
                                                                               'Categorical', env)

                base.go_to_nested_page("importpublications_button", "extraction_upload_btn", env)
                imppubpage.upload_file_with_errors(i, livehta_2532_importool_data, env)
                imppubpage.delete_file(i, livehta_2532_importool_data, "file_status_popup_text", "upload_table_rows",
                                       env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Import publications page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload Non-Oncology Extraction Template validation is completed***",
                                     pass_=True, log=True, screenshot=False)
