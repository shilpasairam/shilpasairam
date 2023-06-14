"""
Test will validate Excluded Publications Page

"""

import os
import pytest
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase

from Pages.LoginPage import LoginPage
from Pages.ExcludedStudiesPage import ExcludedStudiesPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ExcludedStudiesPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getexcludedstudiespath()
    # slrfilepath = ReadConfig.getslrtestdata()

    '''Check Excluded Publications option in admin section is viewable or not'''
    @pytest.mark.C29758
    def test_view_excluded_study_option(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Presence of Manage Excluded Publications option in Admin section"

        LogScreenshot.fLogScreenshot(message=f"***Presence of Manage Excluded Publications option in Admin page "
                                             f"check is started***", pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        # Check Manage Excluded Publications option is present in admin page or not
        base.presence_of_admin_page_option("excluded_studies_link", env)

        LogScreenshot.fLogScreenshot(message=f"***Presence of Manage Excluded Publications option in Admin page "
                                             f"check is completed***", pass_=True, log=True, screenshot=False)

    '''Check Excluded Publications page elements is accessible or not'''
    @pytest.mark.C29759
    def test_access_excludedstudy_page_elements(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        exstdy = ExcludedStudiesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Manage Excluded Publications page accessibility"

        LogScreenshot.fLogScreenshot(message=f"***Presence of Excluded Study Page Elements check is started***",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("excluded_studies_link", env)
        # Go to ExcludedStudies Page
        # base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                exstdy.access_excludedstudy_page_elements(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error while validating the presence of Excluded Study "
                                                 f"Page Elements", pass_=False, log=True, screenshot=True)
            raise Exception("Error while validating the presence of Excluded Study Page Elements")

        LogScreenshot.fLogScreenshot(message=f"***Presence of Excluded Study Page Elements check is completed***",
                                     pass_=True, log=True, screenshot=False)

    '''Addition and Deletion of Excluded Publications File'''
    @pytest.mark.C29760
    @pytest.mark.C29764
    def test_add_and_delete_excluded_study(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        exstdy = ExcludedStudiesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Addition and Deletion of Excluded Publications file"

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Excluded Publications validation is "
                                             f"started***", pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("excluded_studies_link", env)
        # Go to ExcludedStudies Page
        # base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                exstdy.add_multiple_excluded_study_data(i, filepath, env)
                exstdy.del_multiple_excluded_study_data(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Publications page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing Excluded Publications page")

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Excluded Publications validation is "
                                             f"completed***", pass_=True, log=True, screenshot=False)

    '''Addition, Updation and Deletion of Excluded Publications File'''
    @pytest.mark.C29761
    @pytest.mark.C29765
    def test_update_and_delete_excluded_study(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        exstdy = ExcludedStudiesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Addition, Updation and Deletion of Excluded Publications file"

        LogScreenshot.fLogScreenshot(message=f"***Addition, Updation and Deletion of Excluded Publications file "
                                             f"validation is started***", pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("excluded_studies_link", env)
        # base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                exstdy.add_multiple_excluded_study_data(i, filepath, env)
                exstdy.update_multiple_excluded_study_data(i, filepath, env)
                exstdy.del_multiple_excluded_study_data(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Publications page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing Excluded Publications page")

        LogScreenshot.fLogScreenshot(message=f"***Addition, Updation and Deletion of Excluded Publications file "
                                             f"validation is completed***", pass_=True, log=True, screenshot=False)

    '''Compare Excluded Publications File data with Complete Excel Report'''
    @pytest.mark.C29922
    def test_excluded_study_compare_with_excel_report(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        exstdy = ExcludedStudiesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Excluded Publications file comparison with Complete Excel Report"

        LogScreenshot.fLogScreenshot(message=f"***Excluded Publications File comparison started***",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        pop_list = ['pop1']
        
        try:
            for i in pop_list:
                exstdy.compare_excludedstudy_file_with_report(filepath, i, env, 'Oncology')
                exstdy.del_after_studyfile_comparison(filepath, i, env, 'Oncology')
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Publications page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Error while comparing the Excluded Publications file with Completed Excel Report")

        LogScreenshot.fLogScreenshot(message=f"***Excluded Publications File comparison completed***",
                                     pass_=True, log=True, screenshot=False)

    '''Addition and Deletion of Excluded Publications File'''
    @pytest.mark.C38937
    def test_nononcology_add_and_delete_excluded_study(self, extra, env, request, caseid):
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
        # Creating object of ExcludedStudiesPage class
        exstdy = ExcludedStudiesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Addition and Deletion of Excluded Publications file for Non-Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Excluded Publications validation for "
                                             f"Non-Oncology Population is started***",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_manageexcludedpub")

        base.presence_of_admin_page_option("excluded_studies_link", env)
        # Go to ExcludedStudies Page
        # base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                exstdy.add_multiple_excluded_study_data(i, filepath, env)
                exstdy.del_multiple_excluded_study_data(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Publications page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing Excluded Publications page")

        LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Excluded Publications validation for "
                                             f"Non-Oncology Population is completed***",
                                     pass_=True, log=True, screenshot=False)

    '''Addition, Updation and Deletion of Excluded Publications File'''
    @pytest.mark.C38941
    def test_nononcology_update_and_delete_excluded_study(self, extra, env, request, caseid):
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
        # Creating object of ExcludedStudiesPage class
        exstdy = ExcludedStudiesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Addition, Updation and Deletion of Excluded Publications file for " \
                              "Non-Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Addition, Updation and Deletion of Excluded Publications file "
                                             f"for Non-Oncology Population validation is started***",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_manageexcludedpub")
        
        base.presence_of_admin_page_option("excluded_studies_link", env)
        # base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                exstdy.add_multiple_excluded_study_data(i, filepath, env)
                exstdy.update_multiple_excluded_study_data(i, filepath, env)
                exstdy.del_multiple_excluded_study_data(i, filepath, env)
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Publications page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing Excluded Publications page")

        LogScreenshot.fLogScreenshot(message=f"***Addition, Updation and Deletion of Excluded Publications file "
                                             f"for Non-Oncology Population validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    '''Compare Excluded Publications File data with Complete Excel Report'''
    @pytest.mark.C39311
    def test_nononcology_excluded_study_compare_with_excel_report(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        exstdy = ExcludedStudiesPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Excluded Publications file comparison with Complete Excel Report for " \
                              "Non-Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Excluded Publications File comparison for Non-Oncology Population "
                                             f"started***",
                                     pass_=True, log=True, screenshot=False)

        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_manageexcludedpub")

        pop_list = ['pop1']
        
        try:
            for i in pop_list:
                exstdy.compare_excludedstudy_file_with_report(filepath, i, env, 'Non-Oncology')
                exstdy.del_after_studyfile_comparison(filepath, i, env, 'Non-Oncology')
        except Exception:
            LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Publications page",
                                         pass_=False, log=True, screenshot=True)
            raise Exception("Error while comparing the Excluded Publications file with Completed Excel Report")

        LogScreenshot.fLogScreenshot(message=f"***Excluded Publications File comparison for Non-Oncology Population "
                                             f"completed***",
                                     pass_=True, log=True, screenshot=False)
