"""
Test will validate Excluded Studies Page

"""

import os
import pytest
from Pages.Base import Base

from Pages.LoginPage import LoginPage
from Pages.ExcludedStudiesPage import ExcludedStudiesPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ExcludedStudiesPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getexcludedstudiespath()
    # slrfilepath = ReadConfig.getslrtestdata()

    '''Check excluded studies option in admin section is viewable or not'''
    @pytest.mark.C29758
    def test_view_excluded_study_option(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Presence of Manage Excluded Studies option in Admin page "
                                                  f"check is started***", pass_=True, log=True, screenshot=False)

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        # Check Manage Excluded Studies option is present in admin page or not
        self.exstdy.presence_of_elements("excluded_studies_link", env)

        self.LogScreenshot.fLogScreenshot(message=f"***Presence of Manage Excluded Studies option in Admin page "
                                                  f"check is completed***", pass_=True, log=True, screenshot=False)

    '''Check excluded studies page elements is accessible or not'''
    @pytest.mark.C29759
    def test_access_excludedstudy_page_elements(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        # # Removing the files before the test runs
        # if os.path.exists(f'ActualOutputs'):
        #     for root, dirs, files in os.walk(f'ActualOutputs'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Presence of Excluded Study Page Elements check is started***",
                                          pass_=True, log=True, screenshot=False)

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        # Go to ExcludedStudies Page
        self.base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                self.exstdy.access_excludedstudy_page_elements(i, filepath, env)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error while validating the presence of Excluded Study "
                                                      f"Page Elements",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Error while validating the presence of Excluded Study Page Elements")

        self.LogScreenshot.fLogScreenshot(message=f"***Presence of Excluded Study Page Elements check is completed***",
                                          pass_=True, log=True, screenshot=False)

    '''Addition and Deletion of Excluded Studies File'''
    @pytest.mark.C29760
    @pytest.mark.C29764
    def test_add_and_delete_excluded_study(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        # # Removing the files before the test runs
        # if os.path.exists(f'ActualOutputs'):
        #     for root, dirs, files in os.walk(f'ActualOutputs'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Excluded Studies validation is "
                                                  f"started***", pass_=True, log=True, screenshot=False)

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        # Go to ExcludedStudies Page
        self.base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                self.exstdy.add_multiple_excluded_study_data(i, filepath, env)
                self.exstdy.del_multiple_excluded_study_data(i, filepath, env)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Studies page",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing Excluded Studies page")

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion and Deletion of Excluded Studies validation is "
                                                  f"completed***", pass_=True, log=True, screenshot=False)

    '''Addition, Updation and Deletion of Excluded Studies File'''
    @pytest.mark.C29761
    @pytest.mark.C29765
    def test_update_and_delete_excluded_study(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Addition, Updation and Deletion of Excluded Studies file "
                                                  f"validation is started***", pass_=True, log=True, screenshot=False)

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_page("excluded_studies_link", env)

        pop_list = ['pop1']

        try:
            for i in pop_list:
                self.exstdy.add_multiple_excluded_study_data(i, filepath, env)
                self.exstdy.update_multiple_excluded_study_data(i, filepath, env)
                self.exstdy.del_multiple_excluded_study_data(i, filepath, env)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Studies page",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Error in accessing Excluded Studies page")

        self.LogScreenshot.fLogScreenshot(message=f"***Addition, Updation and Deletion of Excluded Studies file "
                                                  f"validation is completed***", pass_=True, log=True, screenshot=False)

    '''Compare Excluded Studies File data with Complete Excel Report'''
    @pytest.mark.C29922
    def test_excluded_study_compare_with_excel_report(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getexcludedstudiespath(env)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ExcludedStudiesPage class
        self.exstdy = ExcludedStudiesPage(self.driver, extra)

        # # Removing the files before the test runs
        # if os.path.exists(f'ActualOutputs'):
        #     for root, dirs, files in os.walk(f'ActualOutputs'):
        #         for file in files:
        #             os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Excluded Studies File comparison started***",
                                          pass_=True, log=True, screenshot=False)

        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        pop_list = ['pop1']
        
        try:
            for i in pop_list:
                self.exstdy.compare_excludedstudy_file_with_report(filepath, i, env)
                self.exstdy.del_after_studyfile_comparison(filepath, i, env)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Excluded Studies page",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Error while comparing the Excluded Studies file with Completed Excel Report")

        self.LogScreenshot.fLogScreenshot(message=f"***Excluded Studies File comparison completed***",
                                          pass_=True, log=True, screenshot=False)
