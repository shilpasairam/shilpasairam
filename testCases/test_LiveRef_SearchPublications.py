"""
Test case deals with Search Publications page functionality of LiveRef application
"""
import os
from re import search
import time
import pandas as pd
import pytest
from datetime import date, timedelta
from Pages.Base import Base
from Pages.LiveRef_SearchPublicationsPage import SearchPublicationsPage
from Pages.LoginPage import LoginPage
from Pages.SLRReportPage import SLRReport
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen


@pytest.mark.usefixtures("init_driver")
class Test_SearchPublications:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    TestData = ReadConfig.getTestdata("liveref_searchpublications_data")

    @pytest.mark.C29813
    @pytest.mark.C29826
    def test_filter_count_value(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
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
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        scenarios = ['scenario1']
        for i in scenarios:
            try:
                self.base.go_to_page("searchpublications_button", env)
                self.base.click("searchpublications_reset_filter", env)
                self.srchpub.filter_count_validation(i, self.TestData, env)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in during validation of filter count",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Error in during validation of filter count")

    @pytest.mark.C29566
    @pytest.mark.C29826
    def test_filter_count_value_with_excel(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
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
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        scenarios = ['scenario1']
        for i in scenarios:
            try:
                self.base.go_to_page("searchpublications_button", env)
                self.base.click("searchpublications_reset_filter", env)
                self.srchpub.filter_count_validation_with_Excel_report(i, self.TestData, env)

            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in during validation of filter count with Excel "
                                                          f"Report", pass_=False, log=True, screenshot=False)
                raise Exception("Error in during validation of filter count with Excel Report")  
    
    @pytest.mark.C27393
    @pytest.mark.C37355
    def test_presence_of_author_and_affiliation_ui(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
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
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        try:
            self.base.go_to_page("searchpublications_button", env)
            self.base.click("searchpublications_reset_filter", env)
            self.srchpub.presence_of_author_and_affiliation_ui("reported_cols_sections", "reported_col_author",
                                                               "reported_col_author_checkbox", env)

        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error during validation of Author and Affiliations option",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error during validation of Author and Affiliations option")

    @pytest.mark.C27394
    @pytest.mark.C37355
    def test_presence_of_author_and_affiliation_column(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
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
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        try:
            self.base.go_to_page("searchpublications_button", env)
            self.base.click("searchpublications_reset_filter", env)
            self.srchpub.presence_of_author_and_affiliation_column("reported_cols_sections", "reported_col_author",
                                                                   "reported_col_author_checkbox", env)

        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error during validation of Author and Affiliations "
                                                      f"column name.", pass_=False, log=True, screenshot=False)
            raise Exception("Error during validation of Author and Affiliations column name")

    @pytest.mark.C27395
    @pytest.mark.C37355
    def test_validate_content_of_author_and_affiliation_for_previous_load(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
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
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        try:
            self.base.go_to_page("searchpublications_button", env)
            self.base.click("searchpublications_reset_filter", env)
            self.srchpub.validate_content_of_author_and_affiliation_for_previous_load("reported_cols_sections",
                                                                                      "reported_col_author",
                                                                                      "reported_col_author_checkbox",
                                                                                      env)

        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error during validation of Author and Affiliations column "
                                                      f"data for previous load.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error during validation of Author and Affiliations column data for previous load")

    @pytest.mark.C27396
    @pytest.mark.C37355
    def test_validate_content_of_author_and_affiliation_for_latest_load(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)
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
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        try:
            self.base.go_to_page("searchpublications_button", env)
            self.base.click("searchpublications_reset_filter", env)
            self.srchpub.validate_content_of_author_and_affiliation_for_latest_load("reported_cols_sections",
                                                                                    "reported_col_author",
                                                                                    "reported_col_author_checkbox", env)

        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error during validation of Author and Affiliations column "
                                                      f"data for new data source.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Error during validation of Author and Affiliations column data for new data source")
