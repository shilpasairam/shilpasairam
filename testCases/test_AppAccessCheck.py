"""
Test case deals with Application access check for LiveSLR and LiveRef
"""

import os
import time
import pytest
from re import search
from Pages.Base import Base
from Pages.ApplicationVersionCheck import AppVersion
from Pages.ExtendedBasePage import ExtendedBase
from Pages.LoginPage import LoginPage
from selenium.webdriver import ActionChains

from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_AppAccess:
    # testdata = ReadConfig.getTestdata("app_access_data")

    # @pytest.mark.parametrize(
    #     "scenarios, name",
    #     [
    #         ("scenario1", "Admin User"),
    #         ("scenario2", "Staff User"),
    #         ("scenario3", "Client User")
    #     ]
    # )
    @pytest.mark.C37915
    def test_livehta_application_access(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        testdata = ReadConfig.getappaccesstestdata(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)        
        # Creating object of AppVersion class
        appver = AppVersion(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)        

        request.node._tcid = caseid
        request.node._title = f"Validate LiveHTA Application Access for Admin, Staff and Client users"

        # Invoking the methods from loginpage
        loginPage.driver.get(baseURL)

        scenarios = ['scenario1', 'scenario2', 'scenario3']

        for scenario in scenarios:
            # Read Type of User details
            usertype1 = exbase.get_individual_col_data(testdata, scenario, 'Sheet1', 'Usertype')
            LogScreenshot.fLogScreenshot(message=f"***Validation of LiveSLR Application Access for {usertype1[0]} "
                                                 f"is Started***",
                                         pass_=True, log=True, screenshot=False)
            # Validating the application version
            appver.validate_liveslr_page_access(scenario, testdata, baseURL, env)

            # Logging out from the application
            loginPage.logout("liveslr_logout_button", env)
            loginPage.driver.close()
            loginPage.driver.switch_to.window(self.driver.window_handles[0])
            loginPage.logout("portal_logout_button", env)
            LogScreenshot.fLogScreenshot(message=f"***Validation of LiveSLR Application Access for {usertype1[0]} "
                                                 f"is Completed***",
                                         pass_=True, log=True, screenshot=False)

        for scenario in scenarios:
            # Read Type of User details
            usertype2 = exbase.get_individual_col_data(testdata, scenario, 'Sheet1', 'Usertype')
            LogScreenshot.fLogScreenshot(message=f"***Validation of LiveRef Application Access for {usertype2[0]} "
                                                 f"is Started***",
                                         pass_=True, log=True, screenshot=False)
            # Validating the application version
            appver.validate_liveref_page_access(scenario, testdata, baseURL, env)

            # Logging out from the application
            loginPage.logout("liveref_logout_button", env)
            loginPage.driver.close()
            loginPage.driver.switch_to.window(self.driver.window_handles[0])
            loginPage.logout("portal_logout_button", env)
            LogScreenshot.fLogScreenshot(message=f"***Validation of LiveRef Application Access for {usertype2[0]} "
                                                 f"is Completed***",
                                         pass_=True, log=True, screenshot=False)

    # @pytest.mark.parametrize(
    #     "scenarios, name",
    #     [
    #         ("scenario1", "Admin User"),
    #         ("scenario2", "Staff User"),
    #         ("scenario3", "Client User")
    #     ]
    # )
    # # @pytest.mark.C37915
    # def test_liveref_application_access(self, scenarios, name, extra, env, request, caseid):
    #     baseURL = ReadConfig.getPortalURL(env)
    #     testdata = ReadConfig.getappaccesstestdata(env)
    #     # Creating object of loginpage class
    #     loginPage = LoginPage(self.driver, extra)        
    #     # Creating object of AppVersion class
    #     appver = AppVersion(self.driver, extra)

    #     request.node._tcid = caseid
    #     request.node._title = f"Validate LiveRef Application Access - {scenarios} - {name}"

    #     # Invoking the methods from loginpage
    #     loginPage.driver.get(baseURL)
        
    #     # Validating the application version
    #     appver.validate_liveref_page_access(scenarios, testdata, baseURL, env)

    #     # Logging out from the application
    #     loginPage.logout("liveref_logout_button", env)

    @pytest.mark.parametrize(
        "scenarios, name",
        [
            ("scenario1", "Admin User"),
            ("scenario2", "Staff User"),
            ("scenario3", "Client User")
        ]
    )
    @pytest.mark.C40338
    def test_liveslr_application_access_from_cytel(self, scenarios, name, extra, env, request, caseid):
        baseURL = "https://www.cytel.com/"
        testdata = ReadConfig.getappaccesstestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)           
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)        
        # Creating object of AppVersion class
        appver = AppVersion(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)        

        request.node._tcid = caseid
        request.node._title = f"Validate LiveSLR Application Access from Cytel.com - {scenarios} - {name}"

        if env == "production":
            loginPage.driver.get(baseURL)            
            base.click("cookie_accept_btn", env)
            LogScreenshot.fLogScreenshot(message=f"Cytel Main Page",
                                         pass_=True, log=True, screenshot=True)            
            login_ele = base.select_element("login_dropdown", env)
            actions = ActionChains(self.driver)
            actions.move_to_element(login_ele).perform()

            hta_ele = base.select_element("livehta_portal", env)
            actions.move_to_element(hta_ele).click().perform()
            
            # Validating the application version
            appver.validate_liveslr_page_access(scenarios, testdata, baseURL, env)

            # Logging out from the application
            loginPage.logout("liveslr_logout_button", env)
        else:
            LogScreenshot.fLogScreenshot(message=f"This Test Case is specifically designed to run only in Production "
                                                 f"environment. For '{str(env).upper()}' environment this TC is being "
                                                 f"ignored as we do not have dedicated cytel.com site available for "
                                                 f"'{str(env).upper()}' environment.",
                                         pass_=True, log=True, screenshot=False)

    @pytest.mark.parametrize(
        "scenarios, name",
        [
            ("scenario1", "Admin User"),
            ("scenario2", "Staff User"),
            ("scenario3", "Client User")
        ]
    )
    @pytest.mark.C40338
    def test_liveref_application_access_from_cytel(self, scenarios, name, extra, env, request, caseid):
        baseURL = "https://www.cytel.com/"
        testdata = ReadConfig.getappaccesstestdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)           
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)        
        # Creating object of AppVersion class
        appver = AppVersion(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)        

        request.node._tcid = caseid
        request.node._title = f"Validate LiveRef Application Access from Cytel.com - {scenarios} - {name}"

        if env == "production":
            loginPage.driver.get(baseURL)
            base.click("cookie_accept_btn", env)
            LogScreenshot.fLogScreenshot(message=f"Cytel Main Page",
                                         pass_=True, log=True, screenshot=True)            
            login_ele = base.select_element("login_dropdown", env)
            actions = ActionChains(self.driver)
            actions.move_to_element(login_ele).perform()

            hta_ele = base.select_element("livehta_portal", env)
            actions.move_to_element(hta_ele).click().perform()
            
            # Validating the application version
            appver.validate_liveref_page_access(scenarios, testdata, baseURL, env)

            # Logging out from the application
            loginPage.logout("liveref_logout_button", env)
        else:
            LogScreenshot.fLogScreenshot(message=f"This Test Case is specifically designed to run only in Production "
                                                 f"environment. For '{str(env).upper()}' environment this TC is being "
                                                 f"ignored as we do not have dedicated cytel.com site available for "
                                                 f"'{str(env).upper()}' environment.",
                                         pass_=True, log=True, screenshot=False)
