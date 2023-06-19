"""
Test case deals with Manage Source Data page functionality of LiveRef application
"""

import os
from re import search
import time
import pandas as pd
import pytest

from datetime import date, timedelta
from Pages.Base import Base
from Pages.LiveRef_ManageSourceDataPage import ManageSourceDataPage
from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen


@pytest.mark.usefixtures("init_driver")
class Test_ManageSourceData:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    TestData = ReadConfig.getTestdata("liveref_data_testing")
    added_source_data = []
    updated_source_data = []

    @pytest.mark.C28989
    def test_managesourcedata(self, extra, env, request, caseid):
        baseURL = ReadConfig.getPortalURL(env)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        mngsrcpage = ManageSourceDataPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate by Manage Source of Data Functionality under LiveRef"

        LogScreenshot.fLogScreenshot(message=f"***Addtion, Edit and Deletion of Manage Source Data validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_portal_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        mngsrcpage.go_to_managesourcedata("managesourcedata_button", env)

        scenarios = ['scenario1']

        for i in scenarios:
            try:
                added_src_data = mngsrcpage.add_valid_managesourcedata(i, self.TestData, "sourcedata_table_rows", env)
                # self.added_source_data.append(added_src_data)
                LogScreenshot.fLogScreenshot(message=f"Added Manage Source data are {added_src_data}",
                                             pass_=True, log=True, screenshot=False)

                updated_src_data = mngsrcpage.edit_valid_managesourcedata(i, added_src_data, self.TestData,
                                                                          "sourcedata_edit", env)
                # self.updated_source_data.append(updated_src_data)
                LogScreenshot.fLogScreenshot(message=f"Updated Manage Source data are {updated_src_data}",
                                             pass_=True, log=True, screenshot=False)

                mngsrcpage.delete_managesourcedata(updated_src_data, "sourcedata_table_rows", env)

                mngsrcpage.add_invalid_managesourcedata(i, self.TestData, env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Source Data page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        LogScreenshot.fLogScreenshot(message=f"**Addtion, Edit and Deletion of Manage Source Data validation is completed**",
                                     pass_=True, log=True, screenshot=False)

    # @pytest.mark.C28989
    # @pytest.mark.C29826
    # def test_add_valid_managesourcedata(self, extra, env, request, caseid):
    #     baseURL = ReadConfig.getPortalURL(env)
    #     # Instantiate the logScreenshot class
    #     LogScreenshot = cLogScreenshot(self.driver, extra)
    #     # Creating object of loginpage class
    #     loginPage = LoginPage(self.driver, extra)
    #     # Creating object of ManagePopulationsPage class
    #     mngsrcpage = ManageSourceDataPage(self.driver, extra)

    #     request.node._tcid = caseid
    #     request.node._title = "Validate by Adding valid Manage Source of Data under LiveRef"

    #     LogScreenshot.fLogScreenshot(message=f"***Addtion of Valid Manage Source Data validation is started***",
    #                                  pass_=True, log=True, screenshot=False)
        
    #     loginPage.driver.get(baseURL)
    #     loginPage.complete_portal_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
    #     mngsrcpage.go_to_managesourcedata("managesourcedata_button", env)

    #     scenarios = ['scenario2']

    #     for i in scenarios:
    #         try:
    #             added_src_data = mngsrcpage.add_valid_managesourcedata(i, self.TestData, "sourcedata_table_rows", env)
    #             self.added_source_data.append(added_src_data)
    #             LogScreenshot.fLogScreenshot(message=f"Added Manage Source data are {self.added_source_data}",
    #                                          pass_=True, log=True, screenshot=False)
    #         except Exception:
    #             LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Source Data page",
    #                                          pass_=False, log=True, screenshot=True)
    #             raise Exception("Element Not Found")
        
    #     LogScreenshot.fLogScreenshot(message=f"***Addtion of Valid Manage Source Data validation is completed***",
    #                                  pass_=True, log=True, screenshot=False)

    # @pytest.mark.C28989
    # @pytest.mark.C29826
    # def test_edit_valid_managesourcedata(self, extra, env, request, caseid):
    #     baseURL = ReadConfig.getPortalURL(env)
    #     # Instantiate the logScreenshot class
    #     LogScreenshot = cLogScreenshot(self.driver, extra)
    #     # Creating object of loginpage class
    #     loginPage = LoginPage(self.driver, extra)
    #     # Creating object of ManagePopulationsPage class
    #     mngsrcpage = ManageSourceDataPage(self.driver, extra)

    #     request.node._tcid = caseid
    #     request.node._title = "Validate by Editing existing Manage Source of Data under LiveRef"

    #     LogScreenshot.fLogScreenshot(message=f"**Updation of Existing Manage Source Data validation is started**",
    #                                  pass_=True, log=True, screenshot=False)
        
    #     loginPage.driver.get(baseURL)
    #     loginPage.complete_portal_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
    #     mngsrcpage.go_to_managesourcedata("managesourcedata_button", env)

    #     scenarios = ['scenario2']

    #     result = [(scenarios[i], self.added_source_data[i]) for i in range(0, len(scenarios))]

    #     for i in result:
    #         try:
    #             updated_src_data = mngsrcpage.edit_valid_managesourcedata(i[0], i[1], self.TestData,
    #                                                                       "sourcedata_edit", env)
    #             self.updated_source_data.append(updated_src_data)
    #             LogScreenshot.fLogScreenshot(message=f"Updated Manage Source data are {self.updated_source_data}",
    #                                          pass_=True, log=True, screenshot=False)
    #         except Exception:
    #             LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Source Data page",
    #                                          pass_=False, log=True, screenshot=True)
    #             raise Exception("Element Not Found")
        
    #     LogScreenshot.fLogScreenshot(message=f"*Updation of Existing Manage Source Data validation is completed*",
    #                                  pass_=True, log=True, screenshot=False)

    # @pytest.mark.C28989
    # @pytest.mark.C29826
    # def test_del_valid_managesourcedata(self, extra, env, request, caseid):
    #     baseURL = ReadConfig.getPortalURL(env)
    #     # Instantiate the logScreenshot class
    #     LogScreenshot = cLogScreenshot(self.driver, extra)
    #     # Creating object of loginpage class
    #     loginPage = LoginPage(self.driver, extra)
    #     # Creating object of ManagePopulationsPage class
    #     mngsrcpage = ManageSourceDataPage(self.driver, extra)

    #     request.node._tcid = caseid
    #     request.node._title = "Validate by Deleting existing Manage Source of Data under LiveRef"

    #     LogScreenshot.fLogScreenshot(message=f"**Deletion of Existing Manage Source Data validation is started**",
    #                                  pass_=True, log=True, screenshot=False)
        
    #     loginPage.driver.get(baseURL)
    #     loginPage.complete_portal_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
    #     mngsrcpage.go_to_managesourcedata("managesourcedata_button", env)

    #     for i in self.updated_source_data:
    #         try:
    #             mngsrcpage.delete_managesourcedata(i, "sourcedata_table_rows", env)       
    #         except Exception:
    #             LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Source Data page",
    #                                          pass_=False, log=True, screenshot=True)
    #             raise Exception("Element Not Found")
        
    #     LogScreenshot.fLogScreenshot(message=f"*Deletion of Existing Manage Source Data validation is completed*",
    #                                  pass_=True, log=True, screenshot=False)


#  ########### For custom report generation ############
# @pytest.mark.usefixtures("init_driver")
# class Test_ManageSourceData:
#     baseURL = ReadConfig.getApplicationURL()
#     username = ReadConfig.getUserName()
#     password = ReadConfig.getPassword()
#     TestData = ReadConfig.getTestdata( "liveref_data" )

#     # instantiate the logger class
#     logger = LogGen.loggen()

#     # import the wrapper testdata file
#     TSFile = pd.read_excel(TestData)
#     # find the test cases that are to be run
#     TCList = TSFile.loc[TSFile['Execute']==1,'TC_ID'].tolist()

#     # @pytest.fixture(autouse=True)
#     # def get_marker(request):
#     #     res = []
#     #     markers = request.node.own_markers
#     #     for marker in markers:
#     #         res.append(marker.name)
#     #     return res

#     @pytest.fixture(params=TCList)
#     def testcase(self,request):
#         return request.param

#     @pytest.mark.C28989
#     def test_verify_error_message_with_old_template(self,testcase,extra,request):

#         LogGen.loggen().info(f"{testcase} started")
#         # read the testcase test data file path from wrapper testdata excel file
#         path = self.TSFile.loc[self.TSFile['TC_ID']==testcase]["FilePath"].to_list()[0]
#         # read the testcase title from wrapper testdata excel file
#         title = self.TSFile.loc[self.TSFile['TC_ID']==testcase,'Title'].to_list()[0]
#         # dataFile = pd.read_csv(path)
#         # storing values from the data file to enter in the report
#         request.node._filepath = path
#         request.node._title = title
#         request.node._tcid = testcase

#         # self.driver = setup
#         # self.driver.get(self.baseURL)

#         # Complete the login process. 
#         self.loginpage = LoginPage(self.driver,extra)
#         self.loginpage.driver.get(self.baseURL)
#         self.loginpage.liveref_complete_login(self.username, self.password, self.baseURL)

#         # instantiate logscreenshot class
#         LogScreenshot = cLogScreenshot(self.driver,extra)
#         # wait for login page to load
#         # time.sleep(10)
#         # Instantiate the Base class
#         self.base = Base(self.driver, extra)

#         # initiate the test case status list
#         tc_status = []
#         scenarios = ['scenario1']
#         for i in scenarios:
            
#             try:
#                 LogScreenshot.fLogScreenshot( message = f"Iteration {i}",
#                     pass_ = True, log = True, screenshot = True)
#                 LogScreenshot.fLogScreenshot( message = f"Path value is {path}",
#                     pass_ = True, log = True, screenshot = True)
#                 LogScreenshot.fLogScreenshot( message = f"Title value is {title}",
#                     pass_ = True, log = True, screenshot = True)
#             except Exception as e:
#                 LogScreenshot.fLogScreenshot( message = f"Error in logging in : {e}",
#                     pass_ = False, log = True, screenshot = True)
#                 tc_status.append("FAIL")
#                 break
        
#             tc_status.append("PASS")

#         LogGen.loggen().info(f"{testcase} completed")

#         if "FAIL" in tc_status:
#             assert False
#             # raise enter message here
#         else:
#             assert True
