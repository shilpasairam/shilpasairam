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
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    TestData = ReadConfig.getTestdata( "liveref_data_testing" )
    added_source = []

    @pytest.mark.C28989
    def test_add_invalid_managesourcedata(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of ManagePopulationsPage class
        self.mngsrcpage = ManageSourceDataPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Addtion of Manage Source Data validation is started***",
                                          pass_=True, log=True, screenshot=False)

        # today = date.today()
        # self.dateval = today.strftime("%m/%d/%Y")
        # self.day_val = today.day
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.liveref_complete_login(self.username, self.password, self.baseURL)

        scenarios = ['scenario1']

        for i in scenarios:
            try:
                self.mngsrcpage.go_to_managesourcedata("managesourcedata_button")
                self.mngsrcpage.add_invalid_managesourcedata(i, self.TestData)
                
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage Source Data page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")




############ For custom report generation ############
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
