"""
Test will validate the Import publications page
"""

import os
import pytest
from Pages.Base import Base
from Pages.ProtocolPage import ProtocolPage

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ProtocolPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getprismadata()

    @pytest.mark.C30243
    def test_upload_prisma_details(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getprismadata(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.prismapage = ProtocolPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Uploading PRISMA details validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_nested_page("protocol_link", "prismas", env)

        pop_val = ['pop1']

        for index, i in enumerate(pop_val):
            try:
                self.prismapage.add_prisma_excel_file(i, filepath, env)

                self.prismapage.upload_prisma_image(i, filepath, index+1, env)
                
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Uploading PRISMA details validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C30243
    def test_del_prisma_details(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getprismadata(env)        
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.prismapage = ProtocolPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_nested_page("protocol_link", "prismas", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                self.prismapage.del_prisma_excel_file(i, "prisma_excel_delete_btn", "prisma_excel_delete_popup",
                                                      filepath, env)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C37920
    def test_picos_page(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getpicosdata(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.picospage = ProtocolPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***PICOS page validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_nested_page("protocol_link", "picos", env)

        pop_val = ['pop1']

        for index, i in enumerate(pop_val):
            try:
                self.picospage.add_picos_details(i, filepath, env)
                
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing PICOS page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PICOS page")
        
        self.LogScreenshot.fLogScreenshot(message=f"***PICOS page validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C37921
    def test_add_invalid_search_strategy_details(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getsearchstrategydata(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.srchpage = ProtocolPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1']

        for index, i in enumerate(pop_val):
            try:
                self.srchpage.add_invalid_search_strategy_details(i, filepath, env)
                
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C37921
    def test_add_valid_search_strategy_details(self, extra, env):
        baseURL = ReadConfig.getApplicationURL(env)
        filepath = ReadConfig.getsearchstrategydata(env)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.srchpage = ProtocolPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        self.loginPage.driver.get(baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        self.base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1']

        for index, i in enumerate(pop_val):
            try:
                self.srchpage.add_valid_search_strategy_details(i, filepath, env)
                
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is completed***",
                                          pass_=True, log=True, screenshot=False)
