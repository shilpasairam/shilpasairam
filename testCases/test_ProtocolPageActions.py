"""
Test will validate the Import publications page
"""

import os
import pytest
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.ProtocolPage import ProtocolPage

from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_ProtocolPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    # filepath = ReadConfig.getprismadata()

    @pytest.mark.C30243
    def test_upload_prisma_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getprismadata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        prismapage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Uploading PRISMA details under Protocol -> PRISMA Page"

        LogScreenshot.fLogScreenshot(message=f"***Uploading PRISMA details validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "prismas", env)

        pop_val = ['pop1']

        for index, i in enumerate(pop_val):
            try:
                prismapage.add_prisma_excel_file(i, filepath, env)

                prismapage.upload_prisma_image(i, filepath, index+1, env)

                prismapage.override_prisma_details(i, filepath, index+1, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")
        
        LogScreenshot.fLogScreenshot(message=f"***Uploading PRISMA details validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C30243
    def test_del_prisma_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getprismadata(env)        
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        prismapage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate Deleting PRISMA details under Protocol -> PRISMA Page"

        LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "prismas", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                prismapage.del_prisma_excel_file(i, "prisma_excel_delete_btn", "prisma_excel_delete_popup",
                                                 filepath, env)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")

        LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37920
    def test_picos_page(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getpicosdata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        picospage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate PICOS functionality under Protocol -> PICOS Page"

        LogScreenshot.fLogScreenshot(message=f"***PICOS page validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "picos", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                picospage.add_picos_details(i, filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PICOS page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PICOS page")
        
        LogScreenshot.fLogScreenshot(message=f"***PICOS page validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37921
    def test_add_invalid_search_strategy_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getsearchstrategydata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        srchpage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate SearchStrategy functionality with invalid data under Protocol -> Search " \
                              "Strategy Page"

        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                srchpage.add_invalid_search_strategy_details(i, filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37921
    def test_add_valid_search_strategy_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getsearchstrategydata(env)
        # Instantiate the Base class
        base = Base(self.driver, extra)        
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        srchpage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate SearchStrategy functionality with valid data under Protocol -> Search " \
                              "Strategy Page"

        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                srchpage.add_valid_search_strategy_details(i, filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37883
    @pytest.mark.C38341
    def test_nononcology_upload_prisma_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        prismapage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology : Validate Uploading PRISMA details under Protocol -> PRISMA Page"

        LogScreenshot.fLogScreenshot(message=f"***Protocol -> PRISMAs page functionality validation for Non-Oncology population is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_prisma")
        
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "prismas", env)

        pop_val = ['pop1']

        for index, i in enumerate(pop_val):
            try:
                prismapage.add_prisma_excel_file(i, filepath, env)

                prismapage.upload_prisma_image(i, filepath, index+1, env)

                prismapage.override_prisma_details(i, filepath, index+1, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")
        
        LogScreenshot.fLogScreenshot(message=f"***Uploading PRISMA details validation for Non-Oncology population is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37883
    @pytest.mark.C38341
    def test_nononcology_del_prisma_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        prismapage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology : Validate Deleting PRISMA details under Protocol -> PRISMA Page"

        LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation for Non-Oncology population is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_prisma")
        
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "prismas", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                prismapage.del_prisma_excel_file(i, "prisma_excel_delete_btn", "prisma_excel_delete_popup",
                                                 filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")
        
        LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation for Non-Oncology population is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C39793
    def test_nononcology_picos_page(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        picospage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology : Validate PICOS functionality under Protocol -> PICOS Page"

        LogScreenshot.fLogScreenshot(message=f"***PICOS page validation for Non-Oncology population is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_picos")

        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "picos", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                picospage.add_picos_details(i, filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PICOS page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PICOS page")
        
        LogScreenshot.fLogScreenshot(message=f"***PICOS page validation for Non-Oncology population is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C38046
    def test_nononcology_add_invalid_search_strategy_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        srchpage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology : Validate SearchStrategy functionality with invalid data under Protocol -> Search " \
                              "Strategy Page"

        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation for Non-Oncology population is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_searchstrategy")

        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                srchpage.add_invalid_search_strategy_details(i, filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation for Non-Oncology population is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C38046
    def test_nononcology_add_valid_search_strategy_details(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
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
        srchpage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Non-Oncology : Validate SearchStrategy functionality with valid data under Protocol -> Search " \
                              "Strategy Page"

        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation for Non-Oncology population is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_searchstrategy")

        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1']

        for i in pop_val:
            try:
                srchpage.add_valid_search_strategy_details(i, filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation for Non-Oncology population is completed***",
                                     pass_=True, log=True, screenshot=False)
