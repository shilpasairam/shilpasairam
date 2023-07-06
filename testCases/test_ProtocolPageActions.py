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
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C30243
    def test_oncology_prisma_details(self, extra, env, request, caseid):
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
        request.node._title = "Validate PRISMA functionalities details under Protocol -> PRISMA Page for " \
                              "Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Upload and Deletion of PRISMA details validation is started***",
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

                prismapage.del_prisma_excel_file(i, "prisma_excel_delete_btn", "prisma_excel_delete_popup",
                                                 filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload and Deletion of PRISMA details validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37920
    def test_picos_page(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getpicosdata(env)
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
        request.node._title = "Validate PICOS functionality under Protocol -> PICOS Page"

        LogScreenshot.fLogScreenshot(message=f"***PICOS page validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "picos", env)

        pop_val = ['pop1', 'pop2', 'pop3', 'pop4']

        for i in pop_val:
            try:
                # Read population details from data sheet
                pop_data = exbase.get_population_data(filepath, 'Sheet1', i)
                # Read study types and file paths to upload
                stdy_data = exbase.get_slrtype_data(filepath, 'Sheet1', i)
                # Removing duplicates to get the proper length of SLR Type data
                stdy_data_ = sorted(list(set(tuple(sorted(sub)) for sub in stdy_data)), key=lambda x: x[1])

                picos_data = picospage.add_picos_details(i, filepath, pop_data, stdy_data_, env, "Oncology")

                # picospage.clear_picos_details(i, filepath, pop_data, stdy_data_, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PICOS page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PICOS page")
        
        LogScreenshot.fLogScreenshot(message=f"***PICOS page validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37921
    def test_search_strategy_page_oncology(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        filepath = ReadConfig.getsearchstrategydata(env)
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
        request.node._title = "Validate SearchStrategy functionality under Protocol -> Search " \
                              "Strategy Page for Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1', 'pop2', 'pop3', 'pop4']

        for i in pop_val:
            try:
                # Read population details from data sheet
                pop_data = exbase.get_population_data(filepath, 'Sheet1', i)
                # Read study types and file paths to upload
                stdy_data = exbase.get_slrtype_data(filepath, 'Sheet1', i)
                # Sorting the SLR Type data to execute orderwise
                stdy_data_ = sorted(list(set(tuple(sorted(sub)) for sub in stdy_data)), key=lambda x: x[1])

                srchpage.add_invalid_search_strategy_details(i, filepath, pop_data, stdy_data_, env)
                uploaded_data = srchpage.add_valid_search_strategy_details(i, filepath, pop_data, stdy_data_, env,
                                                                           "Oncology", "pagelevel")
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C37883
    @pytest.mark.C38341
    def test_nononcology_prisma_details(self, extra, env, request, caseid):
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
        request.node._title = "Non-Oncology : Validate PRISMA functionalities details under Protocol -> PRISMA Page"

        LogScreenshot.fLogScreenshot(message=f"***Upload and Deletion of PRISMA details validation for Non-Oncology "
                                             f"population is completed***",
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

                prismapage.del_prisma_excel_file(i, "prisma_excel_delete_btn", "prisma_excel_delete_popup",
                                                 filepath, env)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PRISMA page")
        
        LogScreenshot.fLogScreenshot(message=f"***Upload and Deletion of PRISMA details validation for Non-Oncology "
                                             f"population is completed***",
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

        pop_val = ['pop1', 'pop2', 'pop3', 'pop4']

        for i in pop_val:
            try:
                # Read population details from data sheet
                pop_data = exbase.get_population_data(filepath, 'Sheet1', i)
                # Read study types and file paths to upload
                stdy_data = exbase.get_slrtype_data(filepath, 'Sheet1', i)
                # Removing duplicates to get the proper length of SLR Type data
                stdy_data_ = sorted(list(set(tuple(sorted(sub)) for sub in stdy_data)), key=lambda x: x[1])

                picospage.add_picos_details(i, filepath, pop_data, stdy_data_, env, "Non-Oncology")
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing PICOS page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing PICOS page")
        
        LogScreenshot.fLogScreenshot(message=f"***PICOS page validation for Non-Oncology population is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C38046
    def test_search_strategy_page_nononcology(self, extra, env, request, caseid):
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
        request.node._title = "Validate SearchStrategy functionality under Protocol -> Search " \
                              "Strategy Page for Non-Oncology Population"

        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation for Non-Oncology population "
                                             f"is started***",
                                     pass_=True, log=True, screenshot=False)
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "nononcology_searchstrategy")

        base.presence_of_admin_page_option("protocol_link", env)
        base.go_to_nested_page("protocol_link", "searchstrategy", env)

        pop_val = ['pop1', 'pop2', 'pop3', 'pop4']

        for i in pop_val:
            try:
                # Read population details from data sheet
                pop_data = exbase.get_population_data(filepath, 'Sheet1', i)
                # Read study types and file paths to upload
                stdy_data = exbase.get_slrtype_data(filepath, 'Sheet1', i)
                # Removing duplicates to get the proper length of SLR Type data
                stdy_data_ = sorted(list(set(tuple(sorted(sub)) for sub in stdy_data)), key=lambda x: x[1])

                srchpage.add_invalid_search_strategy_details(i, filepath, pop_data, stdy_data_, env)
                uploaded_data = srchpage.add_valid_search_strategy_details(i, filepath, pop_data, stdy_data_, env,
                                                                           "Non-Oncology", "pagelevel")
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Search Strategy page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Search Strategy page")
        
        LogScreenshot.fLogScreenshot(message=f"***Search Strategy page validation for Non-Oncology population "
                                             f"is completed***",
                                     pass_=True, log=True, screenshot=False)

    @pytest.mark.C40555
    def test_download_protocol_excel(self, extra, env, request, caseid):
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
        request.node._title = "Validate Downloading Protocol Excel file for Oncology and Non-Oncology population " \
                              "from Search LiveSLR Page"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)
        base.presence_of_admin_page_option("protocol_link", env)
        
        filepath = exbase.get_testdata_filepath(basefile, "download_protocol_excel")

        pop_val = ['pop1', 'pop2', 'pop3', 'pop4', 'pop5']

        for i in pop_val:
            try:
                # Read population details from data sheet
                pop_data = exbase.get_population_data(filepath, 'Sheet1', i)
                # Read study types and file paths to upload
                stdy_data = exbase.get_slrtype_data(filepath, 'Sheet1', i)
                # Removing duplicates to get the proper length of SLR Type data
                stdy_data_ = sorted(list(set(tuple(sorted(sub)) for sub in stdy_data)), key=lambda x: x[1])
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, i, 'Sheet1', 'Project')

                LogScreenshot.fLogScreenshot(message=f"***Protocol Excel file validation for '{project_name[0]}' "
                                                     f"project is started***",
                                             pass_=True, log=True, screenshot=False)

                picospage.download_protocol_file(i, filepath, pop_data, stdy_data_, env, project_name[0])

                LogScreenshot.fLogScreenshot(message=f"***Protocol Excel file validation for '{project_name[0]}' "
                                                     f"project is completed***",
                                             pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing Protocol Excel file",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing Protocol Excel file")

    @pytest.mark.C41136
    def test_liveslr_view_picos(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        picospage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate View PICOS functionality under Search LiveSLR Page for Oncology and " \
                              "Non-Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "download_protocol_excel")

        pop_val = ['pop1', 'pop2', 'pop3', 'pop4', 'pop5']

        for i in pop_val:
            try:
                # Read population details from data sheet
                pop_data = exbase.get_population_data(filepath, 'Sheet1', i)
                # Read study types and file paths to upload
                stdy_data = exbase.get_slrtype_data(filepath, 'Sheet1', i)
                # Removing duplicates to get the proper length of SLR Type data
                stdy_data_ = sorted(list(set(tuple(sorted(sub)) for sub in stdy_data)), key=lambda x: x[1])
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, i, 'Sheet1', 'Project')

                LogScreenshot.fLogScreenshot(message=f"***View PICOS validation for '{project_name[0]}' "
                                                     f"project is started***",
                                             pass_=True, log=True, screenshot=False)

                picospage.validate_view_picos(i, filepath, pop_data, stdy_data_, env, project_name[0])

                LogScreenshot.fLogScreenshot(message=f"***View PICOS validation for '{project_name[0]}' "
                                                     f"project is completed***",
                                             pass_=True, log=True, screenshot=False)
                
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing View PICOS page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing View PICOS page")

    @pytest.mark.C40680
    def test_liveslr_view_searchstrategy(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveSLRAppURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        picospage = ProtocolPage(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate View Search Strategy functionality under Search LiveSLR Page for " \
                              "Oncology and Non-Oncology Population"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", baseURL, env)

        filepath = exbase.get_testdata_filepath(basefile, "download_protocol_excel")

        pop_val = ['pop1', 'pop2', 'pop3', 'pop4', 'pop5']

        for i in pop_val:
            try:
                # Read population details from data sheet
                pop_data = exbase.get_population_data(filepath, 'Sheet1', i)
                # Read study types and file paths to upload
                stdy_data = exbase.get_slrtype_data(filepath, 'Sheet1', i)
                # Removing duplicates to get the proper length of SLR Type data
                stdy_data_ = sorted(list(set(tuple(sorted(sub)) for sub in stdy_data)), key=lambda x: x[1])
                # Read Project name
                project_name = exbase.get_individual_col_data(filepath, i, 'Sheet1', 'Project')

                LogScreenshot.fLogScreenshot(message=f"***View Search Strategy for '{project_name[0]}' "
                                                     f"project is started***",
                                             pass_=True, log=True, screenshot=False)

                picospage.validate_view_search_strategy(i, filepath, pop_data, stdy_data_, env, project_name[0])

                LogScreenshot.fLogScreenshot(message=f"***View Search Strategy for '{project_name[0]}' "
                                                     f"project is completed***",
                                             pass_=True, log=True, screenshot=False)
            except Exception:
                LogScreenshot.fLogScreenshot(message=f"Error in accessing View Search Strategy page",
                                             pass_=False, log=True, screenshot=True)
                raise Exception("Error in accessing View Search Strategy page")
