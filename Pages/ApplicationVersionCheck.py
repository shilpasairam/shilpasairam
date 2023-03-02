from re import search
from signal import raise_signal
import time
from turtle import isvisible

import pytest
import pandas as pd
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.LiveNMAPage import LiveNMA
from Pages.LiveRef_SearchPublicationsPage import SearchPublicationsPage
from Pages.LoginPage import LoginPage
from selenium.webdriver.common.by import By
from selenium.common import NoSuchElementException
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot


class AppVersion(Base):
    app_filepath = ReadConfig.getappversionfilepath()

    """Constructor of the AppVersion class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra) 
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of ExtendedBase class
        self.exbase = ExtendedBase(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)        
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)     
        # Creating object of slrreport class
        self.nma = LiveNMA(self.driver, extra) 
        # Creating object of SearchPublications class
        self.srchpub = SearchPublicationsPage(self.driver, extra)                        
        # instantiate the logger class
        self.logger = LogGen.loggen()
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
    
    def get_expected_application_version(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        version_details = df.loc[df['Application'] == locatorname]['Version'].dropna().to_list()
        return version_details
    
    def get_user_credentials(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        user_id = df.loc[df['Name'] == locatorname]['Username'].dropna().to_list()
        pwd = df.loc[df['Name'] == locatorname]['Password'].dropna().to_list()
        result = [[user_id[i], pwd[i]] for i in range(0, len(user_id))]
        return result    
    
    def validate_version_details(self, locatorname, filepath, app_name, env):
        # Read the expected app version value from test data
        expected_app_version = self.get_expected_application_version(filepath, app_name)
        # Read the actual app version
        actual_app_version = self.base.get_text(locatorname, env)
        # Compare actual value with expected value
        if expected_app_version[0] in actual_app_version:
            self.LogScreenshot.fLogScreenshot(
                message=f"Application version is as expected. Application version is : {actual_app_version}",
                pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(
                message=f"Mismatch found in Applicaiton version. Actual App version is : {actual_app_version} "
                        f"and Expected App version is : {expected_app_version}",
                pass_=False, log=True, screenshot=True)
            raise Exception(f"Mismatch found in Applicaiton version.")

    """Method to check the application version"""
    def app_version_check(self, application, about_button, about_text, about_close_btn, env):
        try:
            self.click(about_button, env, UnivWaitFor=3)
            time.sleep(1)
            # validate the actual app version with expected app version
            self.validate_version_details(about_text, self.app_filepath, application, env)
            time.sleep(1)
            self.click(about_close_btn, env, UnivWaitFor=3)
        except Exception:
            self.LogScreenshot.fLogScreenshot(
                message=f"Mismatch found in Application Version",
                pass_=False, log=True, screenshot=True)
            raise Exception("Mismatch found in Application Version")
        
    def validate_liveslr_page_access(self, locatorname, filepath, url, env):

        credentials = self.get_user_credentials(filepath, locatorname)

        self.loginPage.complete_portal_login(credentials[0][0], credentials[0][1], "launch_live_slr", "Cytel LiveSLR",
                                      url, env)

        # Read population data
        pop_data = self.exbase.get_slrtest_data(filepath, "Sheet1", locatorname, "Population",
                                                "Population_Radio_button")

        # Read slr type data
        slrtype_data = self.exbase.get_slrtest_data(filepath, "Sheet1", locatorname, "slrtype", "slrtype_Radio_button")

        # Read study design data
        stdy_dsgn = self.exbase.get_slrtest_data(filepath, "Sheet1", locatorname, "StudyDesign", "StudyDesign_checkbox")

        # Read reported variables data
        rpt_var = self.exbase.get_slrtest_data(filepath, "Sheet1", locatorname, "ReportedVariables",
                                               "Reportedvariable_checkbox")

        if search("admin", credentials[0][0]):
            self.driver.switch_to.window(self.driver.window_handles[0])
            if self.isvisible("admin_option_dropdown", env, "Admin Dropdown"):
                self.click("admin_option_dropdown", env)
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is present in LIVEHTA portal '
                                                          'for Admin user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is missing in LIVEHTA portal '
                                                          'for Admin user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin option dropdown is missing in LIVEHTA portal for Admin user.') 
            
            self.driver.switch_to.window(self.driver.window_handles[1])  
            if self.isvisible("liveslr_admin_section_text", env, "Admin Section"):
                self.LogScreenshot.fLogScreenshot(message='Admin section is present in LiveSLR homepage for '
                                                          'Admin user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin section is missing in LiveSLR homepage for '
                                                          'Admin user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin section is missing in LiveSLR homepage for Admin user.')

            self.go_to_page("SLR_Homepage", env)
            if self.isvisible(pop_data[0][0], env, "Population"):
                self.exbase.select_data(pop_data[0][0], pop_data[0][0], env)
                self.exbase.select_data(slrtype_data[0][0], slrtype_data[0][0], env)
                self.LogScreenshot.fLogScreenshot(message='Population and SLR Type is present and able to access '
                                                          'in LiveSLR homepage for Admin user.',
                                                  pass_=True, log=True, screenshot=True)
                if self.clickable("NMA_Button", env):
                    self.exbase.select_sub_section(stdy_dsgn[1][0], stdy_dsgn[1][1], env, "study_design_section")
                    self.exbase.select_sub_section(stdy_dsgn[3][0], stdy_dsgn[3][1], env, "study_design_section")
                    self.exbase.select_sub_section(rpt_var[0][0], rpt_var[0][1], env, "reported_variable_section")
                    self.exbase.select_sub_section(rpt_var[1][0], rpt_var[1][1], env, "reported_variable_section")

                    self.nma.launch_nma("launch_live_nma", env, UnivWaitFor=5)
                    self.LogScreenshot.fLogScreenshot(message=f"Admin user is able to access LIVENMA data section.",
                                                      pass_=True, log=True, screenshot=True)
                    self.nma.table_display_check("live_nma_switch_1", "live_nma_data_table", env)
                    self.nma.driver.close()
                    time.sleep(1)
                    self.nma.driver.switch_to.window(self.driver.window_handles[1])
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Admin user is not able to access LIVENMA data section.",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Admin user is not able to access LIVENMA data section.")                        
            else:
                self.LogScreenshot.fLogScreenshot(message='Population and SLR Type is not present in LiveSLR homepage '
                                                          'for Admin user.', pass_=True, log=True, screenshot=True)
                raise Exception('Population and SLR Type is not present in LiveSLR homepage for Admin user.')            

        elif search("staff", credentials[0][0]):
            self.driver.switch_to.window(self.driver.window_handles[0])
            if not self.isvisible("admin_option_dropdown", env, "Admin Dropdown"):
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is absent in LIVEHTA portal for '
                                                          'Staff user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is present in LIVEHTA portal for '
                                                          'Staff user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin option dropdown is present in LIVEHTA portal for Staff user.') 
            
            self.driver.switch_to.window(self.driver.window_handles[1])  
            if not self.isvisible("liveslr_admin_section_text", env, "Admin Section"):
                self.LogScreenshot.fLogScreenshot(message='Admin section is absent in LiveSLR homepage for '
                                                          'Staff user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin section is present in LiveSLR homepage for '
                                                          'Staff user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin section is present in LiveSLR homepage for Staff user.') 

            self.go_to_page("SLR_Homepage", env)
            if self.isvisible(pop_data[0][0], env, "Population"):
                self.exbase.select_data(pop_data[0][0], pop_data[0][0], env)
                self.exbase.select_data(slrtype_data[0][0], slrtype_data[0][0], env)
                self.LogScreenshot.fLogScreenshot(message=f"Population and SLR Type is present and able to access in "
                                                          f"LiveSLR homepage for Staff user.",
                                                  pass_=True, log=True, screenshot=True)
                if self.clickable("NMA_Button", env):
                    self.exbase.select_sub_section(stdy_dsgn[1][0], stdy_dsgn[1][1], env, "study_design_section")
                    self.exbase.select_sub_section(stdy_dsgn[3][0], stdy_dsgn[3][1], env, "study_design_section")
                    self.exbase.select_sub_section(rpt_var[0][0], rpt_var[0][1], env, "reported_variable_section")
                    self.exbase.select_sub_section(rpt_var[1][0], rpt_var[1][1], env, "reported_variable_section")

                    self.nma.launch_nma("launch_live_nma", env, UnivWaitFor=5)
                    self.LogScreenshot.fLogScreenshot(message=f"Staff user is able to access LIVENMA data section.",
                                                      pass_=True, log=True, screenshot=True)
                    self.nma.table_display_check("live_nma_switch_1", "live_nma_data_table", env)
                    self.nma.driver.close()
                    time.sleep(1)
                    self.nma.driver.switch_to.window(self.driver.window_handles[1])                                      
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Staff user is not able to access LIVENMA data section.",
                                                      pass_=False, log=True, screenshot=True)
                    raise Exception(f"Staff user is not able to access LIVENMA data section.")
            
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Population and SLR Type is not present in LiveSLR homepage "
                                                          f"for Staff user.", pass_=False, log=True, screenshot=True)
                raise Exception("Population and SLR Type is not present in LiveSLR homepage for Staff user.")
        
        elif search("client", credentials[0][0]):
            self.driver.switch_to.window(self.driver.window_handles[0])
            if not self.isvisible("admin_option_dropdown", env, "Admin Dropdown"):
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is absent in LIVEHTA portal for '
                                                          'Client user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is present in LIVEHTA portal for '
                                                          'Client user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin option dropdown is present in LIVEHTA portal for Client user.') 
            
            self.driver.switch_to.window(self.driver.window_handles[1])  
            if not self.isvisible("liveslr_admin_section_text", env, "Admin Section"):
                self.LogScreenshot.fLogScreenshot(message='Admin section is absent in LiveSLR homepage for '
                                                          'Client user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin section is present in LiveSLR homepage for '
                                                          'Client user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin section is present in LiveSLR homepage for Client user.')

            self.go_to_page("SLR_Homepage", env)
            if not self.isvisible(pop_data[0][0], env, pop_data[0][0]):
                self.LogScreenshot.fLogScreenshot(message=f"In LiveSLR -> Client user do not have access for "
                                                          f"given Population: '{pop_data[0][0]}' as expected.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"In LiveSLR -> Client user have access for "
                                                          f"given Population: '{pop_data[0][0]}'",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("In LiveSLR -> Client user have access for given Population.")

            self.presence_of_element(pop_data[1][0], env)
            if self.isvisible(pop_data[1][0], env, pop_data[1][0]):
                self.exbase.select_data(pop_data[1][0], pop_data[1][1], env)
                self.exbase.select_data(slrtype_data[0][0], slrtype_data[0][0], env)
                self.LogScreenshot.fLogScreenshot(message=f"In LiveSLR -> Client user have access for given "
                                                          f"Population: '{pop_data[1][0]}'.",
                                                  pass_=True, log=True, screenshot=True)
                if not self.isenabled("NMA_Button", env):
                    self.LogScreenshot.fLogScreenshot(message=f"Select Data for NMA button is not accessible for "
                                                              f"Client user.", pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Select Data for NMA button is accessible for "
                                                              f"Client user.", pass_=False, log=True, screenshot=True)
                    raise Exception(f"Select Data for NMA button is accessible for Client user.")              
            else:
                self.LogScreenshot.fLogScreenshot(message=f"In LiveSLR -> Client user do not have access for given "
                                                          f"Population: '{pop_data[1][0]}'.",
                                                  pass_=True, log=True, screenshot=True)
                raise Exception("In LiveSLR -> Client user do not have access for given Population as expected.")
        else:
            self.LogScreenshot.fLogScreenshot(message='Incorrect user credentials. Please pass the valid credentials.',
                                              pass_=False, log=True, screenshot=False)
            raise Exception('Incorrect user credentials. Please pass the valid credentials.')

    def validate_liveref_page_access(self, locatorname, filepath, url, env):

        credentials = self.get_user_credentials(filepath, locatorname)

        self.loginPage.complete_portal_login(credentials[0][0], credentials[0][1], "launch_liveref", "Cytel LiveRef", url, env)        

        # Read indication data
        ind_data = self.exbase.get_slrtest_data(filepath, "Sheet1", locatorname, "Indication", "Indication_Checkbox")

        # Read population data
        pop_data = self.exbase.get_slrtest_data(filepath, "Sheet1", locatorname, "LiveRef_Population",
                                                "LiveRef_Population_Checkbox")

        if search("admin", credentials[0][0]):
            self.driver.switch_to.window(self.driver.window_handles[0])
            if self.isvisible("admin_option_dropdown", env, "Admin Dropdown"):
                self.click("admin_option_dropdown", env)
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is present in LIVEHTA portal '
                                                          'for Admin user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is missing in LIVEHTA portal '
                                                          'for Admin user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin option dropdown is missing in LIVEHTA portal for Admin user.') 
            
            self.driver.switch_to.window(self.driver.window_handles[1])  
            if self.isvisible("liveref_admin_section_text", env, "Manage Selection"):
                self.LogScreenshot.fLogScreenshot(message='Manage Selection option is present in LiveRef homepage '
                                                          'for Admin user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Manage Selection option is missing in LiveRef homepage '
                                                          'for Admin user.', pass_=False, log=True, screenshot=False)
                raise Exception('Manage Selection option is missing in LiveRef homepage for Admin user.')

            if self.isvisible("liveref_importpublications_button", env, "Import Publications"):
                self.LogScreenshot.fLogScreenshot(message='Import Publications option is present in LiveRef homepage '
                                                          'for Admin user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Import Publications option is missing in LiveRef homepage '
                                                          'for Admin user.', pass_=False, log=True, screenshot=False)
                raise Exception('Import Publications option is missing in LiveRef homepage for Admin user.')

            if self.isvisible("liveref_view_import_status_button", env, "View Import Status"):
                self.LogScreenshot.fLogScreenshot(message='View Import Status option is present in LiveRef homepage '
                                                          'for Admin user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='View Import Status option is missing in LiveRef homepage '
                                                          'for Admin user.', pass_=False, log=True, screenshot=False)
                raise Exception('View Import Status option is missing in LiveRef homepage for Admin user.')

            self.go_to_page("searchpublications_button", env)
            self.click("searchpublications_reset_filter", env)
            if self.isvisible(ind_data[0][0], env, "Indication"):
                self.srchpub.select_data(ind_data[0][0], ind_data[0][1], env)
                self.srchpub.select_sub_section(pop_data[0][0], pop_data[0][1], env, scroll="population_section")
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[0][0]}' is present and able to "
                                                          f"access in LiveRef homepage for Admin user.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[0][0]}' is not present in LiveRef "
                                                          f"homepage for Admin user.",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Indication '{ind_data[0][0]}' is not present in LiveRef homepage for Admin user.")

        elif search("staff", credentials[0][0]):
            self.driver.switch_to.window(self.driver.window_handles[0])
            if not self.isvisible("admin_option_dropdown", env, "Admin Dropdown"):
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is absent in LIVEHTA portal for '
                                                          'Staff user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is present in LIVEHTA portal for '
                                                          'Staff user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin option dropdown is present in LIVEHTA portal for Staff user.') 
            
            self.driver.switch_to.window(self.driver.window_handles[1])  
            if not self.isvisible("liveref_admin_section_text", env, "Manage Selection"):
                self.LogScreenshot.fLogScreenshot(message='Manage Selection option is absent in LiveRef homepage for '
                                                          'Staff user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Manage Selection option is present in LiveRef homepage for '
                                                          'Staff user.', pass_=False, log=True, screenshot=False)
                raise Exception('Manage Selection option is present in LiveRef homepage for Staff user.')

            if not self.isvisible("liveref_importpublications_button", env, "Import Publications"):
                self.LogScreenshot.fLogScreenshot(message='Import Publications option is absent in LiveRef homepage '
                                                          'for Staff user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Import Publications option is present in LiveRef homepage '
                                                          'for Staff user.', pass_=False, log=True, screenshot=False)
                raise Exception('Import Publications option is present in LiveRef homepage for Staff user.')

            if not self.isvisible("liveref_view_import_status_button", env, "View Import Status"):
                self.LogScreenshot.fLogScreenshot(message='View Import Status option is absent in LiveRef homepage '
                                                          'for Staff user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='View Import Status option is present in LiveRef homepage '
                                                          'for Staff user.', pass_=False, log=True, screenshot=False)
                raise Exception('View Import Status option is present in LiveRef homepage for Staff user.')                
        
            self.go_to_page("searchpublications_button", env)
            self.click("searchpublications_reset_filter", env)
            if self.isvisible(ind_data[0][0], env, "Indication"):
                self.srchpub.select_data(ind_data[0][0], ind_data[0][1], env)
                self.srchpub.select_sub_section(pop_data[0][0], pop_data[0][1], env, scroll="population_section")
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[0][0]}' is present and able to "
                                                          f"access in LiveRef homepage for Staff user.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[0][0]}' is not present in LiveRef "
                                                          f"homepage for Staff user.",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Indication '{ind_data[0][0]}' is not present in LiveRef homepage for Staff user.")

        elif search("client", credentials[0][0]):
            self.driver.switch_to.window(self.driver.window_handles[0])
            if not self.isvisible("admin_option_dropdown", env, "Admin Dropdown"):
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is absent in LIVEHTA portal for '
                                                          'Client user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Admin option dropdown is present in LIVEHTA portal for '
                                                          'Client user.', pass_=False, log=True, screenshot=False)
                raise Exception('Admin option dropdown is present in LIVEHTA portal for Client user.') 
            
            self.driver.switch_to.window(self.driver.window_handles[1])  
            if not self.isvisible("liveref_admin_section_text", env, "Manage Selection"):
                self.LogScreenshot.fLogScreenshot(message='Manage Selection option is absent in LiveRef homepage for '
                                                          'Client user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Manage Selection option is present in LiveRef homepage for '
                                                          'Client user.', pass_=False, log=True, screenshot=False)
                raise Exception('Manage Selection option is present in LiveRef homepage for Client user.') 

            if not self.isvisible("liveref_importpublications_button", env, "Import Publications"):
                self.LogScreenshot.fLogScreenshot(message='Import Publications option is absent in LiveRef homepage '
                                                          'for Client user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='Import Publications option is present in LiveRef homepage '
                                                          'for Client user.', pass_=False, log=True, screenshot=False)
                raise Exception('Import Publications option is present in LiveRef homepage for Client user.')

            if not self.isvisible("liveref_view_import_status_button", env, "View Import Status"):
                self.LogScreenshot.fLogScreenshot(message='View Import Status option is absent in LiveRef homepage '
                                                          'for Client user.', pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message='View Import Status option is present in LiveRef homepage '
                                                          'for Client user.', pass_=False, log=True, screenshot=False)
                raise Exception('View Import Status option is present in LiveRef homepage for Client user.')                   
        
            self.go_to_page("searchpublications_button", env)
            self.click("searchpublications_reset_filter", env)
            if not self.isvisible(ind_data[0][0], env, "Indication"):
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[0][0]}' is not present in LiveRef "
                                                          f"homepage for Client user.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[0][0]}' is present in LiveRef "
                                                          f"homepage for Client user.",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Indication '{ind_data[0][0]}' is present in LiveRef homepage for Client user.")

            self.click("searchpublications_reset_filter", env)
            if self.isvisible(ind_data[1][0], env, "Indication"):
                self.srchpub.select_data(ind_data[1][0], ind_data[1][1], env)
                self.srchpub.select_sub_section(pop_data[1][0], pop_data[1][1], env, scroll="population_section")
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[1][0]}' is present and able to "
                                                          f"access in LiveRef homepage for Client user.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Indication '{ind_data[1][0]}' is not present in LiveRef "
                                                          f"homepage for Client user.",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Indication '{ind_data[1][0]}' is not present in LiveRef homepage for Client user.")

        else:
            self.LogScreenshot.fLogScreenshot(message='Incorrect user credentials. Please pass the valid credentials.',
                                              pass_=False, log=True, screenshot=False)
            raise Exception('Incorrect user credentials. Please pass the valid credentials.')

    # def validate_slr_project_selection(self, pop, slrtype, stdy_dsgn, rpt_var, filepath):
        
    #     self.go_to_page("SLR_Homepage")

    #     for i in pop:
    #         self.exbase.select_data(i[0], i[1])
    #         for j in slrtype:
    #             self.exbase.select_data(j[0], j[1])
    #             if j[0] == "Clinical":
    #                 self.exbase.select_sub_section(stdy_dsgn[0][0], stdy_dsgn[0][1], "study_design_section")
    #                 self.exbase.select_sub_section(rpt_var[0][0], rpt_var[0][1], "reported_variable_section")

    #             self.slrreport.generate_download_report("excel_report")
    #             excel_filename = self.slrreport.get_and_validate_filename(filepath)

    #             self.slrreport.generate_download_report("word_report")
    #             word_filename = self.slrreport.get_and_validate_filename(filepath)

    # def validate_nma_project_selection(self, pop, slrtype, stdy_dsgn, rpt_var, filepath):

    #     self.go_to_page("SLR_Homepage")

    #     try:
    #         for i in pop:
    #             self.exbase.select_data(i[0], i[1])
    #             for j in slrtype:
    #                 self.exbase.select_data(j[0], j[1])
    #                 self.exbase.select_sub_section(stdy_dsgn[0][0], stdy_dsgn[0][1], "study_design_section")
    #                 self.exbase.select_sub_section(stdy_dsgn[1][0], stdy_dsgn[1][1], "study_design_section")
    #                 self.exbase.select_sub_section(rpt_var[0][0], rpt_var[0][1], "reported_variable_section")
    #                 self.exbase.select_sub_section(rpt_var[1][0], rpt_var[1][1], "reported_variable_section")

    #                 self.nma.launch_nma("launch_live_nma", UnivWaitFor=5)
    #                 self.nma.table_display_check("live_nma_switch_1", "live_nma_data_table")
    #                 self.nma.validate_nma_selected_criteria_val(filepath, i[0], "live_nma_study_design",
    #                                                             "live_nma_reported_variable", "live_nma_pop_data")
    #                 self.nma.form_fill("add_study_button", filepath, "add_button", "live_nma_data_table_rows",
    #                                     "show_network")
    #                 time.sleep(3)
    #                 self.nma.driver.close()
    #                 self.nma.driver.switch_to.window(self.driver.window_handles[1])

    #     except Exception:
    #         raise Exception("LiveNMA Page action failed")
