import math
import time
import pandas as pd
from selenium.webdriver.support.wait import WebDriverWait
from Pages.Base import Base
from Pages.SLRReportPage import SLRReport
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot


class SearchPublicationsPage(Base):
    
    """Constructor of the SearchPublications Page class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Creating object of slrreport class
        self.slrreport = SLRReport(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)
    
    def select_data(self, locator, locator_button, env):
        time.sleep(2)        
        self.jsclick(locator, env, UnivWaitFor=10)
        if self.isselected(locator_button, env):
            self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                              pass_=True, log=True, screenshot=True)

    def select_sub_section(self, locator, locator_button, env, scroll=None):
        if self.scroll(scroll, env, UnivWaitFor=20):            
            self.jsclick(locator, env, UnivWaitFor=10)
            if self.isselected(locator_button, env):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                  pass_=True, log=True, screenshot=True)
            self.scrollback("searchpublications_page_header", env)
    
    def get_indication_details(self, locatorname, filepath, element_locator, button_locator):
        df = pd.read_excel(filepath)
        webelement = df.loc[df['Name'] == locatorname][element_locator].dropna().to_list()
        webelement_btn = df.loc[df['Name'] == locatorname][button_locator].dropna().to_list()
        result = [[webelement[i], webelement_btn[i]] for i in range(0, len(webelement))]
        return result
    
    def get_population_details(self, locatorname, filepath, element_locator, button_locator, count_locator):
        df = pd.read_excel(filepath)
        webelement = df.loc[df['Name'] == locatorname][element_locator].dropna().to_list()
        webelement_btn = df.loc[df['Name'] == locatorname][button_locator].dropna().to_list()
        webelement_count = df.loc[df['Name'] == locatorname][count_locator].dropna().to_list()
        result = [[webelement[i], webelement_btn[i], webelement_count[i]] for i in range(0, len(webelement))]
        return result

    def validate_liveref_reportname(self, env):
        expectedfilename = "LiveRef_Report.xlsx"
        self.click("liveref_generate_report", env)
        self.slrreport.table_display_check("liveref_web_table", env)
        self.slrreport.generate_download_report("liveref_export_excel_btn", env)
        # excel_filename = self.slrreport.getFilenameAndValidate(180)
        excel_filename = self.slrreport.get_latest_filename(UnivWaitFor=180)
        if excel_filename[16:] == expectedfilename:
            self.LogScreenshot.fLogScreenshot(message=f"Correct file is downloaded",
                                              pass_=True, log=True, screenshot=False)
            return excel_filename
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Filename is not as expected. Expected Filename "
                                                      f"is {expectedfilename} and Actual Filename is "
                                                      f"{excel_filename[16:]}",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Filename is not as expected. Expected Filename is {expectedfilename} and "
                            f"Filename is {excel_filename[16:]}")        
    
    def filter_count_validation(self, locatorname, filepath, env):
        # Read Indication Details
        ind = self.get_indication_details(locatorname, filepath, "Indication", "Indication_Checkbox")

        # Read Population Details
        pop = self.get_population_details(locatorname, filepath, "Population", "Population_Checkbox",
                                          "Population_Count")

        for i in pop:
            self.select_data(ind[0][0], ind[0][1], env)
            self.select_sub_section(i[0], i[1], env, scroll="population_section")
            time.sleep(1)

            selected_pop_count = self.get_text(i[2], env)

            publications_count = self.get_text("publications_count", env)

            actual_publications_msg = self.get_text("publications_display_msg", env)

            if selected_pop_count == publications_count:
                self.LogScreenshot.fLogScreenshot(message=f"Count is matching between selected population filter and "
                                                          f"the count displayed on top of the page. Selected "
                                                          f"Population Count is {selected_pop_count} and Count "
                                                          f"displayed on top of the page is {publications_count}",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Count is not matching between selected population filter "
                                                          f"and the count displayed on top of the page. Selected "
                                                          f"Population Count is {selected_pop_count} and Count "
                                                          f"displayed on top of the page is {publications_count}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Count is not matching between selected population filter and the count "
                                f"displayed on top of the page")
            
            expected_publication_msg = f"Publications: {selected_pop_count} matching"
            if expected_publication_msg == actual_publications_msg:
                self.LogScreenshot.fLogScreenshot(message=f"Publication message is displayed as expected.",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Publication message. Expected "
                                                          f"Publication Message is : {expected_publication_msg} and "
                                                          f"Actual Publication Message is : {actual_publications_msg}",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception(f"Mismatch found in Publication message")
            
            self.click("searchpublications_reset_filter", env)

    def filter_count_validation_with_Excel_report(self, locatorname, filepath, env):
        # Read Indication Details
        ind = self.get_indication_details(locatorname, filepath, "Indication", "Indication_Checkbox")

        # Read Population Details
        pop = self.get_population_details(locatorname, filepath, "Population", "Population_Checkbox",
                                          "Population_Count")

        for i in pop:
            self.select_data(ind[0][0], ind[0][1], env)
            self.select_sub_section(i[0], i[1], env, scroll="population_section")
            time.sleep(1)

            publications_count = self.get_text("publications_count", env)

            excel_filename = self.validate_liveref_reportname(env)
            table_col_data = []
            table_col_eles = self.select_elements("liveref_web_table_col1", env)
            for x in table_col_eles:
                table_col_data.append(x.text)
            # removing empty strings from list    
            web_table_data = list(filter(None, table_col_data))

            excel = pd.read_excel(f'ActualOutputs//{excel_filename}', skiprows=2)
            excel_col_data = excel['Source of Data']
            excel_col_data = [item for item in excel_col_data if str(item) != 'nan']

            if int(publications_count) != int(len(web_table_data)):
                page_counter = math.ceil(int(publications_count)/30)
                for j in range(1, page_counter):
                    self.click("liveref_web_table_next", env)
                    time.sleep(1)
                    table_col_eles1 = self.select_elements("liveref_web_table_col1", env)
                    for k in table_col_eles1:
                        web_table_data.append(k.text)
                # removing empty strings from list    
                web_table_data = list(filter(None, web_table_data))

            if int(len(excel_col_data)) == int(publications_count) == int(len(web_table_data)):
                self.LogScreenshot.fLogScreenshot(message=f"Count is matching between Publication Count, WebTable "
                                                          f"and Excel Report. Selected Population Count is "
                                                          f"{publications_count}, WebTable data row count is "
                                                          f"{len(web_table_data)} and Excel data row count is "
                                                          f"{len(excel_col_data)}",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Count is not matching between Publication Count, WebTable "
                                                          f"and Excel Report. Selected Population Count is "
                                                          f"{publications_count}, WebTable data row count is "
                                                          f"{len(web_table_data)} and Excel data row count is "
                                                          f"{len(excel_col_data)}",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Count is not matching between Publication count, WebTable and Excel Report "
                                f"data row count")
            
            self.click("liveref_back_to_selection_page", env)
            self.click("searchpublications_reset_filter", env)

    def validate_downloaded_filename(self, locatorname, filepath, env):
        # Read Indication Details
        ind = self.get_indication_details(locatorname, filepath, "Indication", "Indication_Checkbox")

        # Read Population Details
        pop = self.get_population_details(locatorname, filepath, "Population", "Population_Checkbox",
                                          "Population_Count")

        try:
            self.select_data(ind[0][0], ind[0][1], env)
            self.select_sub_section(pop[0][0], pop[0][1], env, scroll="population_section")
            time.sleep(1)            
            excel_filename = self.validate_liveref_reportname(env)
            return excel_filename
        except Exception:
            raise Exception("Error while downloading and validating the LiveRef report")

    def presence_of_author_and_affiliation_ui(self, section_locatorname, auth_locator, auth_locator_checkbox, env):

        self.select_sub_section(auth_locator, auth_locator_checkbox, env, section_locatorname)
        time.sleep(1)
        auth_text = self.get_text(auth_locator, env)

        # Checking the presence of 'Authors And Affiliations' option in LiveRef UI
        if 'Authors and Affiliations' == auth_text and self.isselected(auth_locator_checkbox, env, UnivWaitFor=10):
            self.LogScreenshot.fLogScreenshot(message=f"'Authors and Affiliations' option is present with Checkbox",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"'Authors and Affiliations' option is not present with "
                                                      f"Checkbox. Text displayed in UI : {auth_text}",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("'Authors and Affiliations' option is not present with Checkbox")

    def presence_of_author_and_affiliation_column(self, section_locatorname, auth_locator, auth_locator_checkbox, env):

        self.select_sub_section(auth_locator, auth_locator_checkbox, env, section_locatorname)
        time.sleep(1)
        if self.isselected(auth_locator_checkbox, env):
            excel_filename = self.validate_liveref_reportname(env)
            table_col_names = []
            table_col_eles = self.select_elements("liveref_web_table_col_names", env)
            for i in table_col_eles:
                table_col_names.append(i.text)
            
            # Checking the presence of column name in Web Table
            if 'Authors And Affiliations'.upper() in table_col_names:
                self.LogScreenshot.fLogScreenshot(message=f"'Authors And Affiliations' column is present in LiveRef "
                                                          f"Web table", pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"'Authors And Affiliations' column is not present in "
                                                          f"LiveRef Web table", pass_=False, log=True, screenshot=False)
                raise Exception("'Authors And Affiliations' column is not present in LiveRef Web table")
            
            # Checking the presence of column name in downloaded excel report
            excel = pd.read_excel(f'ActualOutputs//{excel_filename}', skiprows=2)
            if 'Authors And Affiliations' in excel.columns:
                self.LogScreenshot.fLogScreenshot(message=f"'Authors And Affiliations' column is present in LiveRef "
                                                          f"Excel Report", pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"'Authors And Affiliations' column is not present in "
                                                          f"LiveRef Excel Report. Column names are : {excel.columns}",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("'Authors And Affiliations' column is not present in LiveRef Excel Report")
        else:
            self.LogScreenshot.fLogScreenshot(message=f"'Authors And Affiliations' checkbox is not selected.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("'Authors And Affiliations' checkbox is not selected.")                    

    def validate_content_of_author_and_affiliation_for_previous_load(self, section_locatorname, auth_locator,
                                                                     auth_locator_checkbox, env):

        self.click("sourceofdata_section", env)
        time.sleep(1)
        self.click("sourceofdata_2020_tab", env)
        self.select_data("sourceofdata_2020_btn", "sourceofdata_2020_radio_btn", env)
        self.select_sub_section(auth_locator, auth_locator_checkbox, env, section_locatorname)
        time.sleep(1)
        if self.isselected(auth_locator_checkbox, env):
            excel_filename = self.validate_liveref_reportname(env)
            table_col_data = []
            table_col_eles = self.select_elements("liveref_web_table_auth_col", env)
            for i in table_col_eles:
                table_col_data.append(i.text)
            # removing empty strings from list    
            final_table_data = list(filter(None, table_col_data))

            excel = pd.read_excel(f'ActualOutputs//{excel_filename}', skiprows=2)
            auth_col_data = excel['Authors And Affiliations']
            auth_col_data = [item for item in auth_col_data if str(item) != 'nan']

            # Checking the presence of content in 'Authors And Affiliations' column
            if len(final_table_data) == 0 == len(auth_col_data):
                self.LogScreenshot.fLogScreenshot(message=f"In Web table and Downloaded Excel report -> 'Authors And "
                                                          f"Affiliations' column data is empty in previous load.",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"In Web table and Downloaded Excel report -> 'Authors And "
                                                          f"Affiliations' column data is not empty in previous load. "
                                                          f"Web table column data : {final_table_data} and Excel "
                                                          f"Report column data : {auth_col_data}",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("In Web table and Downloaded Excel report -> 'Authors And Affiliations' column data "
                                "is not empty in previous load")
        else:
            self.LogScreenshot.fLogScreenshot(message=f"'Authors And Affiliations' checkbox is not selected.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("'Authors And Affiliations' checkbox is not selected.")                    

    def validate_content_of_author_and_affiliation_for_latest_load(self, section_locatorname, auth_locator,
                                                                   auth_locator_checkbox, env):

        self.click("sourceofdata_section", env)
        time.sleep(1)
        self.click("sourceofdata_2022_tab", env)
        self.select_data("sourceofdata_2022_btn", "sourceofdata_2022_radio_btn", env)
        self.select_sub_section(auth_locator, auth_locator_checkbox, env, section_locatorname)
        time.sleep(1)
        if self.isselected(auth_locator_checkbox, env):
            excel_filename = self.validate_liveref_reportname(env)
            table_col_data = []
            table_col_eles = self.select_elements("liveref_web_table_auth_col", env)
            for i in table_col_eles:
                table_col_data.append(i.text)
            # removing empty strings from list    
            final_table_data = list(filter(None, table_col_data))

            excel = pd.read_excel(f'ActualOutputs//{excel_filename}', skiprows=2)
            auth_col_data = excel['Authors And Affiliations']
            auth_col_data = [item for item in auth_col_data if str(item) != 'nan']

            # Checking the presence of content in 'Authors And Affiliations' column
            if len(final_table_data) != 0 != len(auth_col_data):
                self.LogScreenshot.fLogScreenshot(message=f"In Web table and Downloaded Excel report -> 'Authors And "
                                                          f"Affiliations' column contains data in new data source.",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"In Web table and Downloaded Excel report -> 'Authors And "
                                                          f"Affiliations' column does not contains data in new data "
                                                          f"source. Web table column data : {final_table_data} and "
                                                          f"Excel Report column data : {auth_col_data}",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("In Web table and Downloaded Excel report -> 'Authors And Affiliations' column "
                                "contains data in new data source")
        else:
            self.LogScreenshot.fLogScreenshot(message=f"'Authors And Affiliations' checkbox is not selected.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("'Authors And Affiliations' checkbox is not selected.")                    
