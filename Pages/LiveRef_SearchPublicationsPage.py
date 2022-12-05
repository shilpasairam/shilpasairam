from datetime import date
import os
import random
from select import select
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from Pages.Base import Base
from Pages.SLRReportPage import SLRReport
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class SearchPublicationsPage(Base):
    """Constructor of the SearchPublications Page class"""
    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
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
    
    def select_data(self, locator, locator_button):
        time.sleep(2)        
        self.jsclick(locator, UnivWaitFor=10)
        if self.isselected(locator_button):
            self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                                pass_=True, log=True, screenshot=True)

    def select_sub_section(self, locator, locator_button, scroll=None):
        if self.scroll(scroll, UnivWaitFor=20):            
            self.jsclick(locator, UnivWaitFor=10)
            if self.isselected(locator_button):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                    pass_=True, log=True, screenshot=True)
            self.scrollback("searchpublications_page_header")
    
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

    def filter_count_validation(self, locatorname, filepath):
        # Read Indication Details
        ind = self.get_indication_details(locatorname, filepath, "Indication", "Indication_Checkbox")

        # Read Population Details
        pop = self.get_population_details(locatorname, filepath, "Population", "Population_Checkbox", "Population_Count")

        for i in pop:
            self.select_data(ind[0][0], ind[0][1])
            self.select_sub_section(i[0], i[1], scroll="population_section")
            time.sleep(1)

            selected_pop_count = self.get_text(i[2])

            publications_count = self.get_text("publications_count")

            actual_publications_msg = self.get_text("publications_display_msg")

            if selected_pop_count == publications_count:
                self.LogScreenshot.fLogScreenshot(message=f"Count is matching between selected population filter and the count displayed on top of the page. Selected Population Count is {selected_pop_count} and Count displayed on top of the page is {publications_count}", 
                                                    pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Count is not matching between selected population filter and the count displayed on top of the page. Selected Population Count is {selected_pop_count} and Count displayed on top of the page is {publications_count}", 
                                                    pass_=False, log=True, screenshot=True)
                raise Exception(f"Count is not matching between selected population filter and the count displayed on top of the page")
            
            expected_publication_msg = f"Publications: {selected_pop_count} matching"
            if expected_publication_msg == actual_publications_msg:
                self.LogScreenshot.fLogScreenshot(message=f"Publication message is displayed as expected.", 
                                                    pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Mismatch found in Publication message. Expected Publication Message is : {expected_publication_msg} and Actual Publication Message is : {actual_publications_msg}", 
                                                    pass_=False, log=True, screenshot=True)
                raise Exception(f"Mismatch found in Publication message")
            
            self.click("searchpublications_reset_filter")