from datetime import date
import os
import random
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ManagePopulationsPage import ManagePopulationsPage
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select


class ManageSourceDataPage(Base):

    """Constructor of the ManageSourceData Page class"""
    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Creating object of ManagePopulationsPage class
        self.mngpoppage = ManagePopulationsPage(self.driver, extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    def go_to_managesourcedata(self, locator):
        self.click(locator, UnivWaitFor=10)
        time.sleep(5)
    
    def get_details(self, locatorname, filepath, column_locator, column_name):
        df = pd.read_excel(filepath)
        data = df.loc[df['Name'] == locatorname][column_locator].dropna().to_list()
        value = df.loc[df['Name'] == locatorname][column_name].dropna().to_list()
        result = [(data[i], value[i]) for i in range(0, len(data))]
        return result, value
    
    def get_file_details(self, locatorname, filepath, logo_column_name, temp_column_name):
        df = pd.read_excel(filepath)
        logo_file_path = os.getcwd()+(df.loc[df['Name'] == locatorname][logo_column_name].to_list()[0])
        template_file_path = os.getcwd()+(df.loc[df['Name'] == locatorname][temp_column_name].to_list()[0])
        return logo_file_path, template_file_path
    
    def add_invalid_managesourcedata(self, locatorname, filepath):
        # Read manage source details from data sheet
        source_data, source_val = self.get_details(locatorname, filepath, 'Add_source_field', 'Add_source_value')

        # Read filepaths to upload
        logo_path, template_path = self.get_file_details(locatorname, filepath, 'Source_Logo', 'Source_Template')

        self.click("add_sourcedata_btn", UnivWaitFor=10)
        time.sleep(1)

        for j in source_data:
            self.input_text(j[0], f'{j[1]}', UnivWaitFor=10)
        
        # Disabling the fileupload path to send the filepaths via script
        jscmd1 = ReadConfig.get_remove_att_JScommand(24, 'disabled')
        self.jsclick_hide(jscmd1)

        jscmd2 = ReadConfig.get_remove_att_JScommand(26, 'disabled')
        self.jsclick_hide(jscmd2)
            
        self.input_text("source_logo_file", logo_path)
        self.input_text("source_template_file", template_path)
        time.sleep(1)
        self.LogScreenshot.fLogScreenshot(message=f"Entered Source Details are: ",
                                                pass_=True, log=True, screenshot=True)

        self.click("source_save_button")
        time.sleep(3)
