import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec


class LiveSLRPage(Base):

    """Constructor of the LiveSLR Page class"""
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
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 30)

    # def go_to_liveslr(self, locator, env):
    #     self.click(locator, env, UnivWaitFor=10)
    #     self.LogScreenshot.fLogScreenshot(message='LiveSLR Search page is opened',
    #                                       pass_=True, log=True, screenshot=True)

    def get_population_data(self, filepath):
        file = pd.read_excel(filepath)
        pop = list(file['Population'].dropna())
        pop_button = list(file['Population_Radio_button'].dropna())
        population_data = [(pop[i], pop_button[i]) for i in range(0, len(pop))]
        return population_data

    def get_slrtype_data(self, filepath):
        file = pd.read_excel(filepath)
        slrtype = list(file['slrtype'].dropna())
        slrtype_button = list(file['slrtype_Radio_button'].dropna())
        slrtype_data = [(slrtype[i], slrtype_button[i]) for i in range(0, len(slrtype))]
        return slrtype_data

    def get_reported_variables(self, filepath):
        file = pd.read_excel(filepath)
        reported_var = list(file['ReportedVariables'].dropna())
        reported_var_button = list(file['Reportedvariable_checkbox'].dropna())
        # reported_var_data = [(reported_var[i], reported_var_button[i]) for i in range(0, len(reported_var))]
        return reported_var, reported_var_button

    def get_study_design(self, filepath):
        file = pd.read_excel(filepath)
        study_design = list(file['StudyDesign'].dropna())
        study_design_button = list(file['StudyDesign_checkbox'].dropna())
        # study_design_data = [(study_design[i], study_design_button[i]) for i in range(0, len(study_design))]
        return study_design, study_design_button

    def get_data_values(self, filepath):
        file = pd.read_excel(filepath)
        studydesignvalues = list(file['StudyDesignExpectedValue'].dropna())
        reportedvarvalues = list(file['ReportedVarExpectedValue'].dropna())
        return studydesignvalues, reportedvarvalues

    def get_addstudy_data(self, filepath):
        file = pd.read_excel(filepath)
        data = list(file['Add_Study_Data'].dropna())
        value = list(file['Add_Study_Data_Value'].dropna())
        result = [(data[i], value[i]) for i in range(0, len(data))]
        return result, value

    def presence_of_elements(self, locator, env):
        self.wait.until(ec.presence_of_element_located((getattr(By, self.locatortype(locator, env)),
                                                        self.locatorpath(locator, env))))
