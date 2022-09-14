"""
This class is the parent of all pages
It contains all the generic methods and utilites for all the pages
"""
import functools
import time
import pandas as pd
from selenium.common import NoSuchElementException

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig, config


# decorator to wait for action to be executed
def fWaitFor(input_func):
    @functools.wraps(input_func)
    def wrapper(*args, **kwargs):
        try:
            UnivWaitFor = kwargs['UnivWaitFor']
        except KeyError:
            UnivWaitFor = 0
        # simply return the decorated function if waitFor argument isn't explicitly provided
        if UnivWaitFor == 0:
            return input_func(*args, **kwargs)
        # else ignore any error till {UnivWaitFor} seconds. If the functions executes, return the function's output
        else:
            errorPresent = True
            timePassed = 0
            while errorPresent == True and timePassed < UnivWaitFor:
                try:
                    result = input_func(*args, **kwargs)
                    errorPresent = False
                    return result
                except Exception:
                    timePassed += 1
                    time.sleep(1)
            # if function fails even after {UnivWaitFor} seconds, return the error as is
            if errorPresent:
                return input_func(*args, **kwargs)

    return wrapper


class Base:

    def __init__(self, driver, extra):
        self.driver = driver
        self.extra = extra
        self.wait = WebDriverWait(driver, 10)
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)

    # read object repository csv to get a list of lists
    @fWaitFor
    def read_data_from_csv(self, filename, UnivWaitFor=0):
        df = pd.read_csv(filename)
        return df

    @staticmethod
    def getORFilePath():
        OR = config.get('commonInfo', 'OR')
        return OR

    # identify locatorpath from locationname-object in csv
    @fWaitFor
    def locatorpath(self, locatorname, UnivWaitFor=0):
        df = self.read_data_from_csv(ReadConfig.getORFilePath())
        return df.loc[df['Object'] == locatorname]['Path'].to_list()[0]

    # identify locatortype from locationname-object in csv
    @fWaitFor
    def locatortype(self, locatorname, UnivWaitFor=0):
        df = self.read_data_from_csv(ReadConfig.getORFilePath())
        return df.loc[df['Object'] == locatorname]['Type'].to_list()[0]

    # click a web element using locatorname
    @fWaitFor
    def click(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and click on the element
        """
        self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).click()

    # input text to web element using locatorname and value
    @fWaitFor
    def input_text(self, locator, value, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and input text to the element
        """
        self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).clear()
        self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).send_keys(value)

    # Select current date from date-picker
    @fWaitFor
    def select_calendar_date(self, day_val, UnivWaitFor=0):
        """
        Given day value, select the current date from the date-picker
        """
        cal_loc = f"//table[@class='days weeks']//td[@role='gridcell']//span[not(contains(@class, 'is-other-month')) and (text()={day_val})]"
        self.driver.find_element(By.XPATH, cal_loc).click()

    # clear text from web element using locatorname and value
    @fWaitFor
    def clear(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and clear the text
        """
        self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).clear()

    # refresh the webpage
    @fWaitFor
    def refreshpage(self):
        """
        Given locator, identify the locator type and path from the OR file and clear the text
        """
        self.driver.refresh()

    # Assertion validation using pagetitle
    @fWaitFor
    def assertPageTitle(self, pageTitle, UnivWaitFor=0):
        """
        Assert the title of the page
        """
        assert self.driver.title == pageTitle

    # Read the element text using locatorname
    @fWaitFor
    def get_text(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and return the text
        """
        return self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).text

    # function to select an element
    @fWaitFor
    def select_element(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and select the element
        """
        element = self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator))
        return element

    # function to select multiple elements
    @fWaitFor
    def select_elements(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and select the elements
        """
        elements = self.driver.find_elements(getattr(By, self.locatortype(locator)), self.locatorpath(locator))
        return elements

    # JavaScript click a web element using locatorname
    @fWaitFor
    def jsclick(self, locator, message="", UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and click on the element
        """
        try:
            self.driver.execute_script("arguments[0].click();", self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)))
        except NoSuchElementException:
            # self.LogScreenshot.fLogScreenshot(message=f"{locator} is not present",
            #                                   pass_=False, log=True, screenshot=False)
            self.LogScreenshot.fLogScreenshot(message=f"{message}",
                                              pass_=False, log=True, screenshot=False)

    # JavaScript click to hide a element for uploading
    @fWaitFor
    def jsclick_hide(self, command, UnivWaitFor=0):
        """
        Execute the JS statement with given command
        """
        try:
            self.driver.execute_script(command)
        except NoSuchElementException:
            self.LogScreenshot.fLogScreenshot(message=f"{command} is not executable",
                                              pass_=False, log=True, screenshot=False)
    
    # Check whether a web element is clickable or not using locatorname
    @fWaitFor
    def clickable(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and check element is clickable or not
        """
        return self.wait.until(ec.element_to_be_clickable((getattr(By, self.locatortype(locator)), self.locatorpath(locator))))

    # Check whether element is selected or not using locatorname
    @fWaitFor
    def isselected(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and return the bool value
        """
        try:
            return self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).is_selected()
        except NoSuchElementException:
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is not present",
                                              pass_=False, log=True, screenshot=False)

    # Scroll to the element using locatorname
    @fWaitFor
    def scroll(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and scroll to the element
        """
        try:
            self.driver.execute_script("arguments[0].scrollIntoView(true);", self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)))
            self.jsclick(locator)
            return True
        except NoSuchElementException:
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is not present",
                                              pass_=False, log=True, screenshot=False)
            return False

    # Scroll back to the element using locatorname
    @fWaitFor
    def scrollback(self, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and scroll to the element
        """
        self.driver.execute_script("arguments[0].scrollIntoView(true);", self.driver.find_element(By.XPATH, "//p[text()='Search LiveSLR™']"))

    # Check whether a web element is displayed or not using locatorname
    @fWaitFor
    def isdisplayed(self, locator, UnivWaitFor=0):
        """
        Given locator, identify the locator type and path from the OR file and return the bool value
        """
        try:
            return self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).is_displayed()
        # return self.wait.until(ec.presence_of_element_located((getattr(By, self.locatortype(locator)), self.locatorpath(locator))))
        except NoSuchElementException:
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is not present",
                                              pass_=False, log=True, screenshot=False)

    # Assertion validation for text
    @fWaitFor
    def assertText(self, expected_text, actual_text, UnivWaitFor=0):
        """
        Assert the text
        """
        assert expected_text == actual_text
