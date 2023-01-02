"""
Test case deals with Login and Logout functionality of LiveRef application
"""

from re import search
import time
import pytest

from Pages.Base import Base
from Pages.LoginPage import LoginPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_Login:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C29577
    def test_about_liveref_section(self, extra):
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Instantiate the Base class
        self.base = Base(self.driver, extra)
        # instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)

        # Invoking the methods from loginpage
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef")
        
        # Checking the absence of Glossary option
        res_list = []
        eles = self.base.select_elements("glossary_nav_bar", UnivWaitFor=3)
        for ele in eles:
            res_list.append(ele.text)
        
        if not search("Glossary", res_list[0]):
            self.LogScreenshot.fLogScreenshot(message=f"Glossary link is not present as expected", 
                                                pass_=True, log=True, screenshot=True)
        else:
            raise Exception(f"Glossary link is present as expected")

        # Logging out from the application
        self.loginPage.liveref_logout()
