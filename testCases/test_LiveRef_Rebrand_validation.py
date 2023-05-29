"""
Test case deals in validating the LiveRef label rebrand
"""
import os
import pytest
from Pages.LiveRef_Rebrand_Check import LiveRef_Rebrand
from Pages.LoginPage import LoginPage
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_LiveRef_Rebrand:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    TestData = ReadConfig.getTestdata("liveref_searchpublications_data")

    @pytest.mark.C27354
    @pytest.mark.C29826
    def test_validate_liveref_rebrand(self, extra, env, request, caseid):
        baseURL = ReadConfig.getLiveRefAppURL(env)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of LiveRef_Rebrand class
        rebrand = LiveRef_Rebrand(self.driver, extra)

        request.node._tcid = caseid
        request.node._title = "Validate LiveRef Rebranding details"
        
        loginPage.driver.get(baseURL)
        loginPage.complete_login(self.username, self.password, "launch_liveref", "Cytel LiveRef", baseURL, env)
        try:
            rebrand.validate_liveref_rebrand(self.TestData, env)
        except Exception:
            raise Exception("Mismatch found during LiveRef brand validation")
