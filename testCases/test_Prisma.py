"""
Test will validate the Import publications page
"""

import os
import pytest
from Pages.PRISMAsPage import PRISMASPage

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_PRISMAPage:
    # baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getprismadata()

    @pytest.mark.C30243
    def test_upload_prisma_details(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.prismapage = PRISMASPage(self.driver, extra)

        # Removing the files before the test runs
        if os.path.exists(f'ActualOutputs'):
            for root, dirs, files in os.walk(f'ActualOutputs'):
                for file in files:
                    os.remove(os.path.join(root, file))

        self.LogScreenshot.fLogScreenshot(message=f"***Uploading PRISMA details validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.prismapage.go_to_prisma("protocol_link", "prismas")

        pop_val = ['pop1']

        for index, i in enumerate(pop_val):
            try:
                self.prismapage.add_prisma_excel_file(i, self.filepath)

                self.prismapage.upload_prisma_image(i, self.filepath, index+1)
                
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")
        
        self.LogScreenshot.fLogScreenshot(message=f"***Uploading PRISMA details validation is completed***",
                                          pass_=True, log=True, screenshot=False)

    @pytest.mark.C30243
    def test_del_prisma_details(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.prismapage = PRISMASPage(self.driver, extra)

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation is started***",
                                          pass_=True, log=True, screenshot=False)
        
        # self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR")
        self.prismapage.go_to_prisma("protocol_link", "prismas")

        pop_val = ['pop1']

        for i in pop_val:
            try:
                self.prismapage.del_prisma_excel_file(i, "prisma_excel_delete_btn", "prisma_excel_delete_popup",
                                                      self.filepath)
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                                                  pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

        self.LogScreenshot.fLogScreenshot(message=f"***Deletion of PRISMA details validation is completed***",
                                          pass_=True, log=True, screenshot=False)
