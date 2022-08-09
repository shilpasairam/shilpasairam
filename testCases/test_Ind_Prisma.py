"""
Test will validate the Import publications page
"""

import pytest
from Pages.PRISMAsPage import PRISMASPage

from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


@pytest.mark.usefixtures("init_driver")
class Test_PRISMAPage:
    baseURL = ReadConfig.getApplicationURL()
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()
    filepath = ReadConfig.getprismadata()

    @pytest.mark.LIVEHTA_568
    def test_upload_prisma_details(self, extra):
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of loginpage class
        self.loginPage = LoginPage(self.driver, extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, extra)
        # Creating object of ImportPublicationPage class
        self.prismapage = PRISMASPage(self.driver, extra)
        # Read extraction sheet values to upload excel files
        self.excel_file = self.prismapage.get_prisma_excelfile_details(self.filepath) 
        # Read extraction sheet values to upload prisma image for study data
        self.stdy_data = self.prismapage.get_prisma_data_details(self.filepath) 
        
        self.loginPage.driver.get(self.baseURL)
        self.loginPage.complete_login(self.username, self.password)
        self.prismapage.go_to_prisma("protocol_link", "prismas")

        for index, i in enumerate(self.excel_file):
            try:
                self.prismapage.add_prisma_excel_file(index+1, i)

                self.prismapage.upload_prisma_image(self.stdy_data, index+1)
                
            except Exception:
                self.LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                    pass_=False, log=True, screenshot=True)
                raise Exception("Element Not Found")

    @pytest.mark.LIVEHTA_568
    def test_del_prisma_details(self, extra):
            # Instantiate the logScreenshot class
            self.LogScreenshot = cLogScreenshot(self.driver, extra)
            # Creating object of loginpage class
            self.loginPage = LoginPage(self.driver, extra)
            # Creating object of liveslrpage class
            self.liveslrpage = LiveSLRPage(self.driver, extra)
            # Creating object of ImportPublicationPage class
            self.prismapage = PRISMASPage(self.driver, extra)
            # Read extraction sheet values to upload excel files
            self.excel_file = self.prismapage.get_prisma_excelfile_details(self.filepath) 
            # Read extraction sheet values to upload prisma image for study data
            self.stdy_data = self.prismapage.get_prisma_data_details(self.filepath) 
            
            self.loginPage.driver.get(self.baseURL)
            self.loginPage.complete_login(self.username, self.password)
            self.prismapage.go_to_prisma("protocol_link", "prismas")

            for index, i in enumerate(self.excel_file):
                try:
                    self.prismapage.del_prisma_excel_file(index+1, "prisma_excel_delete_btn", "prisma_excel_delete_popup", self.stdy_data)                    
                except Exception:
                    self.LogScreenshot.fLogScreenshot(message=f"Error in accessing PRISMA page",
                        pass_=False, log=True, screenshot=True)
                    raise Exception("Element Not Found")
