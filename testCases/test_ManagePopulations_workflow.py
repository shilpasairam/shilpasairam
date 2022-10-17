# """
# Test will validate below pages functionality:

# 1. Manage Populations Page
# 2. Manage Updates Page
# 3. Manage QA Data Page

# """

# import time
# import pytest
# from datetime import date
# from Pages.ManagePopulationsPage import ManagePopulationsPage

# from Pages.LoginPage import LoginPage
# from Pages.ManageQADataPage import ManageQADataPage
# from Pages.ManageUpdatesPage import ManageUpdatesPage
# from Pages.OpenLiveSLRPage import LiveSLRPage
# from utilities.logScreenshot import cLogScreenshot
# from utilities.readProperties import ReadConfig


# @pytest.mark.usefixtures("init_driver")
# class Test_ManagePopultionsPage_Workflow:
#     baseURL = ReadConfig.getApplicationURL()
#     username = ReadConfig.getUserName()
#     password = ReadConfig.getPassword()
#     filepath = ReadConfig.getimportpublicationsdata()

#     def test_ManagePop_ManageUpdates_MangageQAData(self, extra):
#         # Instantiate the logScreenshot class
#         self.LogScreenshot = cLogScreenshot(self.driver, extra)
#         # Creating object of loginpage class
#         self.loginPage = LoginPage(self.driver, extra)
#         # Creating object of liveslrpage class
#         self.liveslrpage = LiveSLRPage(self.driver, extra)
#         # Creating object of ManagePopulationsPage class
#         self.mngpoppage = ManagePopulationsPage(self.driver, extra)
#         # Creating object of ManageUpdatesPage class
#         self.mngupdates = ManageUpdatesPage(self.driver, extra)
#         # Creating object of ManageQADataPage class
#         self.mngqadata = ManageQADataPage(self.driver, extra)
#         # Read extraction sheet values
#         self.file_upload = self.mngpoppage.get_template_file_details(self.filepath)
#         # Get StudyType and Files path to upload Managae QA Data
#         self.stdy_data = self.mngqadata.get_qa_file_details(self.filepath)

#         today = date.today()
#         self.dateval = today.strftime("%m/%d/%Y")  # .replace('/', '')
#         self.day_val = today.day
        
#         self.loginPage.driver.get(self.baseURL)
#         self.loginPage.complete_login(self.username, self.password, self.baseURL)
#         self.mngpoppage.go_to_managepopulations("managepopulations_button")

#         for i in self.file_upload:
#             try:
#                 added_pop = self.mngpoppage.add_population("add_population_btn", self.filepath, "template_file_upload", i[1], "manage_pop_table_rows")
#                 self.LogScreenshot.fLogScreenshot(message=f"Added population is {added_pop}", pass_=True, log=True, screenshot=False)
                
#                 manage_update_data = self.mngupdates.add_updates("manageupdates_button", "add_update_btn", added_pop, self.day_val, "manage_update_table_rows", self.dateval)
#                 self.LogScreenshot.fLogScreenshot(message=f"Added population udpate is {manage_update_data}", pass_=True, log=True, screenshot=False)
                
#                 self.mngqadata.add_manage_qa_data("manage_qa_data_button", self.stdy_data, self.filepath)

#                 self.mngqadata.del_manage_qa_data("manage_qa_data_button", self.stdy_data, "delete_file_button", "delete_file_popup", self.filepath)
                
#                 self.mngupdates.delete_manage_update("manageupdates_button", added_pop, "delete_updates", "delete_updates_popup", "manage_pop_table_rows")
#                 self.mngpoppage.delete_population("managepopulations_button", "delete_population", self.filepath, "delete_population_popup", "manage_pop_table_rows")

#             except Exception:
#                 self.LogScreenshot.fLogScreenshot(message=f"Error in accessing Manage publications page",
#                     pass_=False, log=True, screenshot=True)
#                 raise Exception("Element Not Found")
        