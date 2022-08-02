import os
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


class ManageQADataPage(Base):

    """Constructor of the ManageQAData Page class"""
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

    def go_to_manageqadata(self, locator):
        self.click(locator, UnivWaitFor=10)
        time.sleep(5)

    def get_qa_file_details(self, filepath):
        file = pd.read_excel(filepath)
        study_type = list(file['Study_Types'].dropna())
        qa_file = list(os.getcwd()+file['QA_Excel_Files'].dropna())
        result = [(study_type[i], qa_file[i]) for i in range(0, len(study_type))]
        return result
    
    def get_qa_file_details_override(self, filepath):
        file = pd.read_excel(filepath)
        study_type = list(file['Study_Types'].dropna())
        qa_file = list(os.getcwd()+file['Override_QA_Excel_Files'].dropna())
        result = [(study_type[i], qa_file[i]) for i in range(0, len(study_type))]
        return result

    def add_manage_qa_data(self, manage_qa_page, study_data, filepath):
        expected_upload_status_text = 'QA File successfully uploaded'
        
        self.click(manage_qa_page)
        # Read population details from data sheet
        new_pop_data, new_pop_val = self.mngpoppage.get_pop_data(filepath)
        try:
            for i in study_data:
                self.refreshpage()
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(new_pop_val[0])

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])

                self.input_text("qa_checklist_name", f"Auto_Test_{i[0]}")
                self.input_text("qa_checklist_citation", f"Auto_Test_Citation{i[0]}")
                self.input_text("qa_checklist_reference", f"Auto_Test_Reference{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(1)

                self.click("upload_save_button")
                time.sleep(1)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File upload is success. Text is {actual_upload_status_text}',
                                            pass_=True, log=True, screenshot=True)
                else:
                    raise Exception("Error in QA file uploading")
        except:
            raise Exception("Unable to upload the Manage QA Data")

    def del_manage_qa_data(self, manage_qa_page, study_data, del_locator, del_locator_popup, filepath):
        expected_delete_status_text = 'QA excel file successfully deleted'

        self.click(manage_qa_page)
        # Read population details from data sheet
        new_pop_data, new_pop_val = self.mngpoppage.get_pop_data(filepath)
        try:
            for i in study_data:
                self.refreshpage()
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(new_pop_val[0])

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])

                self.click(del_locator)
                time.sleep(1)
                self.click(del_locator_popup)
                time.sleep(1)

                actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)

                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File Deletion is success. Text is {actual_delete_status_text}',
                                            pass_=True, log=True, screenshot=True)
                else:
                    raise Exception("Error in QA file Deletion")
        
        except:
            raise Exception("Unable to delete the existing QA file")

    def add_multiple_manage_qa_data(self, study_data, pop_index):
        expected_upload_status_text = 'QA File successfully uploaded'
        
        # # Read population details from data sheet
        # new_pop_data, new_pop_val = self.mngpoppage.get_pop_data(filepath)
        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_index(pop_index)
                time.sleep(1)

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"Auto_Test_{i[0]}")
                self.input_text("qa_checklist_citation", f"Auto_Test_Citation{i[0]}")
                self.input_text("qa_checklist_reference", f"Auto_Test_Reference{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(2)

                self.click("upload_save_button")
                time.sleep(2)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File upload is success. Text is {actual_upload_status_text}',
                                            pass_=True, log=True, screenshot=True)
                else:
                    raise Exception("Error in QA file uploading")
        except:
            raise Exception("Unable to upload the Manage QA Data")

    def overwrite_multiple_manage_qa_data(self, study_data, pop_index):
        expected_upload_status_text = 'QA File successfully uploaded'
        
        # # Read population details from data sheet
        # new_pop_data, new_pop_val = self.mngpoppage.get_pop_data(filepath)
        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_index(pop_index)
                time.sleep(1)

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("qa_checklist_name", f"Auto_Test_{i[0]}")
                self.input_text("qa_checklist_citation", f"Auto_Test_Citation{i[0]}")
                self.input_text("qa_checklist_reference", f"Auto_Test_Reference{i[0]}")
                self.input_text("qa_excel_file_upload", i[1])
                time.sleep(2)

                self.click("upload_save_button")
                time.sleep(2)
                actual_upload_status_text = self.get_text("get_status_text", UnivWaitFor=10)

                if actual_upload_status_text == expected_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'Overwriting the existing QA File upload is success.',
                                            pass_=True, log=True, screenshot=True)
                else:
                    raise Exception("Error while Overwriting the QA file")
        except:
            raise Exception("Unable to overwrite the Manage QA Data")

    def del_multiple_manage_qa_data(self, study_data, del_locator, del_locator_popup, pop_index):
        expected_delete_status_text = 'QA excel file successfully deleted'

        # # Read population details from data sheet
        # new_pop_data, new_pop_val = self.mngpoppage.get_pop_data(filepath)
        try:
            for i in study_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("select_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_index(pop_index)
                time.sleep(1)

                stdy_ele = self.select_element("select_stdy_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.click(del_locator)
                time.sleep(1)
                self.click(del_locator_popup)
                time.sleep(1)

                actual_delete_status_text = self.get_text("get_status_text", UnivWaitFor=10)

                if actual_delete_status_text == expected_delete_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f'QA File Deletion is success. Text is {actual_delete_status_text}',
                                            pass_=True, log=True, screenshot=True)
                else:
                    raise Exception("Error in QA file Deletion")
        
        except:
            raise Exception("Unable to delete the existing QA file")