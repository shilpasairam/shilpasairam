import os
import time
import pandas as pd
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support.ui import Select


class PRISMASPage(Base):

    """Constructor of the ImportPublication Page class"""
    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    def go_to_prisma(self, locator, button):
        self.click(locator, UnivWaitFor=10)
        self.jsclick(button)
        time.sleep(5)

    def get_prisma_excelfile_details(self, filepath):
        file = pd.read_excel(filepath)
        sheet_name = list(os.getcwd()+file['Prisma_Excel_File'].dropna())
        return sheet_name
    
    def get_prisma_data_details(self, filepath):
        file = pd.read_excel(filepath)
        sheet_name = list(file['Study_Types'].dropna())
        sheet1 = list(file['OriginalStudiesNumbers'].dropna())
        sheet2 = list(file['RecordsNumber'].dropna())
        sheet3 = list(file['fullTextReviewRecordsNumber'].dropna())
        sheet4 = list(file['totalRecordsNumber'].dropna())
        sheet5 = list(os.getcwd()+file['Prisma_Image'].dropna())
        prisma_data = [(sheet_name[i], sheet1[i], sheet2[i], sheet3[i], sheet4[i], sheet5[i]) for i in range(0, len(sheet_name))]
        return prisma_data

    def add_prisma_excel_file(self, pop_index, prisma_excel):
        expected_excel_upload_status_text = "PRISMA successfully uploaded"
        
        self.refreshpage()
        time.sleep(3)
        try:
            pop_ele = self.select_element("prisma_pop_dropdown")
            select = Select(pop_ele)
            select.select_by_index(pop_index)

            self.input_text("prisma_excel_file", prisma_excel)
            self.click("prisma_excel_upload_btn")
            time.sleep(2)
            actual_excel_upload_status_text = self.get_text("prisma_excel_status_text", UnivWaitFor=10)

            self.assertText(expected_excel_upload_status_text, actual_excel_upload_status_text)
            self.LogScreenshot.fLogScreenshot(message=f"Excel File upload is success",
                    pass_=True, log=True, screenshot=True)
        except Exception:
            raise Exception("Unable to upload PRISMA Excel file")
    
    def del_prisma_excel_file(self, pop_index, excel_del_locator, excel_del_popup, stdy_data):
        expected_excel_del_status_text = "PRISMA excel file successfully deleted"
        
        self.refreshpage()
        time.sleep(3)
        try:
            pop_ele = self.select_element("prisma_pop_dropdown")
            select = Select(pop_ele)
            select.select_by_index(pop_index)

            self.del_prisma_image(pop_index, stdy_data, "prisma_image_delete_btn", "prisma_image_delete_popup")
            time.sleep(2)

            self.click(excel_del_locator)
            time.sleep(2)
            self.click(excel_del_popup)
            time.sleep(2)

            actual_excel_del_status_text = self.get_text("prisma_excel_status_text", UnivWaitFor=10)

            self.assertText(expected_excel_del_status_text, actual_excel_del_status_text)
            self.LogScreenshot.fLogScreenshot(message=f"Excel File Delete is success. Text is : {actual_excel_del_status_text}",
                    pass_=True, log=True, screenshot=True)
            time.sleep(5)
        except Exception:
            raise Exception("Unable to upload PRISMA Excel file")

    def upload_prisma_image(self, stdy_data, index):
        expected_image_upload_status_text = "PRISMA successfully updated"
        try:
            for i in stdy_data:
                stdy_ele = self.select_element("prisma_study_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.input_text("original_studies", i[1]+index, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("records_number", i[2]+index, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("fulltext_review", i[3]+index, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("total_record_number", i[4]+index, UnivWaitFor=5)
                time.sleep(1)
                self.input_text("prisma_image", i[5], UnivWaitFor=5)
                time.sleep(1)

                self.click("prisma_image_save_btn")
                time.sleep(2)
                
                actual_image_upload_status_text = self.get_text("prisma_image_status_text", UnivWaitFor=10)

                self.assertText(expected_image_upload_status_text, actual_image_upload_status_text)
                self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Upload is success",
                    pass_=True, log=True, screenshot=True)
        except Exception:
            raise Exception("Unable to upload PRISMA image")
    
    def del_prisma_image(self, pop_index, stdy_data, del_locator, del_popup):
        expected_image_del_status_text = "PRISMA successfully deleted"
        try:
            for i in stdy_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("prisma_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_index(pop_index)
                time.sleep(1)

                stdy_ele = self.select_element("prisma_study_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.click(del_locator)
                time.sleep(2)
                self.click(del_popup)
                time.sleep(2)
                
                actual_image_del_status_text = self.get_text("prisma_image_status_text", UnivWaitFor=10)

                self.assertText(expected_image_del_status_text, actual_image_del_status_text)
                self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Delete is success. Text is : {actual_image_del_status_text}",
                    pass_=True, log=True, screenshot=True)
                time.sleep(5)
        except Exception:
            raise Exception("Unable to upload PRISMA image")
