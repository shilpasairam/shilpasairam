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

    # Reading Population data for PRISMAs Page
    def get_prisma_pop_data(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        pop = df.loc[df['Name'] == locatorname]['Prisma_Population'].dropna().to_list()
        return pop
    
    def get_prisma_excelfile_details(self, filepath, locatorname, column_name):
        df = pd.read_excel(filepath)
        sheet_name = os.getcwd()+(df.loc[df['Name'] == locatorname][column_name].to_list()[0])
        return sheet_name
    
    def get_prisma_data_details(self, filepath, locatorname):
        df = pd.read_excel(filepath)
        sheet_name = df.loc[df['Name'] == locatorname]['Study_Types'].to_list()
        sheet1 = df.loc[df['Name'] == locatorname]['OriginalStudiesNumbers'].to_list()
        sheet2 = df.loc[df['Name'] == locatorname]['RecordsNumber'].to_list()
        sheet3 = df.loc[df['Name'] == locatorname]['fullTextReviewRecordsNumber'].to_list()
        sheet4 = df.loc[df['Name'] == locatorname]['totalRecordsNumber'].to_list()
        sheet5 = df.loc[df['Name'] == locatorname]['Prisma_Image'].to_list()
        prisma_data = [(sheet_name[i], int(sheet1[i]), int(sheet2[i]), int(sheet3[i]), int(sheet4[i]), (os.getcwd()+sheet5[i])) for i in range(0, len(sheet_name))]
        return prisma_data

    def add_prisma_excel_file(self, locatorname, filepath):
        expected_excel_upload_status_text = "PRISMA successfully uploaded"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)
        
        self.refreshpage()
        time.sleep(3)
        try:
            pop_ele = self.select_element("prisma_pop_dropdown")
            select = Select(pop_ele)
            select.select_by_visible_text(pop_name[0])

            # Read the PRISMA file path required to upload
            prisma_excel = self.get_prisma_excelfile_details(filepath, locatorname, 'Prisma_Excel_File')

            self.input_text("prisma_excel_file", prisma_excel, UnivWaitFor=10)
            self.click("prisma_excel_upload_btn")
            time.sleep(3)
            actual_excel_upload_status_text = self.get_text("prisma_excel_status_text", UnivWaitFor=30)
            # time.sleep(2)

            # self.assertText(expected_excel_upload_status_text, actual_excel_upload_status_text)
            # self.LogScreenshot.fLogScreenshot(message=f"Excel File upload is success",
            #         pass_=True, log=True, screenshot=True)
            if actual_excel_upload_status_text == expected_excel_upload_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"PRISMA Excel File Upload is success.",
                                        pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading PRISMA Excel File.",
                                        pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message while uploading PRISMA Excel File.")
        except Exception:
            raise Exception("Unable to upload PRISMA Excel file")
    
    def del_prisma_excel_file(self, locatorname, excel_del_locator, excel_del_popup, filepath):
        expected_excel_del_status_text = "PRISMA excel file successfully deleted"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)
        
        self.refreshpage()
        time.sleep(3)
        try:
            pop_ele = self.select_element("prisma_pop_dropdown")
            select = Select(pop_ele)
            select.select_by_visible_text(pop_name[0])

            self.del_prisma_image(locatorname, filepath, "prisma_image_delete_btn", "prisma_image_delete_popup")
            time.sleep(2)

            self.click(excel_del_locator)
            time.sleep(2)
            self.click(excel_del_popup)
            time.sleep(2)

            actual_excel_del_status_text = self.get_text("prisma_excel_status_text", UnivWaitFor=30)
            # time.sleep(2)

            # self.assertText(expected_excel_del_status_text, actual_excel_del_status_text)
            # self.LogScreenshot.fLogScreenshot(message=f"Excel File Delete is success. Text is : {actual_excel_del_status_text}",
            #         pass_=True, log=True, screenshot=True)
            if actual_excel_del_status_text == expected_excel_del_status_text:
                self.LogScreenshot.fLogScreenshot(message=f"PRISMA Excel File Delete is success.",
                                        pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting PRISMA Excel File.",
                                        pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message while deleting PRISMA Excel File.")
            time.sleep(5)
        except Exception:
            raise Exception("Unable to delete PRISMA Excel file")

    def upload_prisma_image(self, locatorname, filepath, index):
        expected_image_upload_status_text = "PRISMA successfully updated"

        # Read the Data required to upload for study types
        stdy_data = self.get_prisma_data_details(filepath, locatorname)

        try:
            for i in stdy_data:
                stdy_ele = self.select_element("prisma_study_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(2)

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
                time.sleep(3)
                
                actual_image_upload_status_text = self.get_text("prisma_image_status_text", UnivWaitFor=30)
                # time.sleep(2)

                # self.assertText(expected_image_upload_status_text, actual_image_upload_status_text)
                # self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Upload is success",
                #     pass_=True, log=True, screenshot=True)

                if actual_image_upload_status_text == expected_image_upload_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Upload is success.",
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while uploading PRISMA Image File.",
                                            pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message while uploading PRISMA Image File.")
        except Exception:
            raise Exception("Unable to upload PRISMA image")
    
    def del_prisma_image(self, locatorname, filepath, del_locator, del_popup):
        expected_image_del_status_text = "PRISMA successfully deleted"

        # Read population details from data sheet
        pop_name = self.get_prisma_pop_data(filepath, locatorname)

        stdy_data = self.get_prisma_data_details(filepath, locatorname)

        try:
            for i in stdy_data:
                self.refreshpage()
                time.sleep(3)
                pop_ele = self.select_element("prisma_pop_dropdown")
                select = Select(pop_ele)
                select.select_by_visible_text(pop_name[0])
                time.sleep(1)

                stdy_ele = self.select_element("prisma_study_type_dropdown")
                select = Select(stdy_ele)
                select.select_by_visible_text(i[0])
                time.sleep(1)

                self.click(del_locator)
                time.sleep(2)
                self.click(del_popup)
                time.sleep(2)
                
                actual_image_del_status_text = self.get_text("prisma_image_status_text", UnivWaitFor=30)
                # time.sleep(2)

                # self.assertText(expected_image_del_status_text, actual_image_del_status_text)
                # self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Delete is success. Text is : {actual_image_del_status_text}",
                #     pass_=True, log=True, screenshot=True)

                if actual_image_del_status_text == expected_image_del_status_text:
                    self.LogScreenshot.fLogScreenshot(message=f"PRISMA Image File Delete is success.",
                                            pass_=True, log=True, screenshot=True)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Unable to find status message while deleting PRISMA Image File.",
                                            pass_=False, log=True, screenshot=True)
                    raise Exception("Unable to find status message while deleting PRISMA Image File.")
                time.sleep(5)
        except Exception:
            raise Exception("Unable to delete PRISMA image")
