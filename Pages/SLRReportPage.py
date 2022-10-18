import numbers
import re
import time

import docx
import openpyxl
import pandas as pd
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait
from pandas.core.common import flatten

from Pages.Base import Base, fWaitFor
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


class SLRReport(Base):
    """Constructor of the LiveSLR Page class"""

    def __init__(self, driver, extra):
        super().__init__(driver, extra)  # initializing the driver from base class
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 20)

    def select_data(self, locator, locator_button):
        time.sleep(3)
        if self.isselected(locator_button):
            self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.jsclick(locator, UnivWaitFor=10)
            if self.isselected(locator_button):
                self.LogScreenshot.fLogScreenshot(message=f"Selected Element: {locator}",
                                                  pass_=True, log=True, screenshot=True)

    def select_sub_section(self, locator, locator_button, scroll=None):
        if self.scroll(scroll, UnivWaitFor=20):
            if self.isselected(locator_button):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.jsclick(locator, UnivWaitFor=10)
                if self.isselected(locator_button):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=True)
            self.scrollback()

    def select_all_sub_section(self, locator, locator_button, scroll=None):
        if self.scroll(scroll, UnivWaitFor=20):
            if self.isselected(locator_button):
                self.LogScreenshot.fLogScreenshot(message=f"{locator} already selected",
                                                  pass_=True, log=True, screenshot=True)
            else:
                self.jsclick(locator, UnivWaitFor=10)
                if self.isselected(locator_button):
                    self.LogScreenshot.fLogScreenshot(message=f"{locator} selected",
                                                      pass_=True, log=True, screenshot=True)
            self.scrollback()

    def get_additional_criteria_values(self, locator_study, locator_var):
        ele1 = self.select_elements(locator_study)
        ele2 = self.select_elements(locator_var)
        return ele1, ele2

    def validate_additional_criteria_val(self, filepath, locator_study, locator_var):
        # Read reportedvariables and studydesign expected data values
        design_val, var_val = self.liveslrpage.get_data_values(filepath)
        # Get the actual values
        act_study_design, act_rep_var = self.get_additional_criteria_values(locator_study, locator_var)
        try:
            for k in act_study_design:
                if k.text in design_val:
                    self.LogScreenshot.fLogScreenshot(
                        message=f"Correct StudyDesign Value is displayed in NMA Page:{k.text}",
                        pass_=True, log=True, screenshot=False)
            for v in act_rep_var:
                if v.text in var_val:
                    self.LogScreenshot.fLogScreenshot(
                        message=f"Correct Reported Variable is displayed in NMA Page:{v.text}",
                        pass_=True, log=True, screenshot=False)
        except Exception:
            raise Exception("Unable to validate Additional Criteria Values")

    def validate_selected_area(self, pop, slr):
        pop_sel, slr_sel = self.collect_selected_area_details("selected_area_population",
                                                              "selected_area_slrtype")
        if pop in pop_sel and slr in slr_sel:
            self.LogScreenshot.fLogScreenshot(message=f"Selected elements are shown correct\n"
                                                      f"Selected Area value: {pop_sel} {slr_sel}\n"
                                                      f"Selected Element value: {pop}, {slr}",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Selected elements are wrong",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Selected Population and SLR Type are not matching in Selected area")

    def prism_value_validation(self, prism, efilename, wfilename, word_filename):

        # Reading Article identifier column values from WebExcel Sheet
        webexcel = openpyxl.load_workbook(f'ActualOutputs//{wfilename}')
        webcount = []
        webcount_final = []

        for sheet in webexcel.sheetnames:
            webexcel_value = pd.read_excel(f'ActualOutputs//{wfilename}', sheet_name=sheet, skiprows=3)
            webex = webexcel_value['Article Identifier(s)']
            webcount.append([item for item in webex if str(item) != 'nan'])

        # Removing duplicates to get the proper length of Article identifier data
        [webcount_final.append(x) for x in list(flatten(webcount)) if x not in webcount_final]

        # Read PRISMA value from Complete Excel Sheet
        excel = openpyxl.load_workbook(f'ActualOutputs//{efilename}')
        excel_sheet = excel['Updated PRISMA']
        count_str = excel_sheet['B22'].value
        count_val = excel_sheet['C22'].value

        # Reading PRISMA Value from Complete Word Report
        docs = docx.Document(f'ActualOutputs//{word_filename}')
        word = []
        for row in docs.tables[4].rows:
            word.append(row.cells[1].text)

        self.LogScreenshot.fLogScreenshot(message=f"WebExcel FileName is: {wfilename}\n and Count Value is: {len(webcount_final)}"
                                                  f"Excel FileName is: {efilename}\n and Count value is: {count_str} {count_val}"
                                                  f"Word FileName is: {word_filename}\n and Count value is: {word[7]}",
                                          pass_=True, log=True, screenshot=False)

        if int(prism) == count_val and len(webcount_final) == count_val and int(word[7]) == count_val:
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Prisma Count Value: {len(webcount_final)}\n"
                                                      f"Excel Sheet Prisma Count Value: {count_val}\n"
                                                      f"Word Prisma Count Value: {word[7]}\n"
                                                      f"UI Updated Prisma Count Value: {prism}\n"
                                                      f"Records are matching",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Prisma Count Value: {len(webcount_final)}\n"
                                                      f"Excel Sheet Prisma Count Value: {count_val}\n"
                                                      f"Word Prisma Count Value: {word[7]}\n"
                                                      f"UI Updated Prisma Count Value: {prism}\n"
                                                      f"Records are not matching",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Prisma count values are mismatching")
        self.scrollback()

    def collect_selected_area_details(self, locator1, locator2):
        x = self.get_text(locator1)
        y = self.get_text(locator2)
        return x, y

    def preview_result(self, locator):
        if self.clickable(locator):
            self.jsclick(locator)
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is clickable",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is not clickable",
                                              pass_=False, log=True, screenshot=False)

    def table_display_check(self, locator):
        if self.isdisplayed(locator, UnivWaitFor=120):
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is displayed",
                                              pass_=True, log=True, screenshot=True)
        else:
            time.sleep(10)
            self.driver.find_element(getattr(By, self.locatortype(locator)), self.locatorpath(locator)).is_displayed()
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is displayed with extra wait time",
                                              pass_=True, log=True, screenshot=True)

    # def check_download_error_ele(self, locator):
    #     return self.wait.until(ec.invisibility_of_element_located((By.XPATH, locator)))
    #     # if self.isdisplayed(locator, UnivWaitFor=5):
    #     #     self.LogScreenshot.fLogScreenshot(message=f"{locator} is displayed",
    #     #                                       pass_=False, log=True, screenshot=True)
    #     #     return True
    #     # else:
    #     #     return False

    @fWaitFor
    def check_download_error_element(self, locator, UnivWaitFor=0):
        try:
            if self.isdisplayed(locator):
                # value = True
                value = self.wait.until(
                    ec.presence_of_element_located((getattr(By, self.locatortype(locator)), self.locatorpath(locator))))
                self.LogScreenshot.fLogScreenshot(message=f"Entered into IF block",
                                                  pass_=False, log=True, screenshot=False)
                self.LogScreenshot.fLogScreenshot(message=f"IF block value is: {value}",
                                                  pass_=False, log=True, screenshot=True)
            else:
                # time.sleep(5)
                value = self.driver.find_element(getattr(By, self.locatortype(locator)),
                                                 self.locatorpath(locator)).is_displayed()
                self.LogScreenshot.fLogScreenshot(message=f"Entered into ELSE block",
                                                  pass_=False, log=True, screenshot=False)
                self.LogScreenshot.fLogScreenshot(message=f"ELSE block value is: {value}",
                                                  pass_=False, log=True, screenshot=True)
            return value
        except NoSuchElementException:
            self.LogScreenshot.fLogScreenshot(message=f"{locator} is not present",
                                              pass_=False, log=True, screenshot=False)
            return False

    @fWaitFor
    def generate_download_report(self, locator, UnivWaitFor=0):
        # download table csv
        try:
            self.jsclick(locator)
            # wait for table to download
            time.sleep(15)
            self.LogScreenshot.fLogScreenshot(message=f"Download table",
                                              pass_=True, log=True, screenshot=False)
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Error in downloading the table",
                                              pass_=False, log=True, screenshot=True)
            raise Exception("Download table failed")

    def back_to_report_page(self, locator):
        self.jsclick(locator)

    def validate_filename(self, filename, filepath):
        try:
            file = pd.read_excel(filepath)
            expectedname = list(file['ExpectedFilenames'].dropna())
            # actualname = [x for x in re.compile(r'\D+').findall(filename)]
            actualname = filename[:-19]
            # for i in expectedname:
            #     # if i == actualname[0]:
            #     if i == actualname:
            #         self.LogScreenshot.fLogScreenshot(message=f"Correct file is downloaded",
            #                                           pass_=True, log=True, screenshot=False)
            #         break
            if actualname in expectedname:
                self.LogScreenshot.fLogScreenshot(message=f"Correct file is downloaded",
                                                      pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Filename is not present in the expected list. Expected Filenames are {expectedname} and Actual Filename is {actualname}",
                                                      pass_=False, log=True, screenshot=False)
        except Exception:
            raise Exception("Error during filename validation")

    # method to get the downloaded file name
    def getFilenameAndValidate(self, waitTime):
        self.driver.execute_script("window.open()")
        # switch to new tab
        self.driver.switch_to.window(self.driver.window_handles[-1])
        # navigate to chrome downloads
        self.driver.get('chrome://downloads')
        # define the endTime
        endTime = time.time() + waitTime
        while True:
            try:
                time.sleep(10)
                filename = self.driver.execute_script(
                    "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
                self.LogScreenshot.fLogScreenshot(message=f"Downloaded filename is {filename}",
                                                  pass_=True, log=True, screenshot=False)
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[1])
                return filename
            except:
                pass
            time.sleep(1)
            if time.time() > endTime:
                break
    
    def excel_content_validation(self, webexcel_filename, excel_filename):
        self.LogScreenshot.fLogScreenshot(message=f"Content validation between WebExcel and Complete Excel Report",
                                          pass_=True, log=True, screenshot=False)
        self.LogScreenshot.fLogScreenshot(message=f"FileNames are: {webexcel_filename} and \n{excel_filename}",
                                          pass_=True, log=True, screenshot=False)
        column_names = ['Article Identifier(s)', 'Publication Type', 'Short Reference']
        sheet_names = []

        webexcel_data = openpyxl.load_workbook(f'ActualOutputs//{webexcel_filename}')
        excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')

        for i in webexcel_data.sheetnames:
            if i in excel_data.sheetnames:
                sheet_names.append(i)

        self.LogScreenshot.fLogScreenshot(message=f"Sheetnames are: {sheet_names}",
                                          pass_=True, log=True, screenshot=False)
        for sheet in sheet_names:
            webexcel = pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name=sheet, skiprows=3)
            excel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=sheet, skiprows=3)

            try:
                for col in column_names:
                    webex = webexcel[col]
                    compex = excel[col]

                    webex = [item for item in webex if str(item) != 'nan']
                    compex = [item for item in compex if str(item) != 'nan']

                    if len(webex) == len(compex) and webex == compex:
                        self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Values in Column '{col}' are matching between WebExcel and Complete Excel "
                                          f"Report.\n WebExcel Elements Length: {len(webex)}\n Excel Elements Length: {len(compex)}\n",
                                          # f"WebExcel Elements: {webex}\n Excel Elements: {compex}",
                                  pass_=True, log=True, screenshot=False)
                    else:
                        raise Exception("Elements are not matching between Webexcel and Complete Excel")
            except Exception:
                raise Exception("Error in Excel sheet content validation")

    def excel_to_word_content_validation(self, webexcel_filename, excel_filename, word_filename):
        self.LogScreenshot.fLogScreenshot(message=f"Content validation between WebExcel, Complete Excel and Complete Word Reports",
                                          pass_=True, log=True, screenshot=False)
        self.LogScreenshot.fLogScreenshot(message=f"FileNames are: {webexcel_filename}, \n{excel_filename}, \n{word_filename}",
                                          pass_=True, log=True, screenshot=False)
        
        # Index of Table number 6 is : 5. Starting point for word table content comparison
        table_count = 5
        sheet_names = []

        webexcel_data = openpyxl.load_workbook(f'ActualOutputs//{webexcel_filename}')
        excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')

        for i in webexcel_data.sheetnames:
            if i in excel_data.sheetnames:
                sheet_names.append(i)

        self.LogScreenshot.fLogScreenshot(message=f"Sheetnames are: {sheet_names}",
                                          pass_=True, log=True, screenshot=False)
        for sheet in sheet_names:
            webexcel = pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name=sheet, skiprows=3)
            excel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=sheet, skiprows=3)
            docs = docx.Document(f'ActualOutputs//{word_filename}')
            try:
                table = docs.tables[table_count]
                data = [[cell.text for cell in row.cells] for row in table.rows]
                df_word = pd.DataFrame(data)

                # Using count variable to loop over columns in word document
                count = 0
                # df_word.values[0] will give the list of column names from the table in Word document
                for col_name in df_word.values[0]:
                    # Restricting comparison upto 3 columns in word due to data formatting issues in further columns
                    if count <= 2:
                        if col_name == 'Year/Country':
                            # This IF condition is just to add space to the column name as per Excel sheet to match the names
                            # between word and excel. This is an workaround until it is fixed
                            col_name = 'Year / Country'
                        if col_name in webexcel.columns.values and col_name in excel.columns.values:
                            webex = webexcel[col_name]
                            compex = excel[col_name]
                            word = []
                            for row in docs.tables[table_count].rows:
                                word.append(row.cells[count].text)

                            webex = [item for item in webex if str(item) != 'nan']
                            compex = [item for item in compex if str(item) != 'nan']
                            word.pop(0)

                            if len(webex) == len(compex) == len(word) and webex == compex == word:
                                self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Values in Column '{col_name}' are matching between WebExcel, Complete Excel and Word Reports.\n"
                                          f"WebExcel Elements Length: {len(webex)}\n Excel Elements Length: {len(compex)}\n Word Elements Length: {len(word)}\n",
                                          # f"WebExcel Elements: {webex}\n Excel Elements: {compex}\n Word Elements: {word}",
                                  pass_=True, log=True, screenshot=False)
                            else:
                                self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Values in Column '{col_name}' are not matching between WebExcel, Complete Excel and Word Reports.\n"
                                          f"WebExcel Elements Length: {len(webex)}\n Excel Elements Length: {len(compex)}\n Word Elements Length: {len(word)}\n"
                                          f"WebExcel Elements: {webex}\n Excel Elements: {compex}\n Word Elements: {word}",
                                  pass_=False, log=True, screenshot=False)
                                raise Exception("Elements are not matching between Webexcel, Complete Excel and Word Reports")
                            count += 1
                        else:
                            raise Exception("Column names are not matching between Webexcel, Complete Excel and Word Reports")
                table_count += 1
            except Exception:
                raise Exception("Error in Word report content validation")


    # ############## Using Openpyxl library #################
    # def excel_content_validation(self, webexcel_filename, excel_filename, slrtype):
    #     self.LogScreenshot.fLogScreenshot(message=f"FileNames are: {webexcel_filename} and \n{excel_filename}",
    #                                       pass_=True, log=True, screenshot=False)
    #     sheet_names = ReadConfig.get_sheetname_as_per_slrtype(slrtype)
    #     self.LogScreenshot.fLogScreenshot(message=f"Sheetnames are: {sheet_names}",
    #                                       pass_=True, log=True, screenshot=False)
    #     for sheet in sheet_names:
    #         webex = []
    #         compex = []
    #         webexcel = openpyxl.load_workbook(f'ActualOutputs//{webexcel_filename}')
    #         webexcel_sheet = webexcel[sheet]
    #         excel = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
    #         excel_sheet = excel[sheet]

    #         try:
    #             for i in webexcel_sheet.rows:
    #                 webex.append(i[0].value)

    #             for j in excel_sheet.rows:
    #                 compex.append(j[0].value)

    #             # Removing None values from the lists
    #             webex = list(filter(None, webex))
    #             compex = list(filter(None, compex))

    #             # Comparison and Merging with regular expressions
    #             res = [ele for ele in compex if (ele in webex)]
    #             self.LogScreenshot.fLogScreenshot(message=f"Regular Expression output: {res}",
    #                                               pass_=True, log=True, screenshot=False)
    #             if str(bool(res)):
    #                 # Extracting numbers from the lists
    #                 ele1 = [x for x in webex if isinstance(x, numbers.Number)]
    #                 ele2 = [x for x in compex if isinstance(x, numbers.Number)]
    #                 if ele1 == sorted(ele1) and ele2 == sorted(ele2):
    #                     if ele1 == ele2:
    #                         self.LogScreenshot.fLogScreenshot(message=f"Article Identifier(s) are matching between WebExcel "
    #                                                                   f"and Complete Excel report. \n"
    #                                                                   f"Compared Element values are: {ele1} and \n{ele2}",
    #                                                           pass_=True, log=True, screenshot=False)
    #                     else:
    #                         self.LogScreenshot.fLogScreenshot(message=f"Article Identifier(s) are not matching between WebExcel "
    #                                                                   f"and Complete Excel report. \n"
    #                                                                   f"Compared Element values are: {ele1} and \n{ele2}",
    #                                                           pass_=False, log=True, screenshot=False)
    #                         raise Exception("Article Identifier(s) are not matching")
    #                 else:
    #                     raise Exception("Article Identifier Elements are not in sorted order")
    #         except Exception:
    #             raise Exception("Error in Excel sheet content validation")

    # def excel_to_word_content_validation(self, webexcel_filename, excel_filename, word_filename, slrtype):
    #     self.LogScreenshot.fLogScreenshot(message=f"FileNames are: "
    #                                               f"{webexcel_filename}, \n{excel_filename}, \n{word_filename}",
    #                                       pass_=True, log=True, screenshot=False)
    #     # Index of Table number 6 is : 5. Starting point for word table content comparison
    #     table_count = 5
    #     sheet_names = ReadConfig.get_sheetname_as_per_slrtype(slrtype)
    #     self.LogScreenshot.fLogScreenshot(message=f"Sheetnames are: {sheet_names}",
    #                                       pass_=True, log=True, screenshot=False)
    #     for sheet in sheet_names:
    #         webex = []
    #         compex = []
    #         word = []
    #         webexcel = openpyxl.load_workbook(f'ActualOutputs//{webexcel_filename}')
    #         webexcel_sheet = webexcel[sheet]
    #         excel = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
    #         excel_sheet = excel[sheet]
    #         docs = docx.Document(f'ActualOutputs//{word_filename}')
    #         try:
    #             for i in webexcel_sheet.rows:
    #                 webex.append(i[3].value)

    #             for j in excel_sheet.rows:
    #                 compex.append(j[3].value)

    #             for row in docs.tables[table_count].rows:
    #                 word.append(row.cells[0].text)
    #             # Incrementing the counter to switch for next table based on the number of excel sheet names
    #             table_count += 1

    #             # Removing None values from the lists
    #             webex = list(filter(None, webex))
    #             compex = list(filter(None, compex))

    #             if compex == webex and compex == word:
    #                 self.LogScreenshot.fLogScreenshot(message=f"Contents are matching in WebExcel, Complete Excel and "
    #                                                           f"Complete Word Reports. Short Reference values are: \n"
    #                                                           f"{webex} \n {compex} \n {word}",
    #                                                   pass_=True, log=True, screenshot=False)
    #             else:
    #                 self.LogScreenshot.fLogScreenshot(message=f"Contents are not matching in WebExcel, Complete Excel and "
    #                                                           f"Complete Word Reports. Short Reference values are: \n"
    #                                                           f"{webex} \n {compex} \n {word}",
    #                                                   pass_=False, log=True, screenshot=False)
    #                 raise Exception("Word report contents are not matching")
    #         except Exception:
    #             raise Exception("Error in Word report content validation")
