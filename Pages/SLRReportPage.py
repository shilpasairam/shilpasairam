import ast
import numbers
import os
import re
import time

import docx
import openpyxl
import pandas as pd
from pathlib import Path
from re import search
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from pandas.core.common import flatten

from Pages.Base import Base, fWaitFor
from Pages.ExtendedBasePage import ExtendedBase
from Pages.ImportPublicationsPage import ImportPublicationPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig


class SLRReport(Base):

    """Constructor of the SLRReport class"""
    def __init__(self, driver, extra):
        # initializing the driver from base class
        super().__init__(driver, extra)  
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of liveslrpage class
        self.liveslrpage = LiveSLRPage(self.driver, self.extra)
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Creating object of ExtendedBase class
        self.exbase = ExtendedBase(self.driver, extra)  
        # Creating object of ImportPublicationPage class
        self.imppubpage = ImportPublicationPage(self.driver, extra)                     
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
            self.scrollback("SLR_page_header")

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
            self.scrollback("SLR_page_header")

    def get_additional_criteria_values(self, locator_study, locator_var):
        ele1 = self.select_elements(locator_study)
        ele2 = self.select_elements(locator_var)
        return ele1, ele2

    def get_source_template(self, filepath, col_name):
        file = pd.read_excel(filepath)
        expectedfilepath = list(os.getcwd()+file[col_name].dropna())
        return expectedfilepath
    
    def get_duplicates_from_list(self, input_list):
        return list(set([x for x in input_list if input_list.count(x) > 1]))

    def list_comparison_between_reports_data(self, source_list, compex_list, webex_list=None, word=None):
        idx = 0
        res_index = []
        res_list = []
        if webex_list is not None and word is not None:
            for i in source_list:
                if i != compex_list[idx] and i != webex_list[idx] and i != word[idx]:
                    res_index.append(idx)
                idx = idx + 1
        elif webex_list is None and word is not None:
            for i in source_list:
                if i != compex_list[idx] and i != word[idx]:
                    res_index.append(idx)
                idx = idx + 1
        elif webex_list is not None and word is None:
            for i in source_list:
                if i != compex_list[idx] and i != webex_list[idx]:
                    res_index.append(idx)
                idx = idx + 1
        else:
            for i in source_list:
                if i != compex_list[idx]:
                    res_index.append(idx)
                idx = idx + 1

        if webex_list is not None and word is not None:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])
                res_list[index].append(webex_list[n])
                res_list[index].append(word[n])
        elif webex_list is None and word is not None:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])
                res_list[index].append(word[n])
        elif webex_list is not None and word is None:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])
                res_list[index].append(webex_list[n])
        else:
            for index, n in enumerate(res_index):
                res_list.append([source_list[n]])
                res_list[index].append(compex_list[n])

        return res_list
    
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

        # Reading Study Identifier column values from WebExcel Sheet
        webexcel = openpyxl.load_workbook(f'ActualOutputs//{wfilename}')
        webcount = []
        webcount_final = []

        for sheet in webexcel.sheetnames:
            webexcel_value = pd.read_excel(f'ActualOutputs//{wfilename}', sheet_name=sheet, skiprows=3)
            webex = webexcel_value['Study Identifier']
            webcount.append([item for item in webex if str(item) != 'nan'])

        # Removing duplicates to get the proper length of Study Identifier data
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

        self.LogScreenshot.fLogScreenshot(message=f"WebExcel FileName is: {wfilename} and Count Value is: "
                                                  f"{len(webcount_final)} \nExcel FileName is: {efilename} and "
                                                  f"Count value is: {count_str} {count_val}"
                                                  f"Word FileName is: {word_filename}\n and Count value is: {word[8]}",
                                          pass_=True, log=True, screenshot=False)

        if int(prism) == count_val and len(webcount_final) == count_val and int(word[8]) == count_val:
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Prisma Count Value: {len(webcount_final)}\n"
                                                      f"Excel Sheet Prisma Count Value: {count_val}\n"
                                                      f"Word Prisma Count Value: {word[8]}\n"
                                                      f"UI Updated Prisma Count Value: {prism}\n"
                                                      f"Records are matching",
                                              pass_=True, log=True, screenshot=True)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Prisma Count Value: {len(webcount_final)}\n"
                                                      f"Excel Sheet Prisma Count Value: {count_val}\n"
                                                      f"Word Prisma Count Value: {word[8]}\n"
                                                      f"UI Updated Prisma Count Value: {prism}\n"
                                                      f"Records are not matching",
                                              pass_=False, log=True, screenshot=False)
            raise Exception("Prisma count values are mismatching")
        self.scrollback("SLR_page_header")

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
        except Exception:
            self.LogScreenshot.fLogScreenshot(message=f"Filename is not present in the expected list. Expected "
                                                      f"Filenames are {expectedname} and Actual "
                                                      f"Filename is {actualname}",
                                                      pass_=False, log=True, screenshot=False)
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
                    "return document.querySelector('downloads-manager').shadowRoot.querySelector"
                    "('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
                self.LogScreenshot.fLogScreenshot(message=f"Downloaded filename is {filename}",
                                                  pass_=True, log=True, screenshot=False)
                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[1])
                return filename
            except Exception:
                pass
            time.sleep(1)
            if time.time() > endTime:
                break
    
    def excel_content_validation(self, filepath, index, webexcel_filename, excel_filename):
        self.LogScreenshot.fLogScreenshot(message=f"Content validation between Source File and downloaded WebExcel "
                                                  f"and Complete Excel Reports",
                                          pass_=True, log=True, screenshot=False)
        
        source_template = self.get_source_template(filepath, 'ExpectedSourceTemplateFile_Excel')

        self.LogScreenshot.fLogScreenshot(message=f"Source Filename is: {Path(f'{source_template[0]}').stem}, "
                                                  f"Downloaded FileNames are: {webexcel_filename} "
                                                  f"and \n{excel_filename}",
                                          pass_=True, log=True, screenshot=False)

        '''Reference-LIVEHTA1161: Check position of 'Back to TOC' button and presence of column names at row 4'''
        wb = openpyxl.load_workbook(f'ActualOutputs//{webexcel_filename}')
        for i in wb.sheetnames:
            # Check the position of 'Back to TOC' button
            ex = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
            ex_sheet = ex[i]
            if ex_sheet['H1'].value == 'Back To Toc':
                self.LogScreenshot.fLogScreenshot(message=f"'Back To Toc' option is present in '{i}' sheet",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"'Back To Toc' option is not present in '{i}' "
                                                          f"sheet",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"'Back To Toc' option is not present in '{i}' sheet")
            
            # Check the presence of column names at row 4
            df = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=i, skiprows=3)
            if 'Study Identifier' in df.columns.values:
                self.LogScreenshot.fLogScreenshot(message=f"Column names are present at row 4.",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Column names are not present at row 4",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Column names are not present at row 4")           

        # Actual excel content validation step starts here
        source_data = openpyxl.load_workbook(f'{source_template[0]}')

        self.LogScreenshot.fLogScreenshot(message=f"Source file Sheetnames are: {source_data.sheetnames}",
                                          pass_=True, log=True, screenshot=False)
        
        expected_data = pd.read_excel(f'{source_template[0]}', sheet_name=source_data.sheetnames[index])
        webexcel = pd.concat(pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name=None, skiprows=3),
                             ignore_index=True)
        excel = pd.concat(pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=None, skiprows=3),
                          ignore_index=True)

        try:
            # Check the length of 1st column from the report to make sure number of rows are as expected
            source_len = expected_data["Study Identifier"]
            webex_len = webexcel["Study Identifier"]
            compex_len = excel["Study Identifier"]

            source_len = [item for item in source_len if str(item) != 'nan']
            webex_len = [item for item in webex_len if str(item) != 'nan']
            compex_len = [item for item in compex_len if str(item) != 'nan']

            if len(source_len) == len(webex_len) == len(compex_len):
                self.LogScreenshot.fLogScreenshot(message=f"Elements length is matching between Source Excel, "
                                                          f"Web_Excel and Complete Excel Report. "
                                                          f"Source Elements Length: {len(source_len)}\n "
                                                          f"WebExcel Elements Length: {len(webex_len)}\n "
                                                          f"Excel Elements Length: {len(compex_len)}\n",
                                                  pass_=True, log=True, screenshot=False)

                # Content validation starts from here
                for col in expected_data.columns.values:
                    source = expected_data[col]
                    webex = webexcel[col]
                    compex = excel[col]

                    source = [item for item in source if str(item) != 'nan']
                    webex = [item for item in webex if str(item) != 'nan']
                    compex = [item for item in compex if str(item) != 'nan']

                    comparison_result = self.list_comparison_between_reports_data(source, compex, webex_list=webex)

                    if len(comparison_result) == 0:
                        self.LogScreenshot.fLogScreenshot(message=f"From '{source_data.sheetnames[index]}' Report, "
                                                                  f"Values in Column '{col}' are matching between "
                                                                  f"Source, WebExcel and Complete Excel Report.\n",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"From '{source_data.sheetnames[index]}' Report, "
                                                                  f"Values in Column '{col}' are not matching between "
                                                                  f"Source, WebExcel and Complete Excel Report.\n"
                                                                  f"Mismatch values are arranged in following "
                                                                  f"order -> Source File, WebExcel, Complete Excel. "
                                                                  f"{comparison_result}",
                                                          pass_=False, log=True, screenshot=False)
                        raise Exception("Elements are not matching between Webexcel and Complete Excel")
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Elements length is not matching between Source Excel, "
                                                          f"Web_Excel and Complete Excel Report. "
                                                          f"Source Elements Length: {len(source_len)}\n "
                                                          f"WebExcel Elements Length: {len(webex_len)}\n "
                                                          f"Excel Elements Length: {len(compex_len)}\n",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"Elements length is not matching between Source Excel, Web_Excel and "
                                f"Complete Excel Report.")
        except Exception:
            raise Exception("Error in Excel sheet content validation")

    def word_content_validation(self, filepath, index, word_filename):
        self.LogScreenshot.fLogScreenshot(message=f"Content validation between Source File and "
                                                  f"Complete Word Report",
                                          pass_=True, log=True, screenshot=False)

        source_template = self.get_source_template(filepath, 'ExpectedSourceTemplateFile_Word')

        self.LogScreenshot.fLogScreenshot(message=f"FileName is: {word_filename}",
                                          pass_=True, log=True, screenshot=False)

        source_excel = openpyxl.load_workbook(f'{source_template[index]}')                                          
        
        # Index of Table number 6 is : 5. Starting point for word table content comparison
        table_count = 5

        self.LogScreenshot.fLogScreenshot(message=f"Sheetnames are: {source_excel.sheetnames}",
                                          pass_=True, log=True, screenshot=False)
        for sheet in source_excel.sheetnames:
            src_data = pd.read_excel(f'{source_template[index]}', sheet_name=sheet)
            docs = docx.Document(f'ActualOutputs//{word_filename}')
            try:
                table = docs.tables[table_count]
                data = [[cell.text for cell in row.cells] for row in table.rows]
                df_word = pd.DataFrame(data)

                # Check the length of 1st column from the report to make sure number of rows are as expected
                src_data_len = src_data[df_word.values[0][0]]
                word_len = []
                for row in docs.tables[table_count].rows:
                    word_len.append(row.cells[0].text)

                src_data_len = [item for item in src_data_len if str(item) != 'nan']
                word_len.pop(0)

                if len(src_data_len) == len(word_len):
                    self.LogScreenshot.fLogScreenshot(message=f"Elements length is matching between Source Excel and "
                                                              f"Complete Word Report. Source Excel Elements Length: "
                                                              f"{len(src_data_len)}\n Word Elements Length: "
                                                              f"{len(word_len)}\n",
                                                      pass_=True, log=True, screenshot=False)

                    # Content validation starts from here
                    # Using count variable to loop over columns in word document
                    count = 0
                    # df_word.values[0] will give the list of column names from the table in Word document
                    for col_name in df_word.values[0]:
                        # # Restricting comparison upto 3 columns in word due to data formatting
                        # # issues in further columns
                        # if count <= 2:
                        #     if col_name == 'Year/Country':
                        #         # This IF condition is just to add space to the column name as per Excel sheet to
                        #         # match the names between word and excel. This is an workaround until it is fixed
                        #         col_name = 'Year / Country'
                        if col_name in src_data.columns.values:
                            src_data_final = src_data[col_name]
                            word = []
                            for row in docs.tables[table_count].rows:
                                word.append(row.cells[count].text)

                            src_data_final = [item for item in src_data_final if str(item) != 'nan']
                            # Converting Integer list to String list
                            src_data_final = [str(x) for x in src_data_final]                            
                            word.pop(0)

                            comparison_result = self.list_comparison_between_reports_data(src_data_final, word)

                            if len(comparison_result) == 0:
                                self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Values in "
                                                                          f"Column '{col_name}' are matching "
                                                                          f"between Source Excel and Word Report.\n",
                                                                  pass_=True, log=True, screenshot=False)
                            else:
                                self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Values in "
                                                                          f"Column '{col_name}' are not matching "
                                                                          f"between Source Excel and Word Report.\n "
                                                                          f"Mismatch values are arranged in "
                                                                          f"following order -> Word, Complete Excel "
                                                                          f"and WebExcel Report. {comparison_result}",
                                                                  pass_=False, log=True, screenshot=False)
                                raise Exception("Elements are not matching between Source Excel "
                                                "and Word Reports")
                            count += 1
                        else:
                            raise Exception("Column names are not matching between Source Excel "
                                            "and Word Reports")
                    table_count += 1
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Elements length is not matching between Source Excel "
                                                              f"and Complete Word Report. Source Excel Elements "
                                                              f"Length: {len(src_data_len)}\n Word Elements "
                                                              f"Length: {len(word_len)}\n",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Elements length is not matching between Source Excel and "
                                    f"Complete Word Report.")
            except Exception:
                raise Exception("Error in Word report content validation")

    def presencof_publicationtype_col_in_wordreport(self, webexcel_filename, excel_filename, word_filename):
        self.LogScreenshot.fLogScreenshot(message=f"Check presence of Publication Type column in Word Report",
                                          pass_=True, log=True, screenshot=False)
        self.LogScreenshot.fLogScreenshot(message=f"FileNames are: {webexcel_filename}, \n{excel_filename}, "
                                                  f"\n{word_filename}",
                                          pass_=True, log=True, screenshot=False)
        
        # Index of Table number 6 is : 5. Starting point for word table content comparison
        table_counts = {5: "Table 2-2 Clinical study characteristics",
                        6: "Table 2-3 Summary of patient demographics and baseline characteristics",
                        7: "Table 2-4 Efficacy reported in clinical studies",
                        8: "Table 2-5 Safety reported in clinical studies"}
        sheet = 'Clinical Report'

        for table_count, value in table_counts.items():
            webexcel = pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name=sheet, skiprows=3)
            excel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=sheet, skiprows=3)
            docs = docx.Document(f'ActualOutputs//{word_filename}')
            try:
                table = docs.tables[table_count]
                data = [[cell.text for cell in row.cells] for row in table.rows]
                df_word = pd.DataFrame(data)

                # Using count variable 1 will give the details of 'Publication Type' column in word document
                count = 1
                # 'Publication Type' is the column which will be used for validation
                col_name = 'Publication Type'

                # df_word.values[0] will give the list of column names from the table in Word document
                if col_name in df_word.values[0]:
                    self.LogScreenshot.fLogScreenshot(message=f"Column '{col_name}' is present after 'Short Reference'"
                                                              f" in Word report -> '{value}'",
                                                      pass_=True, log=True, screenshot=False)
                    if col_name in webexcel.columns.values and col_name in excel.columns.values:
                        webex = webexcel[col_name]
                        compex = excel[col_name]
                        word = []
                        for row in docs.tables[table_count].rows:
                            word.append(row.cells[count].text)

                        webex = [item for item in webex if str(item) != 'nan']
                        compex = [item for item in compex if str(item) != 'nan']
                        word.pop(0)

                        if len(webex) == len(compex) == len(word):
                            self.LogScreenshot.fLogScreenshot(message=f"Elements length is matching between Web_Excel, "
                                                                      f"Complete Excel and Complete Word Report. "
                                                                      f"WebExcel Elements Length: {len(webex)}\n "
                                                                      f"Excel Elements Length: {len(compex)}\n "
                                                                      f"Word Elements Length: {len(word)}\n",
                                                              pass_=True, log=True, screenshot=False)      

                            comparison_result = self.list_comparison_between_reports_data(word, compex,
                                                                                          webex_list=webex)

                            if len(comparison_result) == 0:
                                self.LogScreenshot.fLogScreenshot(message=f"From '{value}', Values in Column "
                                                                          f"'{col_name}' are matching between "
                                                                          f"WebExcel, Complete Excel "
                                                                          f"and Word Reports.\n",
                                                                  pass_=True, log=True, screenshot=False)
                            else:
                                self.LogScreenshot.fLogScreenshot(message=f"From '{value}', Values in Column "
                                                                          f"'{col_name}' are not matching between "
                                                                          f"WebExcel, Complete Excel and Word Reports."
                                                                          f"\n Mismatch values are arranged in "
                                                                          f"following order -> Word, Complete Excel "
                                                                          f"and WebExcel Report. {comparison_result}",
                                                                  pass_=False, log=True, screenshot=False)
                                raise Exception("Elements are not matching between Webexcel, Complete Excel and "
                                                "Word Reports")
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f"Elements length is not matching between "
                                                                      f"Web_Excel, Complete Excel and Complete "
                                                                      f"Word Report. WebExcel Elements Length: "
                                                                      f"{len(webex)}\n Excel Elements Length: "
                                                                      f"{len(compex)}\n Word Elements Length: "
                                                                      f"{len(word)}\n",
                                                              pass_=False, log=True, screenshot=False)
                            raise Exception(f"Elements length is not matching between Web_Excel, Complete Excel "
                                            f"and Complete Word Report.")
                    else:
                        raise Exception("Column names are not matching between Webexcel, Complete Excel and "
                                        "Word Reports")
                table_count += 1
            except Exception:
                raise Exception("Error in Word report content validation")

    def check_sorting_order_in_excel_report(self, webexcel_filename, excel_filename):
        self.LogScreenshot.fLogScreenshot(message=f"Check the sorting order in Complete Excel and Web_Excel reports",
                                          pass_=True, log=True, screenshot=False)
        self.LogScreenshot.fLogScreenshot(message=f"FileNames are: {webexcel_filename} and \n{excel_filename}",
                                          pass_=True, log=True, screenshot=False)
   
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
            compexcel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=sheet, skiprows=3)

            # Check sorting order from downloaded excel report
            webex_identifier = webexcel["Study Identifier"]
            webex_identifier = [item for item in webex_identifier if str(item) != 'nan']

            webex_shortreference = webexcel["Short Reference"]
            webex_shortreference = [item for item in webex_shortreference if str(item) != 'nan']

            compex_identifier = compexcel["Study Identifier"]
            compex_identifier = [item for item in compex_identifier if str(item) != 'nan']

            compex_shortreference = compexcel["Short Reference"]
            compex_shortreference = [item for item in compex_shortreference if str(item) != 'nan']

            # Finding duplicate values
            compex_dup_identifier = self.get_duplicates_from_list(compex_identifier)
            compex_dup_shortref = self.get_duplicates_from_list(compex_shortreference)

            webex_dup_identifier = self.get_duplicates_from_list(webex_identifier)
            webex_dup_shortref = self.get_duplicates_from_list(webex_shortreference)

            if webex_identifier == sorted(webex_identifier) and compex_identifier == sorted(compex_identifier):
                self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Contents in column Study "
                                                          f"Identifier are in sorted order",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Contents in column Study "
                                                          f"Identifier are not in sorted order",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"From Sheet '{sheet}', Contents in column Study Identifier are not in sorted order")
            
            # When Study Identifier contains duplicate values then corresponding value in Short Reference
            # column should be in sorted order
            if len(compex_dup_identifier) != 0:
                for m in compex_dup_identifier:
                    col_val = compexcel[compexcel["Study Identifier"] == m]
                    col_val_res = col_val["Short Reference"]
                    col_val_res = [item for item in col_val_res if str(item) != 'nan']
                    if col_val_res == sorted(col_val_res):
                        self.LogScreenshot.fLogScreenshot(message=f"For '{m}' in Study Identifier, corresponding "
                                                                  f"contents in column 'Short Reference' are in "
                                                                  f"sorted order in complete excel report",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"For '{m}' in Study Identifier, corresponding "
                                                                  f"contents in column 'Short Reference' are not "
                                                                  f"in sorted order in complete excel report",
                                                          pass_=True, log=True, screenshot=False)
                        raise Exception(f"For '{m}' in Study Identifier, corresponding contents in column "
                                        f"'Short Reference' are not in sorted order in complete excel report")
                
                # When Short Reference contains duplicate values then corresponding value in Publication Type
                # column should be in sorted order
                if len(compex_dup_shortref) != 0:
                    for n in compex_dup_shortref:
                        col_val = compexcel[compexcel["Short Reference"] == n]
                        col_val_res = col_val["Publication Type"]
                        col_val_res = [item for item in col_val_res if str(item) != 'nan']
                        if col_val_res == sorted(col_val_res):
                            self.LogScreenshot.fLogScreenshot(message=f"For '{n}' in Short Reference, corresponding "
                                                                      f"contents in column 'Publication Type' are in "
                                                                      f"sorted order in complete excel report",
                                                              pass_=True, log=True, screenshot=False)
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f"For '{n}' in Short Reference, corresponding "
                                                                      f"contents in column 'Publication Type' are not "
                                                                      f"in sorted order in complete excel report",
                                                              pass_=True, log=True, screenshot=False)
                            raise Exception(f"For '{n}' in Short Reference, corresponding contents in column "
                                            f"'Publication Type' are not in sorted order in complete excel report")

            # When Study Identifier contains duplicate values then corresponding value in Short Reference column
            # should be in sorted order
            if len(webex_dup_identifier) != 0:
                for m in webex_dup_identifier:
                    col_val = webexcel[webexcel["Study Identifier"] == m]
                    col_val_res = col_val["Short Reference"]
                    col_val_res = [item for item in col_val_res if str(item) != 'nan']
                    if col_val_res == sorted(col_val_res):
                        self.LogScreenshot.fLogScreenshot(message=f"For '{m}' in Study Identifier, corresponding "
                                                                  f"contents in column 'Short Reference' are in "
                                                                  f"sorted order in Web_excel report",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"For '{m}' in Study Identifier, corresponding "
                                                                  f"contents in column 'Short Reference' are not in "
                                                                  f"sorted order in Web_excel report",
                                                          pass_=True, log=True, screenshot=False)
                        raise Exception(f"For '{m}' in Study Identifier, corresponding contents in column "
                                        f"'Short Reference' are not in sorted order in Web_excel report")
                
                # When Short Reference contains duplicate values then corresponding value in Publication Type
                # column should be in sorted order
                if len(webex_dup_shortref) != 0:
                    for n in webex_dup_shortref:
                        col_val = webexcel[webexcel["Short Reference"] == n]
                        col_val_res = col_val["Publication Type"]
                        col_val_res = [item for item in col_val_res if str(item) != 'nan']
                        if col_val_res == sorted(col_val_res):
                            self.LogScreenshot.fLogScreenshot(message=f"For '{n}' in Short Reference, corresponding "
                                                                      f"contents in column 'Publication Type' are in "
                                                                      f"sorted order in Web_excel report",
                                                              pass_=True, log=True, screenshot=False)
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f"For '{n}' in Short Reference, corresponding "
                                                                      f"contents in column 'Publication Type' are not "
                                                                      f"in sorted order in Web_excel report",
                                                              pass_=True, log=True, screenshot=False)
                            raise Exception(f"For '{n}' in Short Reference, corresponding contents in column "
                                            f"'Publication Type' are not in sorted order in Web_excel report")

    def test_prisma_ele_comparison_between_Excel_and_Word_Report(self, pop_data, slr_type, add_criteria, filepath):

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(2)
        self.select_data(f"{pop_data[0][0]}", f"{pop_data[0][1]}")
        self.select_data(f"{slr_type[0][0]}", f"{slr_type[0][1]}")
        self.select_sub_section(f"{add_criteria[0][0]}", f"{add_criteria[0][1]}", f"{add_criteria[0][2]}")
        self.select_sub_section(f"{add_criteria[1][0]}", f"{add_criteria[1][1]}", f"{add_criteria[1][2]}")
        self.select_sub_section(f"{add_criteria[2][0]}", f"{add_criteria[2][1]}", f"{add_criteria[2][2]}")
        self.select_sub_section(f"{add_criteria[3][0]}", f"{add_criteria[3][1]}", f"{add_criteria[3][2]}")

        self.generate_download_report("excel_report")
        time.sleep(5)
        excel_filename = self.getFilenameAndValidate(180)
        self.validate_filename(excel_filename, filepath)

        self.generate_download_report("word_report")
        time.sleep(5)
        word_filename = self.getFilenameAndValidate(180)
        self.validate_filename(word_filename, filepath)

        excel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='Updated PRISMA', usecols='C', skiprows=11)
        excel_studies_col = excel['Unique studies']
        # Removing NAN values
        excel_studies_col = [item for item in excel_studies_col if str(item) != 'nan']
        # Converting list values from float to int
        excel_studies_col = [int(x) for x in excel_studies_col]

        # Reading PRISMA Value from Complete Word Report
        docs = docx.Document(f'ActualOutputs//{word_filename}')
        word = []
        for row in docs.tables[4].rows:
            word.append(row.cells[1].text)
        # Removing unwanted values
        word.pop(2)
        word.pop(0)
        # Converting list values from str to int
        word = [int(x) for x in word]

        if excel_studies_col == word:
            self.LogScreenshot.fLogScreenshot(message=f"'Updated Prisma' count is matching between Excel and Word "
                                                      f"report. Excel report 'Updated Prisma' tab count is "
                                                      f"{excel_studies_col} and Word Report 'Updated Prisma' "
                                                      f"table count is {word}",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"'Updated Prisma' count is not matching between Excel and "
                                                      f"Word report. Excel report 'Updated Prisma' tab count is "
                                                      f"{excel_studies_col} and Word Report 'Updated Prisma' "
                                                      f"table count is {word}",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"'Updated Prisma' count is not matching between Excel and Word report. Excel report "
                            f"'Updated Prisma' tab count is {excel_studies_col} and Word Report 'Updated Prisma' "
                            f"table count is {word}")

    def test_prisma_ele_comparison_between_Excel_and_UI(self, pop_data, slr_type, add_criteria, filepath):

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(2)
        self.select_data(f"{pop_data[0][0]}", f"{pop_data[0][1]}")
        self.select_data(f"{slr_type[0][0]}", f"{slr_type[0][1]}")
        self.select_sub_section(f"{add_criteria[0][0]}", f"{add_criteria[0][1]}", f"{add_criteria[0][2]}")
        self.select_sub_section(f"{add_criteria[1][0]}", f"{add_criteria[1][1]}", f"{add_criteria[1][2]}")
        self.select_sub_section(f"{add_criteria[2][0]}", f"{add_criteria[2][1]}", f"{add_criteria[2][2]}")
        self.select_sub_section(f"{add_criteria[3][0]}", f"{add_criteria[3][1]}", f"{add_criteria[3][2]}")

        self.scroll("New_total_selected")
        locators = ['Original_SLR_Count', 'Sub_Pop_Count', 'Line_of_Therapy_Count', 'Intervention_Count',
                    'Study_Design_Count', 'Reported_Var_Count', 'New_total_selected']
        ui_add_critera_values = []
        for i in locators:
            ui_add_critera_values.append(self.get_text(i))
        # Converting list values from str to int
        ui_add_critera_values = [int(x) for x in ui_add_critera_values]
        
        self.LogScreenshot.fLogScreenshot(message=f"Original SLR Count : {ui_add_critera_values[0]}, Sub Pop Count : "
                                                  f"{ui_add_critera_values[1]}, LOT Count : {ui_add_critera_values[2]},"
                                                  f" Intervention Count : {ui_add_critera_values[3]}, "
                                                  f"Study Design Count : {ui_add_critera_values[4]}, Reported "
                                                  f"Variable Count : {ui_add_critera_values[5]}, Total Selected "
                                                  f"Count : {ui_add_critera_values[6]}",
                                          pass_=True, log=True, screenshot=True)
        # original_slr_count = self.get_text("Original_SLR_Count")
        # sub_pop_count = self.get_text("Sub_Pop_Count")
        # lot_count = self.get_text("Line_of_Therapy_Count")
        # intervention_count = self.get_text("Intervention_Count")
        # stdy_dsgn_count = self.get_text("Study_Design_Count")
        # rptd_var_count = self.get_text("Reported_Var_Count")
        # total_sel_count = self.get_text("New_total_selected")

        # self.LogScreenshot.fLogScreenshot(message=f"Original SLR Count : {original_slr_count}, "
        #                                           f"Sub Pop Count : {sub_pop_count}, LOT Count : {lot_count}, "
        #                                           f"Intervention Count : {intervention_count}, Study Design "
        #                                           f"Count : {stdy_dsgn_count}, Reported Variable Count : "
        #                                           f"{rptd_var_count}, Total Selected Count : {total_sel_count}",
        #                                   pass_=True, log=True, screenshot=True)

        self.generate_download_report("excel_report")
        time.sleep(5)
        excel_filename = self.getFilenameAndValidate(180)
        self.validate_filename(excel_filename, filepath)

        excel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='Updated PRISMA', usecols='C', skiprows=11)
        excel_studies_col = excel['Unique studies']
        # Removing NAN values
        excel_studies_col = [item for item in excel_studies_col if str(item) != 'nan']
        # Converting list values from float to int
        excel_studies_col = [int(x) for x in excel_studies_col]

        if excel_studies_col == ui_add_critera_values:
            self.LogScreenshot.fLogScreenshot(message=f"Count seen in excel report under 'Updated Prisma tab' is "
                                                      f"same as 'Updated PRISMA table' in UI",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Count seen in excel report under 'Updated Prisma tab' is "
                                                      f"not matching with 'Updated PRISMA table' in UI. Excel Report "
                                                      f"values are {excel_studies_col} and UI values are "
                                                      f"{ui_add_critera_values}",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Count seen in excel report under 'Updated Prisma tab' is not matching with "
                            f"'Updated PRISMA table' in UI")

    def test_prisma_count_comparison_between_prismatab_and_excludedstudiesliveslr(self, pop_data, slr_type,
                                                                                  add_criteria, filepath):

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(2)
        self.select_data(f"{pop_data[0][0]}", f"{pop_data[0][1]}")
        self.select_data(f"{slr_type[0][0]}", f"{slr_type[0][1]}")
        self.select_sub_section(f"{add_criteria[0][0]}", f"{add_criteria[0][1]}", f"{add_criteria[0][2]}")
        self.select_sub_section(f"{add_criteria[1][0]}", f"{add_criteria[1][1]}", f"{add_criteria[1][2]}")
        self.select_sub_section(f"{add_criteria[2][0]}", f"{add_criteria[2][1]}", f"{add_criteria[2][2]}")
        self.select_sub_section(f"{add_criteria[3][0]}", f"{add_criteria[3][1]}", f"{add_criteria[3][2]}")

        self.generate_download_report("excel_report")
        time.sleep(5)
        excel_filename = self.getFilenameAndValidate(180)
        self.validate_filename(excel_filename, filepath)

        prisma_tab = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='Updated PRISMA', usecols='C',
                                   skiprows=11)
        excludedstudies_liveslr_tab = pd.read_excel(f'ActualOutputs//{excel_filename}',
                                                    sheet_name='Excluded studies - LiveSLR', skiprows=2)
        prisma_tab_col = prisma_tab['Unique studies']
        # Removing NAN values
        prisma_tab_col = [item for item in prisma_tab_col if str(item) != 'nan']
        # Converting list values from float to int
        prisma_tab_col = [int(x) for x in prisma_tab_col]

        # Get Unique Study Identifier value from corresponding to Reason for Rejection values
        excludedstudies_liveslr = excludedstudies_liveslr_tab["Reason for Rejection"]
        excludedstudies_liveslr = [item for item in excludedstudies_liveslr if str(item) != 'nan']
        excludedstudies_liveslr_final = []
        # Removing the duplicates
        [excludedstudies_liveslr_final.append(x) for x in list(flatten(excludedstudies_liveslr))
         if x not in excludedstudies_liveslr_final]

        self.LogScreenshot.fLogScreenshot(message=f"Unique Reason for Rejection column values are : "
                                                  f"{excludedstudies_liveslr_final}",
                                          pass_=True, log=True, screenshot=False)

        for m in excludedstudies_liveslr_final:
            unique_study_identifier = []
            col_val = excludedstudies_liveslr_tab[excludedstudies_liveslr_tab["Reason for Rejection"] == m]
            study_identifier = col_val["Study Identifier"]
            study_identifier = [item for item in study_identifier if str(item) != 'nan']
            # Removing duplicates to get the proper length of Study Identifier data
            [unique_study_identifier.append(x) for x in list(flatten(study_identifier))
             if x not in unique_study_identifier]
            if len(unique_study_identifier) in prisma_tab_col[1:6]:
                self.LogScreenshot.fLogScreenshot(message=f"From 'Excluded studies - LiveSLR' sheet -> for '{m}' "
                                                          f"Reason for Rejection the Selected studies(Unique Study "
                                                          f"Identifier)count is matching with the selected count in "
                                                          f"Updated PRISMA tab",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"From 'Excluded studies - LiveSLR' sheet -> for '{m}' "
                                                          f"Reason for Rejection the Selected studies(Unique Study "
                                                          f"Identifier)count is not matching with the selected count "
                                                          f"in Updated PRISMA tab",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"From 'Excluded studies - LiveSLR' sheet -> for '{m}' Reason for Rejection the "
                                f"Selected studies(Unique Study Identifier)count is not matching with the selected "
                                f"count in Updated PRISMA tab")

    def test_prisma_tab_format_changes(self, pop_data, slr_type, add_criteria, filepath):

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(2)
        self.select_data(f"{pop_data[0][0]}", f"{pop_data[0][1]}")
        self.select_data(f"{slr_type[0][0]}", f"{slr_type[0][1]}")
        self.select_sub_section(f"{add_criteria[0][0]}", f"{add_criteria[0][1]}", f"{add_criteria[0][2]}")
        self.select_sub_section(f"{add_criteria[1][0]}", f"{add_criteria[1][1]}", f"{add_criteria[1][2]}")
        self.select_sub_section(f"{add_criteria[2][0]}", f"{add_criteria[2][1]}", f"{add_criteria[2][2]}")
        self.select_sub_section(f"{add_criteria[3][0]}", f"{add_criteria[3][1]}", f"{add_criteria[3][2]}")

        self.generate_download_report("excel_report")
        time.sleep(5)
        excel_filename = self.getFilenameAndValidate(180)
        self.validate_filename(excel_filename, filepath)

        # Read PRISMA value from Complete Excel Sheet
        excel = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
        excel_sheet = excel['Updated PRISMA']
        row_13 = excel_sheet['B13'].value
        row_12 = excel_sheet['C12'].value
        row_14 = excel_sheet['B14'].value
        row_15 = excel_sheet['B15'].value
        row_21 = excel_sheet['B21'].value
        row_23 = excel_sheet['B23'].value
        
        if "From original SLR" in row_13:
            self.LogScreenshot.fLogScreenshot(message=f"Format has been updated from 'Original SLR' to 'From "
                                                      f"original SLR'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Format has not been updated from 'Original SLR' to 'From "
                                                      f"original SLR'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Format has not been updated from 'Original SLR' to 'From original SLR'")

        if "Unique studies" in row_12:
            self.LogScreenshot.fLogScreenshot(message=f"Column Title has been added with name 'Unique studies'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Column Title has not been added with name 'Unique studies'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Column Title has not been added with name 'Unique studies'")

        if row_14 is None:
            self.LogScreenshot.fLogScreenshot(message=f"An empty row below 'From original SLR' has been added",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"An empty row below 'From original SLR' has not been added",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"An empty row below 'From original SLR' has not been added")

        if "Excluded based on" in row_15:
            self.LogScreenshot.fLogScreenshot(message=f"An additional row has been added above 'sub-population' and "
                                                      f"the column title is 'Excluded based on'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"An additional row has not been added above 'sub-population'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"An additional row has not been added above 'sub-population'")

        if row_21 is None:
            self.LogScreenshot.fLogScreenshot(message=f"An empty row above 'New total selected' has been added",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"An empty row above 'New total selected' has not been added",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"An empty row above 'New total selected' has not been added")

        if "Note: LiveSLR counts the number of original studies, however, for each original study, there might be " \
           "multiple extractions covering original or sub-group patient populations. Therefore, the number of the " \
           "new total selected and the number of excluded studies may not add up to be the total original SLR count." \
           " Please contact your Cytel SLR project lead if you require more assistance on this." in row_23:
            self.LogScreenshot.fLogScreenshot(message=f"Static text '{row_23}' has been added.",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Static text '{row_23}' has not been added.",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Static text '{row_23}' has not been added.")                                            

    def test_interventional_to_clinical_changes(self, filepath):

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(1)
        self.select_data("NewImportLogic_1 - Test_Automation_1", "NewImportLogic_1 - Test_Automation_1_radio_button")
        time.sleep(1)
        type_of_slr_text = self.get_text("slrtype_first_option_text")
        self.select_data("Clinical", "Clinical_radio_button")

        self.click("NMA_Button")
        nma_param_clinical_list = self.get_text("NMA_parameters_list_clinical")

        self.generate_download_report("excel_report")
        time.sleep(5)
        excel_filename = self.getFilenameAndValidate(180)
        self.validate_filename(excel_filename, filepath)

        self.preview_result("preview_results")
        self.table_display_check("Table")
        web_table_title = self.get_text("web_table_title")
        self.generate_download_report("Export_as_excel")
        time.sleep(5)
        webexcel_filename = self.getFilenameAndValidate(180)
        self.validate_filename(webexcel_filename, filepath)
        self.back_to_report_page("Back_to_search_page")

        if search("Clinical", type_of_slr_text) and not search("Interventional", type_of_slr_text):
            self.LogScreenshot.fLogScreenshot(message=f"From 'Search LiveSLR' page -> Under Select Type of SLR first "
                                                      f"option is updated from Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"From 'Search LiveSLR' page -> Under Select Type of SLR first "
                                                      f"option is not updated from Interventional to 'Clinical'. "
                                                      f"Actual value is {type_of_slr_text}",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"From 'Search LiveSLR' page -> Under Select Type of SLR first option is not updated "
                            f"from Interventional to 'Clinical'")
                
        if search("Clinical", nma_param_clinical_list) and not search("Interventional", nma_param_clinical_list):
            self.LogScreenshot.fLogScreenshot(message=f"From 'Search LiveSLR' page -> Go to 'Select Data for  NMA' -> "
                                                      f"Under select NMA parameters section is updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"From 'Search LiveSLR' page -> Go to 'Select Data for  NMA' -> "
                                                      f"Under select NMA parameters section is not updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"From 'Search LiveSLR' page -> Go to 'Select Data for  NMA' -> Under select NMA "
                            f"parameters section is not updated from Interventional to 'Clinical'")

        self.click("protocol_link")
        pages = [['picos', 'picos_pop_dropdown', 'picos_study_type_dropdown'],
                 ['searchstrategy', 'searchstrategy_pop_dropdown', 'searchstrategy_study_type_dropdown'],
                 ['prismas', 'prisma_pop_dropdown', 'prisma_study_type_dropdown']]
        for i in pages:
            self.click(i[0])
            pop_ele = self.select_element(i[1])
            select1 = Select(pop_ele)
            select1.select_by_visible_text("NewImportLogic_1 - Test_Automation_1")

            stdy_ele = self.select_element(i[2])
            select2 = Select(stdy_ele)
            opts = select2.options
            dropdown_li = []
            for j in opts:
                dropdown_li.append(j.text)
            
            if "Clinical" in dropdown_li and "Interventional" not in dropdown_li:
                self.LogScreenshot.fLogScreenshot(message=f"From page '{i[0]}' -> Under Study type dropdown "
                                                          f"Interventional is changed to 'Clinical'",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"From page '{i[0]}' -> Under Study type dropdown "
                                                          f"Interventional is not changed to 'Clinical'",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"From page '{i[0]}' -> Under Study type dropdown Interventional is not changed "
                                f"to 'Clinical'")

        self.click("manage_qa_data_button")
        pop_ele = self.select_element("select_pop_dropdown")
        select1 = Select(pop_ele)
        select1.select_by_visible_text("NewImportLogic_1 - Test_Automation_1")

        stdy_ele = self.select_element("select_stdy_type_dropdown")
        select2 = Select(stdy_ele)
        opts = select2.options
        qa_data_dropdown_li = []
        for m in opts:
            qa_data_dropdown_li.append(m.text)
        
        if "Clinical" in qa_data_dropdown_li and "Interventional" not in qa_data_dropdown_li:
            self.LogScreenshot.fLogScreenshot(message=f"From 'Manage QA Data' page -> Under Study type dropdown "
                                                      f"Interventional is changed to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"From 'Manage QA Data' page -> Under Study type dropdown "
                                                      f"Interventional is not changed to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"From 'Manage QA Data' page -> Under Study type dropdown Interventional is not changed "
                            f"to 'Clinical'")
        
        if search("Clinical", web_table_title) and not search("Interventional", web_table_title):
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Report -> Title of 'Preview Results' section is "
                                                      f"updated from Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Report -> Title of 'Preview Results' section is not "
                                                      f"updated from Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"WebExcel Report -> Title of 'Preview Results' section is not updated from "
                            f"Interventional to 'Clinical'")
        
        # Read value from Complete Excel and WebExcel
        excel = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
        webexcel = openpyxl.load_workbook(f'ActualOutputs//{webexcel_filename}')
        excel_sheetnames = excel.sheetnames
        webexcel_sheetnames = webexcel.sheetnames

        webex_sheet = webexcel.active
        webex_title = webex_sheet['A1'].value
        if search("Clinical", webex_title) and not search("Interventional", webex_title) and "Clinical Report" in \
                webexcel_sheetnames:
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Report -> Title and sheet name is updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"WebExcel Report -> Title and sheet name is not updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"WebExcel Report -> Title and sheet name is not updated from Interventional to 'Clinical'")

        excel_sheet_toc = excel['TOC']
        toc_heading = excel_sheet_toc['A1'].value
        toc_content_6 = excel_sheet_toc['B10'].value

        if search("Clinical", toc_heading) and not search("Interventional", toc_heading):
            self.LogScreenshot.fLogScreenshot(message=f"TOC sheet heading is updated from Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"TOC sheet heading is not updated from Interventional to "
                                                      f"'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"TOC sheet heading is not updated from Interventional to 'Clinical'")

        if search("Clinical", toc_content_6) and not search("Interventional", toc_content_6):
            self.LogScreenshot.fLogScreenshot(message=f"TOC sheet -> table of content number 6 is updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"TOC sheet -> table of content number 6 is not updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"TOC sheet -> table of content number 6 is not updated from Interventional to 'Clinical'")

        excel_sheet_picos = excel['PICOS']
        picos_heading = excel_sheet_picos['A3'].value

        if search("CLINICAL", picos_heading) and not search("INTERVENTIONAL", picos_heading):
            self.LogScreenshot.fLogScreenshot(message=f"PICOS sheet heading is updated from INTERVENTIONAL to "
                                                      f"'CLINICAL'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"PICOS sheet heading is not updated from INTERVENTIONAL to "
                                                      f"'CLINICAL'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"PICOS sheet heading is not updated from INTERVENTIONAL to 'CLINICAL'")

        excel_sheet_inc_exc = excel['INC-EXC']
        inc_sheet_heading = excel_sheet_inc_exc['A1'].value
        inc_table_heading = excel_sheet_inc_exc['A4'].value

        if search("Clinical", inc_sheet_heading) and not search("Interventional", inc_sheet_heading):
            self.LogScreenshot.fLogScreenshot(message=f"INCLUSION AND EXCLUSION CRITERIA sheet heading is updated "
                                                      f"from Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"INCLUSION AND EXCLUSION CRITERIA sheet heading is not "
                                                      f"updated from Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"INCLUSION AND EXCLUSION CRITERIA sheet heading is not updated from Interventional to "
                            f"'Clinical'")

        if search("Clinical", inc_table_heading) and not search("Interventional", inc_table_heading):
            self.LogScreenshot.fLogScreenshot(message=f"INCLUSION AND EXCLUSION CRITERIA table heading is updated "
                                                      f"from Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"INCLUSION AND EXCLUSION CRITERIA table heading is not updated "
                                                      f"from Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"INCLUSION AND EXCLUSION CRITERIA table heading is not updated from Interventional "
                            f"to 'Clinical'")

        excel_sheet_strategy = excel['SEARCH STRATEGIES']
        strategy_sheet_heading = excel_sheet_strategy['A1'].value
        strategy_table_heading = excel_sheet_strategy['A3'].value

        if search("Clinical", strategy_sheet_heading) and not search("Interventional", strategy_sheet_heading):
            self.LogScreenshot.fLogScreenshot(message=f"SEARCH STRATEGY sheet heading is updated from Interventional "
                                                      f"to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"SEARCH STRATEGY sheet heading is not updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"SEARCH STRATEGY sheet heading is not updated from Interventional to 'Clinical'")

        if search("Clinical", strategy_table_heading) and not search("Interventional", strategy_table_heading):
            self.LogScreenshot.fLogScreenshot(message=f"SEARCH STRATEGY table heading is updated from Interventional "
                                                      f"to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"SEARCH STRATEGY table heading is not updated from "
                                                      f"Interventional to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"SEARCH STRATEGY table heading is not updated from Interventional to 'Clinical'")

        excel_sheet_report = excel['Clinical Report']
        report_sheet_heading = excel_sheet_report['A1'].value 

        if "Clinical Report" in excel_sheetnames and "Interventional Report" not in excel_sheetnames:
            self.LogScreenshot.fLogScreenshot(message=f"Report sheet name is updated from Interventional to 'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Report sheet name is not updated from Interventional to "
                                                      f"'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Report sheet name is not updated from Interventional to 'Clinical'")

        if search("Clinical", report_sheet_heading) and not search("Interventional", report_sheet_heading):
            self.LogScreenshot.fLogScreenshot(message=f"Report sheet heading is updated from Interventional to "
                                                      f"'Clinical'",
                                              pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Report sheet heading is not updated from Interventional "
                                                      f"to 'Clinical'",
                                              pass_=False, log=True, screenshot=False)
            raise Exception(f"Report sheet heading is not updated from Interventional to 'Clinical'")

    def test_publication_identifier_count_in_updated_prisma_tab(self, pop_data, slr_type, add_criteria, filepath):

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(2)
        self.select_data(f"{pop_data[0][0]}", f"{pop_data[0][1]}")
        self.select_data(f"{slr_type[0][0]}", f"{slr_type[0][1]}")
        self.select_sub_section(f"{add_criteria[0][0]}", f"{add_criteria[0][1]}", f"{add_criteria[0][2]}")
        self.select_sub_section(f"{add_criteria[1][0]}", f"{add_criteria[1][1]}", f"{add_criteria[1][2]}")
        self.select_sub_section(f"{add_criteria[2][0]}", f"{add_criteria[2][1]}", f"{add_criteria[2][2]}")
        self.select_sub_section(f"{add_criteria[3][0]}", f"{add_criteria[3][1]}", f"{add_criteria[3][2]}")

        self.generate_download_report("excel_report")
        time.sleep(5)
        excel_filename = self.getFilenameAndValidate(180)
        self.validate_filename(excel_filename, filepath)

        prisma_tab = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='Updated PRISMA', usecols='D',
                                   skiprows=11)
        excludedstudies_liveslr_tab = pd.read_excel(f'ActualOutputs//{excel_filename}',
                                                    sheet_name='Excluded studies - LiveSLR', skiprows=2)
        prisma_tab_col = prisma_tab['Publications']
        # Removing NAN values
        prisma_tab_col = [item for item in prisma_tab_col if str(item) != 'nan']
        # Converting list values from float to int
        prisma_tab_col = [int(x) for x in prisma_tab_col]

        # Get Unique Study Identifier value from corresponding to Reason for Rejection values
        excludedstudies_liveslr = excludedstudies_liveslr_tab["Reason for Rejection"]
        excludedstudies_liveslr = [item for item in excludedstudies_liveslr if str(item) != 'nan']
        excludedstudies_liveslr_final = []
        # Removing the duplicates
        [excludedstudies_liveslr_final.append(x) for x in list(flatten(excludedstudies_liveslr))
         if x not in excludedstudies_liveslr_final]

        self.LogScreenshot.fLogScreenshot(message=f"Unique Reason for Rejection column values are : "
                                                  f"{excludedstudies_liveslr_final}",
                                          pass_=True, log=True, screenshot=False)

        for m in excludedstudies_liveslr_final:
            unique_pub_identifier = []
            col_val = excludedstudies_liveslr_tab[excludedstudies_liveslr_tab["Reason for Rejection"] == m]
            pub_identifier = col_val["Publication Identifier"]
            pub_identifier = [item for item in pub_identifier if str(item) != 'nan']
            # Checking if list is nested or not
            if any(isinstance(i, str) for i in pub_identifier):
                # Converting a list containing String type values into list/tuple type
                res = []
                for xy in pub_identifier:
                    res.append(ast.literal_eval(xy))
                # Removing duplicates to get the proper length of Study Identifier data
                [unique_pub_identifier.append(x) for x in list(flatten(res))
                 if x not in unique_pub_identifier]
            else:
                # Removing duplicates to get the proper length of Study Identifier data
                [unique_pub_identifier.append(x) for x in list(flatten(pub_identifier))
                 if x not in unique_pub_identifier]

            if len(unique_pub_identifier) in prisma_tab_col[1:6]:
                self.LogScreenshot.fLogScreenshot(message=f"From 'Excluded studies - LiveSLR' sheet -> for '{m}' "
                                                          f"Reason for Rejection the unique Publication "
                                                          f"Identifier count is matching with the count in "
                                                          f"Updated PRISMA tab -> 'Publications' column. The count "
                                                          f"value is : {len(unique_pub_identifier)}",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"From 'Excluded studies - LiveSLR' sheet -> for '{m}' "
                                                          f"Reason for Rejection the unique Publication "
                                                          f"Identifier count is not matching with the count in "
                                                          f"Updated PRISMA tab -> 'Publications' column. The count "
                                                          f"value is : {len(unique_pub_identifier)}",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception(f"From 'Excluded studies - LiveSLR' sheet -> for '{m}' Reason for Rejection the "
                                f"unique Publication Identifier count is not matching with the "
                                f"count in Updated PRISMA tab -> 'Publications' column")            

    def validate_population_col_in_wordreport(self, filepath, locatorname):
        self.LogScreenshot.fLogScreenshot(message=f"Validate contents of Population/Sub-group column in Word Report",
                                          pass_=True, log=True, screenshot=False)
        source_template = self.exbase.get_source_template(filepath, 'Sheet1', locatorname)

        # Read population details from data sheet
        extraction_file = self.exbase.get_file_details_to_upload(filepath, locatorname)       

        # Read population data values
        pop_list = self.exbase.get_population_data(filepath, 'Sheet1', locatorname)
        # Read slrtype data values
        slrtype = self.exbase.get_slrtype_data(filepath, 'Sheet1', locatorname)         

        self.refreshpage()
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")
        self.exbase.upload_file(extraction_file[0][0], extraction_file[0][1]) 

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(2)
        self.select_data(f"{pop_list[0][0]}", f"{pop_list[0][1]}")
        self.select_data(slrtype[0][0], f"{slrtype[0][1]}")
        
        self.generate_download_report("word_report")
        time.sleep(5)
        word_filename = self.getFilenameAndValidate(180)
        self.validate_filename(word_filename, filepath)

        table_count = 5  

        sourcefile = pd.read_excel(f'{source_template[0]}') 
        docs = docx.Document(f'ActualOutputs//{word_filename}')
        try:
            table = docs.tables[table_count]
            data = [[cell.text for cell in row.cells] for row in table.rows]
            df_word = pd.DataFrame(data)

            # Check whether the column name is updated or not
            if 'Population/Sub-group' in df_word.values[0] and 'Population' not in df_word.values[0]:
                self.LogScreenshot.fLogScreenshot(message=f"Column name has been updated from 'Population' to "
                                                          f"'Population/Sub-group'",
                                                  pass_=True, log=True, screenshot=False)
            else:
                self.LogScreenshot.fLogScreenshot(message=f"Column name has not been updated from 'Population' to "
                                                          f"'Population/Sub-group'",
                                                  pass_=False, log=True, screenshot=False)
                raise Exception("Column name has not been updated from 'Population' to 'Population/Sub-group'")

            # Using count variable to loop over columns in word document
            count = 0        
            # df_word.values[0] will give the list of column names from the table in Word document        
            for j in df_word.values[0]:
                # Check the column name is present in Expected test data file
                if j in sourcefile.columns.values:
                    sourcedata = sourcefile[j]
                    word = []
                    for row in docs.tables[table_count].rows:
                        word.append(row.cells[count].text)

                    # Removing NAN/None values if any
                    sourcedata = [item for item in sourcedata if str(item) != 'nan']
                    # Converting Integer list to String list
                    sourcedata = [str(x) for x in sourcedata]
                    # Popping up the column name
                    word.pop(0)

                    comparison_result = self.list_comparison_between_reports_data(sourcedata, word)

                    if len(comparison_result) == 0:
                        self.LogScreenshot.fLogScreenshot(message=f"Content from Table 2-2 is matching with expected "
                                                                  f"test data.",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"Content from Table 2-2 is not matching with "
                                                                  f"expected test data. Mismatch values are "
                                                                  f"arranged in following order -> Source Excel and "
                                                                  f"Word Report. {comparison_result}",
                                                          pass_=False, log=True, screenshot=False)
                        raise Exception("Elements are not matching between Source data and Word Report")
                    count += 1
                else:
                    raise Exception("Column names are not matching between Source data and "
                                    "Word Report")
            self.refreshpage()
            self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

            self.exbase.delete_file(extraction_file[0][2])

            # Go to live slr page
            self.liveslrpage.go_to_liveslr("SLR_Homepage")
            time.sleep(2)
        except Exception:
            raise Exception("Error in Word report content validation")

    def validate_control_chars_in_wordreport(self, filepath, locatorname):
        self.LogScreenshot.fLogScreenshot(message=f"Validate the accessibility of downloaded reports when extraction "
                                                  f"file contains control characters",
                                          pass_=True, log=True, screenshot=False)

        # Read population details from data sheet
        extraction_file = self.exbase.get_file_details_to_upload(filepath, locatorname)       

        # Read population data values
        pop_list = self.exbase.get_population_data(filepath, 'Sheet1', locatorname)
        # Read slrtype data values
        slrtype = self.exbase.get_slrtype_data(filepath, 'Sheet1', locatorname)         

        self.refreshpage()
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")
        self.exbase.upload_file(extraction_file[0][0], extraction_file[0][1]) 

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")
        time.sleep(2)
        try:
            for i in pop_list:
                self.select_data(i[0], i[1])
                for j in slrtype:
                    self.select_data(j[0], j[1])

                    self.generate_download_report("excel_report")
                    time.sleep(5)
                    excel_filename = self.getFilenameAndValidate(180)
                    self.validate_filename(excel_filename, filepath)

                    self.generate_download_report("word_report")
                    time.sleep(5)
                    word_filename = self.getFilenameAndValidate(180)
                    self.validate_filename(word_filename, filepath)

                    self.preview_result("preview_results")
                    self.table_display_check("Table")
                    self.generate_download_report("Export_as_excel")
                    time.sleep(5)
                    webexcel_filename = self.getFilenameAndValidate(180)
                    self.validate_filename(webexcel_filename, filepath)
                    self.back_to_report_page("Back_to_search_page")

                    '''Checking whether we are able to read data from downloaded reports or not'''
                    webexcel = pd.read_excel(f'ActualOutputs//{webexcel_filename}')
                    excel = pd.read_excel(f'ActualOutputs//{excel_filename}')
                    docs = docx.Document(f'ActualOutputs//{word_filename}')

                    # Reading Table 2-2 in word report
                    table = docs.tables[5]
                    data = [[cell.text for cell in row.cells] for row in table.rows]
                    df_word = pd.DataFrame(data)

                    if not webexcel.empty and not excel.empty and not df_word.empty:
                        self.LogScreenshot.fLogScreenshot(message=f"Able to open WebExcel, Complete Excel, Word "
                                                                  f"Reports and read data successfully.",
                                                          pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"Error while opening and reading the data from "
                                                                  f"WebExcel, Complete Excel, Word Reports.",
                                                          pass_=False, log=True, screenshot=False)
                        raise Exception(f"Error while opening and reading the data from WebExcel, Complete Excel, "
                                        f"Word Reports.")
        except Exception:
            raise Exception("Unable to select element")                  

        self.refreshpage()
        self.imppubpage.go_to_importpublications("importpublications_button", "extraction_upload_btn")

        self.exbase.delete_file(extraction_file[0][2])

        # Go to live slr page
        self.liveslrpage.go_to_liveslr("SLR_Homepage")

    # # ############## Using Openpyxl library #################
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
    #
    #         try:
    #             for i in webexcel_sheet.rows:
    #                 webex.append(i[0].value)
    #
    #             for j in excel_sheet.rows:
    #                 compex.append(j[0].value)
    #
    #             # Removing None values from the lists
    #             webex = list(filter(None, webex))
    #             compex = list(filter(None, compex))
    #
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
    #                         self.LogScreenshot.fLogScreenshot(message=f"Study Identifier are matching between "
    #                                                                   f"WebExcel and Complete Excel report. \n"
    #                                                                   f"Compared Element values are: {ele1} "
    #                                                                   f"and \n{ele2}",
    #                                                           pass_=True, log=True, screenshot=False)
    #                     else:
    #                         self.LogScreenshot.fLogScreenshot(message=f"Study Identifier are not matching between "
    #                                                                   f"WebExcel and Complete Excel report. \n"
    #                                                                   f"Compared Element values are: {ele1} "
    #                                                                   f"and \n{ele2}",
    #                                                           pass_=False, log=True, screenshot=False)
    #                         raise Exception("Study Identifier are not matching")
    #                 else:
    #                     raise Exception("Study Identifier Elements are not in sorted order")
    #         except Exception:
    #             raise Exception("Error in Excel sheet content validation")
    #
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
    #
    #             for j in excel_sheet.rows:
    #                 compex.append(j[3].value)
    #
    #             for row in docs.tables[table_count].rows:
    #                 word.append(row.cells[0].text)
    #             # Incrementing the counter to switch for next table based on the number of excel sheet names
    #             table_count += 1
    #
    #             # Removing None values from the lists
    #             webex = list(filter(None, webex))
    #             compex = list(filter(None, compex))
    #
    #             if compex == webex and compex == word:
    #                 self.LogScreenshot.fLogScreenshot(message=f"Contents are matching in WebExcel, Complete "
    #                                                           f"Excel and Complete Word Reports. "
    #                                                           f"Short Reference values are: \n"
    #                                                           f"{webex} \n {compex} \n {word}",
    #                                                   pass_=True, log=True, screenshot=False)
    #             else:
    #                 self.LogScreenshot.fLogScreenshot(message=f"Contents are not matching in WebExcel, Complete "
    #                                                           f"Excel and Complete Word Reports. "
    #                                                           f"Short Reference values are: \n"
    #                                                           f"{webex} \n {compex} \n {word}",
    #                                                   pass_=False, log=True, screenshot=False)
    #                 raise Exception("Word report contents are not matching")
    #         except Exception:
    #             raise Exception("Error in Word report content validation")
