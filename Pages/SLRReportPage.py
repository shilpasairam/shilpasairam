import numbers
import os
import re
import time

import docx
import openpyxl
import pandas as pd
from pathlib import Path
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

    def get_source_template(self, filepath):
        file = pd.read_excel(filepath)
        expectedfilepath = list(os.getcwd()+file['ExpectedSourceTemplateFile'].dropna())
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
        
        source_template = self.get_source_template(filepath)

        self.LogScreenshot.fLogScreenshot(message=f"Source Filename is: {Path(f'{source_template[0]}').stem}, "
                                                  f"Downloaded FileNames are: {webexcel_filename} "
                                                  f"and \n{excel_filename}",
                                          pass_=True, log=True, screenshot=False)

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
                                                                  f"Duplicate values are arranged in following "
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

    def excel_to_word_content_validation(self, webexcel_filename, excel_filename, word_filename):
        self.LogScreenshot.fLogScreenshot(message=f"Content validation between WebExcel, Complete Excel and "
                                                  f"Complete Word Reports",
                                          pass_=True, log=True, screenshot=False)
        self.LogScreenshot.fLogScreenshot(message=f"FileNames are: {webexcel_filename}, \n{excel_filename}, "
                                                  f"\n{word_filename}",
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

                # Check the length of 1st column from the report to make sure number of rows are as expected
                webex_len = webexcel[df_word.values[0][0]]
                compex_len = excel[df_word.values[0][0]]
                word_len = []
                for row in docs.tables[table_count].rows:
                    word_len.append(row.cells[0].text)

                webex_len = [item for item in webex_len if str(item) != 'nan']
                compex_len = [item for item in compex_len if str(item) != 'nan']
                word_len.pop(0)

                if len(webex_len) == len(compex_len) == len(word_len):
                    self.LogScreenshot.fLogScreenshot(message=f"Elements length is matching between Web_Excel, "
                                                              f"Complete Excel and Complete Word Report. "
                                                              f"WebExcel Elements Length: {len(webex_len)}\n "
                                                              f"Excel Elements Length: {len(compex_len)}\n "
                                                              f"Word Elements Length: {len(word_len)}\n",
                                                      pass_=True, log=True, screenshot=False)

                    # Content validation starts from here
                    # Using count variable to loop over columns in word document
                    count = 0
                    # df_word.values[0] will give the list of column names from the table in Word document
                    for col_name in df_word.values[0]:
                        # Restricting comparison upto 3 columns in word due to data formatting issues in further columns
                        if count <= 2:
                            if col_name == 'Year/Country':
                                # This IF condition is just to add space to the column name as per Excel sheet to
                                # match the names between word and excel. This is an workaround until it is fixed
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

                                comparison_result = self.list_comparison_between_reports_data(word, compex,
                                                                                              webex_list=webex)

                                if len(comparison_result) == 0:
                                    self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Values in "
                                                                              f"Column '{col_name}' are matching "
                                                                              f"between WebExcel, Complete Excel "
                                                                              f"and Word Reports.\n",
                                                                      pass_=True, log=True, screenshot=False)
                                else:
                                    self.LogScreenshot.fLogScreenshot(message=f"From Sheet '{sheet}', Values in "
                                                                              f"Column '{col_name}' are not matching "
                                                                              f"between WebExcel, Complete Excel and "
                                                                              f"Word Reports.\n Duplicate values are "
                                                                              f"arranged in following order -> "
                                                                              f"WebExcel, Complete Excel and Word "
                                                                              f"Report. {comparison_result}",
                                                                      pass_=False, log=True, screenshot=False)
                                    raise Exception("Elements are not matching between Webexcel, Complete Excel "
                                                    "and Word Reports")
                                count += 1
                            else:
                                raise Exception("Column names are not matching between Webexcel, Complete Excel "
                                                "and Word Reports")
                    table_count += 1
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Elements length is not matching between Web_Excel, "
                                                              f"Complete Excel and Complete Word Report. "
                                                              f"WebExcel Elements Length: {len(webex_len)}\n "
                                                              f"Excel Elements Length: {len(compex_len)}\n "
                                                              f"Word Elements Length: {len(word_len)}\n",
                                                      pass_=False, log=True, screenshot=False)
                    raise Exception(f"Elements length is not matching between Web_Excel, Complete Excel and "
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

                        if len(webex) == len(compex) == len(word) and webex == compex == word:
                            self.LogScreenshot.fLogScreenshot(message=f"From '{value}', Values in Column '{col_name}' "
                                                                      f"are matching between WebExcel, Complete Excel "
                                                                      f"and Word Reports.\n"
                                                                      f"WebExcel Elements Length: {len(webex)}\n "
                                                                      f"Excel Elements Length: {len(compex)}\n "
                                                                      f"Word Elements Length: {len(word)}\n",
                                                              pass_=True, log=True, screenshot=False)
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f"From '{value}', Values in Column '{col_name}' "
                                                                      f"are not matching between WebExcel, "
                                                                      f"Complete Excel and Word Reports.\n"
                                                                      f"WebExcel Elements Length: {len(webex)}\n "
                                                                      f"Excel Elements Length: {len(compex)}\n "
                                                                      f"Word Elements Length: {len(word)}\n"
                                                                      f"WebExcel Elements: {webex}\n "
                                                                      f"Excel Elements: {compex}\n "
                                                                      f"Word Elements: {word}",
                                                              pass_=False, log=True, screenshot=False)
                            raise Exception("Elements are not matching between Webexcel, Complete Excel and "
                                            "Word Reports")
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
