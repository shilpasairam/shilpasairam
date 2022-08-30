import numbers
import os
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


class UtilityOutcome(Base):

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

    def get_population_data_specific_sheet(self, filepath, sheet):
        file = pd.read_excel(filepath, sheet_name=sheet)
        pop = list(file['Population'].dropna())
        pop_button = list(file['Population_Radio_button'].dropna())
        population_data = [(pop[i], pop_button[i]) for i in range(0, len(pop))]
        return population_data

    def get_slrtype_data_specific_sheet(self, filepath, sheet):
        file = pd.read_excel(filepath, sheet_name=sheet)
        slrtype = list(file['slrtype'].dropna())
        slrtype_button = list(file['slrtype_Radio_button'].dropna())
        slrtype_data = [(slrtype[i], slrtype_button[i]) for i in range(0, len(slrtype))]
        return slrtype_data

    def get_reported_variables_specific_sheet(self, filepath, sheet):
        file = pd.read_excel(filepath, sheet_name=sheet)
        reported_var = list(file['ReportedVariables'].dropna())
        reported_var_button = list(file['Reportedvariable_checkbox'].dropna())
        # reported_var_data = [(reported_var[i], reported_var_button[i]) for i in range(0, len(reported_var))]
        return reported_var, reported_var_button
    
    def get_util_source_template(self, filepath, sheet):
        file = pd.read_excel(filepath, sheet_name=sheet)
        expectedfilepath = list(os.getcwd()+file['ExpectedSourceTemplateFile'].dropna())
        return expectedfilepath
    
    def qol_presenceof_utility_summary_tab(self, excel_filename, util_filepath):
        source_template = self.get_util_source_template(util_filepath, 'NewImportLogic')
        
        self.LogScreenshot.fLogScreenshot(message=f"*****Check Presence of Utility Summary Tab in Complete Excel Report*****",
                                          pass_=True, log=True, screenshot=False)
        excel_data = openpyxl.load_workbook(f'ActualOutputs//{excel_filename}')
        if 'Utility Summary' in excel_data.sheetnames:
            self.LogScreenshot.fLogScreenshot(message=f"'Utility Summary' tab is present in Downloaded Complete Excel Report",
                                          pass_=True, log=True, screenshot=False)
        else:
            raise Exception("'Utility Summary' tab is missing in Complete Excel Report")
        
        self.LogScreenshot.fLogScreenshot(message=f"*****Check Presence of expected columns in Utility Summary Tab*****",
                                          pass_=True, log=True, screenshot=False)
        sourcefile = pd.read_excel(f'{source_template[0]}', sheet_name='ExpectedUtilitySummary')
        actualexcel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='Utility Summary', skiprows=3)
        
        source_col_val = [item for item in list(sourcefile.columns.values) if str(item) != 'nan']
        compex_col_val = [item for item in list(actualexcel.columns.values) if str(item) != 'nan']
        # Removing the unnamed column from the downloaded report. This column got created for 'Back To TOC' option
        compex_col_val.pop(-1)

        if source_col_val == compex_col_val:
            self.LogScreenshot.fLogScreenshot(message=f"Expected column names are present in Complete Excel Report. Column names are: {compex_col_val}",
                                          pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f"Expected column names are not present in Complete Excel Report."
                                          f"Source Excel Column names are : {source_col_val},\n Complete Excel Column names are : {compex_col_val}",
                                          pass_=False, log=True, screenshot=False)
            raise Exception("Expected column names are not present in Complete Excel Report")
    
    def qol_utility_summary_validation(self, webexcel_filename, excel_filename, util_filepath, word_filename):
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Started*****",
                                          pass_=True, log=True, screenshot=False)
        source_template = self.get_util_source_template(util_filepath, 'NewImportLogic')
        
        # QOL Report sheet comparison with Expected results
        self.LogScreenshot.fLogScreenshot(message=f"*****QOL Report sheet Comparison*****",
                                          pass_=True, log=True, screenshot=False)
        sourcefile1 = pd.read_excel(f'{source_template[0]}', sheet_name='ExpectedReportData')
        webexcel1 = pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name='QOL Report', skiprows=3)
        actualexcel1 = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='QOL Report', skiprows=4)
        
        cols1 = list(sourcefile1.columns.values)

        sourcefile_col1 = sourcefile1['Article Identifier(s)']
        actualexcel_col1 = actualexcel1['Article Identifier(s)']
        actualwebexcel_col1 = webexcel1['Article Identifier(s)']

        sourcefile_col1 = [item for item in sourcefile_col1 if str(item) != 'nan']
        actualexcel_col1 = [item for item in actualexcel_col1 if str(item) != 'nan']
        actualwebexcel_col1 = [item for item in actualwebexcel_col1 if str(item) != 'nan']

        if sourcefile_col1 == sorted(sourcefile_col1) and actualexcel_col1 == sorted(actualexcel_col1) and actualwebexcel_col1 == sorted(actualwebexcel_col1):
            self.LogScreenshot.fLogScreenshot(message=f"From 'QOL Report' sheet, contents in column 'Article Identifier(s)' are in sorted order",
                                            pass_=True, log=True, screenshot=False)
            for col in cols1:
                source_col1 = sourcefile1[col]
                actual_excel_col1 = actualexcel1[col]
                actual_webexcel_col = webexcel1[col]

                source_col1 = [item for item in source_col1 if str(item) != 'nan']
                actual_excel_col1 = [item for item in actual_excel_col1 if str(item) != 'nan']
                actual_webexcel_col = [item for item in actual_webexcel_col if str(item) != 'nan']

                if len(actual_excel_col1) == len(source_col1) == len(actual_webexcel_col) and actual_excel_col1 == source_col1 == actual_webexcel_col:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are matching between Source Template, Complete Excel and WebExcel Report"
                                            f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_excel_col1)},\n Web_Excel Elements Length {len(actual_webexcel_col)}",
                                            pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are not matching. "
                                            f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_excel_col1)},\n Web_Excel Elements Length {len(actual_webexcel_col)}"
                                            f"Source Template column value is {source_col1},\n Actual excel file column value is {actual_excel_col1},\n Actual webexcel file column value is {actual_webexcel_col}",
                                            pass_=False, log=True, screenshot=False)
                    raise Exception("Contents are not matching between Source template, Complete Excel and WebExcel report")
        else:
            raise Exception("From 'QOL Report' sheet, contents in column 'Article Identifier(s)' are not in Sorted order")
        
        # Utility Summary sheet comparison with Expected results
        self.LogScreenshot.fLogScreenshot(message=f"*****Utility Summary Sheet Comparison*****",
                                          pass_=True, log=True, screenshot=False)
        sourcefile = pd.read_excel(f'{source_template[0]}', sheet_name='ExpectedUtilitySummary')
        actualexcel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='Utility Summary', skiprows=3)
        
        cols = list(sourcefile.columns.values)

        sourcefile_col = sourcefile['Short Reference']
        actualexcel_col = actualexcel['Short Reference']

        sourcefile_col = [item for item in sourcefile_col if str(item) != 'nan']
        actualexcel_col = [item for item in actualexcel_col if str(item) != 'nan']

        if sourcefile_col == sorted(sourcefile_col) and actualexcel_col == sorted(actualexcel_col):
            self.LogScreenshot.fLogScreenshot(message=f"From 'Utility Summary' sheet, contents in column 'Short Reference' are in sorted order",
                                            pass_=True, log=True, screenshot=False)
            for col in cols:
                source_col = sourcefile[col]
                actual_col = actualexcel[col]

                source_col = [item for item in source_col if str(item) != 'nan']
                actual_col = [item for item in actual_col if str(item) != 'nan']

                if len(source_col) == len(actual_col) and source_col == actual_col:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are matching between Source Template and Complete Excel Report"
                                          f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_col)}",
                                          pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are not matching. "
                                          f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_col)}"
                                          f"Source Template column value is {source_col},\n Actual downloaded file column value is {actual_col}",
                                          pass_=False, log=True, screenshot=False)
                    raise Exception("Contents are not matching between Source template and Complete Excel report")
        else:
            raise Exception("From 'Utility Summary' sheet, contents in column 'Short Reference' are not in Sorted order")

        # Word report content comparison for Utility Table
        self.LogScreenshot.fLogScreenshot(message=f"*****Word Report Utility Table Comparison*****",
                                          pass_=True, log=True, screenshot=False)
        docs = docx.Document(f'ActualOutputs//{word_filename}')
        try:
            table = docs.tables[6]
            data = [[cell.text for cell in row.cells] for row in table.rows]
            df_word = pd.DataFrame(data)

            # Using count variable to loop over columns in word document
            count = 0
            # df_word.values[0] will give the list of column names from the table in Word document
            for col_name in df_word.values[0]:
                if col_name in sourcefile.columns.values and col_name in actualexcel.columns.values:
                    source = sourcefile[col_name]
                    actual = actualexcel[col_name]
                    word = []
                    for row in docs.tables[6].rows:
                        word.append(row.cells[count].text)
                    
                    source = [item for item in source if str(item) != 'nan']
                    actual = [item for item in actual if str(item) != 'nan']
                    word.pop(0)

                    # Validating the sorted order for 'Short Reference' column from Word Report -> Utility Table
                    if col_name == 'Short Reference':
                        if word == sorted(word):
                            self.LogScreenshot.fLogScreenshot(message=f"From Word Report -> 'Utility Table' section, contents in column 'Short Reference' are in sorted order",
                                            pass_=True, log=True, screenshot=False)
                        else:
                            raise Exception("From Word Report -> 'Utility Table' section, contents in column 'Short Reference' are not in sorted order")

                    if len(source) == len(actual) == len(word) and source == actual == word:
                        self.LogScreenshot.fLogScreenshot(message=f"Contents in Column '{col_name}' are matching between Source Excel, Complete Excel and Word Reports.\n"
                                    f"Source Excel Elements Length: {len(source)},\n Complete Excel Elements Length: {len(actual)},\n Word Elements Length: {len(word)}\n",
                                    # f"Source Excel Elements: {source}\n Excel Elements: {actual}\n Word Elements: {word}",
                            pass_=True, log=True, screenshot=False)
                    else:
                        self.LogScreenshot.fLogScreenshot(message=f"Contents in Column '{col_name}' are not matching between Source Excel, Complete Excel and Word Reports.\n"
                                    f"Source Excel Elements Length: {len(source)},\n Excel Elements Length: {len(actual)},\n Word Elements Length: {len(word)}\n"
                                    f"Source Excel Elements: {source},\n Complete Excel Elements: {actual},\n Word Elements: {word}",
                            pass_=False, log=True, screenshot=False)
                        raise Exception("Elements are not matching between Source Excel, Complete Excel and Word Reports")
                    count += 1
                else:
                    raise Exception("Column names are not matching between Source Excel, Complete Excel and Word Reports")
        except Exception:
                raise Exception("Error in Word report content validation")
        
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Completed*****",
                                          pass_=True, log=True, screenshot=False)

    def qol_utility_summary_validation_old_imports(self, webexcel_filename, excel_filename, util_filepath):
        self.LogScreenshot.fLogScreenshot(message=f"*****Validation as per the OLD Import logic*****",
                                          pass_=True, log=True, screenshot=False)
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Started*****",
                                          pass_=True, log=True, screenshot=False)
        source_template = self.get_util_source_template(util_filepath, 'OldImportLogic')
        
        # QOL Report sheet comparison with Expected results
        self.LogScreenshot.fLogScreenshot(message=f"*****QOL Report sheet Comparison*****",
                                          pass_=True, log=True, screenshot=False)
        sourcefile1 = pd.read_excel(f'{source_template[0]}', sheet_name='ExpectedReportData')
        webexcel1 = pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name='QOL Report', skiprows=3)
        actualexcel1 = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='QOL Report', skiprows=4)
        
        cols1 = list(sourcefile1.columns.values)

        sourcefile_col1 = sourcefile1['Article Identifier(s)']
        actualexcel_col1 = actualexcel1['Article Identifier(s)']
        actualwebexcel_col1 = webexcel1['Article Identifier(s)']

        sourcefile_col1 = [item for item in sourcefile_col1 if str(item) != 'nan']
        actualexcel_col1 = [item for item in actualexcel_col1 if str(item) != 'nan']
        actualwebexcel_col1 = [item for item in actualwebexcel_col1 if str(item) != 'nan']

        if sourcefile_col1 == sorted(sourcefile_col1) and actualexcel_col1 == sorted(actualexcel_col1) and actualwebexcel_col1 == sorted(actualwebexcel_col1):
            self.LogScreenshot.fLogScreenshot(message=f"From 'QOL Report' sheet, contents in column 'Article Identifier(s)' are in sorted order",
                                            pass_=True, log=True, screenshot=False)
            for col in cols1:
                source_col1 = sourcefile1[col]
                actual_excel_col1 = actualexcel1[col]
                actual_webexcel_col = webexcel1[col]

                source_col1 = [item for item in source_col1 if str(item) != 'nan']
                actual_excel_col1 = [item for item in actual_excel_col1 if str(item) != 'nan']
                actual_webexcel_col = [item for item in actual_webexcel_col if str(item) != 'nan']

                if len(actual_excel_col1) == len(source_col1) == len(actual_webexcel_col) and actual_excel_col1 == source_col1 == actual_webexcel_col:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are matching between Source Template, Complete Excel and WebExcel Report"
                                            f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_excel_col1)},\n Web_Excel Elements Length {len(actual_webexcel_col)}",
                                            pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are not matching. "
                                            f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_excel_col1)},\n Web_Excel Elements Length {len(actual_webexcel_col)}"
                                            f"Source Template column Value is {source_col1},\n Actual excel file column value is {actual_excel_col1},\n Actual webexcel file column value is {actual_webexcel_col}",
                                            pass_=False, log=True, screenshot=False)
                    raise Exception("Contents are not matching between Source template, Complete Excel and WebExcel report")
        else:
            raise Exception("From 'QOL Report' sheet, contents in column 'Article Identifier(s)' are not in Sorted order")
        
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Completed*****",
                                          pass_=True, log=True, screenshot=False)

    def econ_utility_summary_validation(self, webexcel_filename, excel_filename, util_filepath, word_filename):
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Started*****",
                                          pass_=True, log=True, screenshot=False)
        source_template = self.get_util_source_template(util_filepath, 'NewImportLogic')
        
        # ECON Report sheet comparison with Expected results
        self.LogScreenshot.fLogScreenshot(message=f"*****ECON Report sheet Comparison*****",
                                          pass_=True, log=True, screenshot=False)
        sourcefile = pd.read_excel(f'{source_template[0]}', sheet_name='ExpectedReportData')
        webexcel = pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name='CEA CUA Report', skiprows=3)
        actualexcel = pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name='CEA CUA Report', skiprows=4)
        
        cols = list(sourcefile.columns.values)

        sourcefile_col = sourcefile['Article Identifier(s)']
        actualexcel_col = actualexcel['Article Identifier(s)']
        actualwebexcel_col = webexcel['Article Identifier(s)']

        sourcefile_col = [item for item in sourcefile_col if str(item) != 'nan']
        actualexcel_col = [item for item in actualexcel_col if str(item) != 'nan']
        actualwebexcel_col = [item for item in actualwebexcel_col if str(item) != 'nan']

        if sourcefile_col == sorted(sourcefile_col) and actualexcel_col == sorted(actualexcel_col) and actualwebexcel_col == sorted(actualwebexcel_col):
            self.LogScreenshot.fLogScreenshot(message=f"From 'CEA CUA Report' sheet, contents in column 'Article Identifier(s)' are in sorted order",
                                            pass_=True, log=True, screenshot=False)
            for col in cols:
                source_col = sourcefile[col]
                actual_excel_col = actualexcel[col]
                actual_webexcel_col = webexcel[col]

                source_col = [item for item in source_col if str(item) != 'nan']
                actual_excel_col = [item for item in actual_excel_col if str(item) != 'nan']
                actual_webexcel_col = [item for item in actual_webexcel_col if str(item) != 'nan']

                if len(actual_excel_col) == len(source_col) == len(actual_webexcel_col) and actual_excel_col == source_col == actual_webexcel_col:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are matching between Source Template, Complete Excel and WebExcel Report"
                                            f"Source Template Elements Length {len(source_col)},\n Complete Excel Elements Length {len(actual_excel_col)},\n Web_Excel Elements Length {len(actual_webexcel_col)}",
                                            pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are not matching. "
                                            f"Source Template Elements Length {len(source_col)},\n Complete Excel Elements Length {len(actual_excel_col)},\n Web_Excel Elements Length {len(actual_webexcel_col)}"
                                            f"Source Template column Value is {source_col},\n Actual excel file column value is {actual_excel_col},\n Actual webexcel file column value is {actual_webexcel_col}",
                                            pass_=False, log=True, screenshot=False)
                    raise Exception("Contents are not matching between Source template, Complete Excel and WebExcel Report")
        else:
            raise Exception("From 'CEA CUA Report' sheet, contents in column 'Article Identifier(s)' are not in Sorted order")
        
        # Word report content comparison with Source template
        self.LogScreenshot.fLogScreenshot(message=f"*****ECON Word Report Comparison with Source template*****",
                                          pass_=True, log=True, screenshot=False)
        # Index of Table number 6 is : 5. Starting point for word table content comparison
        table_count = 5
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
                    if col_name in sourcefile.columns.values and col_name in actualexcel.columns.values and col_name in webexcel.columns.values:
                        source = sourcefile[col_name]
                        actualex = actualexcel[col_name]
                        actwebex = webexcel[col_name]
                        word = []
                        for row in docs.tables[table_count].rows:
                            word.append(row.cells[count].text)
                        
                        source = [item for item in source if str(item) != 'nan']
                        actualex = [item for item in actualex if str(item) != 'nan']
                        actwebex = [item for item in actwebex if str(item) != 'nan']
                        word.pop(0)

                        if len(source) == len(actualex) == len(word) == len(actwebex) and source == actualex == word == actwebex:
                            self.LogScreenshot.fLogScreenshot(message=f"Contents in Column '{col_name}' are matching between Source Excel, Complete Excel, Web-Excel and Word Reports.\n"
                                        f"Source Excel Elements Length: {len(source)},\n Complete Excel Elements Length: {len(actualex)},\n Web-Excel Elements Length: {len(actwebex)},\n Word Elements Length: {len(word)}\n",
                                        # f"Source Excel Elements: {source}\n Excel Elements: {actual}\n Word Elements: {word}",
                                pass_=True, log=True, screenshot=False)
                        else:
                            self.LogScreenshot.fLogScreenshot(message=f"Contents in Column '{col_name}' are not matching between Source Excel, Complete Excel, Web-Excel and Word Reports.\n"
                                        f"Source Excel Elements Length: {len(source)},\n Excel Elements Length: {len(actualex)},\n Web-Excel Elements Length: {len(actwebex)},\n Word Elements Length: {len(word)}\n"
                                        f"Source Excel Elements: {source},\n Complete Excel Elements: {actualex},\n Web-Excel Elements: {actwebex},\n Word Elements: {word}",
                                pass_=False, log=True, screenshot=False)
                            raise Exception("Elements are not matching between Source Excel, Complete Excel and Word Reports")
                        count += 1
                    else:
                        raise Exception("Column names are not matching between Source Excel, Complete Excel, WebExcel and Word Reports")
        except Exception:
                raise Exception("Error in Word report content validation")
        
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Completed*****",
                                          pass_=True, log=True, screenshot=False)

    def econ_utility_summary_validation_old_imports(self, webexcel_filename, excel_filename, util_filepath, word_filename):
        self.LogScreenshot.fLogScreenshot(message=f"*****Validation as per the OLD Import logic*****",
                                          pass_=True, log=True, screenshot=False)
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Started*****",
                                          pass_=True, log=True, screenshot=False)
        source_template = self.get_util_source_template(util_filepath, 'OldImportLogic')
        
        # ECON Report sheet comparison with Expected results
        self.LogScreenshot.fLogScreenshot(message=f"*****ECON Report sheet Comparison*****",
                                          pass_=True, log=True, screenshot=False)
        sourcefile1 = pd.read_excel(f'{source_template[0]}', sheet_name='ExpectedReportData')
        webexcel1 =  pd.concat(pd.read_excel(f'ActualOutputs//{webexcel_filename}', sheet_name=None, skiprows=3), ignore_index=True)
        actualexcel1 =  pd.concat(pd.read_excel(f'ActualOutputs//{excel_filename}', sheet_name=None, skiprows=4), ignore_index=True)
        
        cols1 = list(sourcefile1.columns.values)

        sourcefile_col1 = sourcefile1['Article Identifier(s)']
        actualexcel_col1 = actualexcel1['Article Identifier(s)']
        actualwebexcel_col1 = webexcel1['Article Identifier(s)']

        sourcefile_col1 = [item for item in sourcefile_col1 if str(item) != 'nan']
        actualexcel_col1 = [item for item in actualexcel_col1 if str(item) != 'nan']
        actualwebexcel_col1 = [item for item in actualwebexcel_col1 if str(item) != 'nan']

        if sourcefile_col1 == sorted(sourcefile_col1) and actualexcel_col1 == sorted(actualexcel_col1) and actualwebexcel_col1 == sorted(actualwebexcel_col1):
            self.LogScreenshot.fLogScreenshot(message=f"Contents in column 'Article Identifier(s)' are in sorted order",
                                            pass_=True, log=True, screenshot=False)
            for col in cols1:
                source_col1 = sourcefile1[col]
                actual_excel_col1 = actualexcel1[col]
                actual_webexcel_col1 = webexcel1[col]

                source_col1 = [item for item in source_col1 if str(item) != 'nan']
                actual_excel_col1 = [item for item in actual_excel_col1 if str(item) != 'nan']
                actual_webexcel_col1 = [item for item in actual_webexcel_col1 if str(item) != 'nan']

                if len(actual_excel_col1) == len(source_col1) == len(actual_webexcel_col1) and actual_excel_col1 == source_col1 == actual_webexcel_col1:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are matching between Source Template, Complete Excel and WebExcel Report"
                                            f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_excel_col1)},\n Web_Excel Elements Length {len(actual_webexcel_col1)}",
                                            pass_=True, log=True, screenshot=False)
                else:
                    self.LogScreenshot.fLogScreenshot(message=f"Contents in column '{col}' are not matching. "
                                            f"Source Template Elements Length {len(source_col1)},\n Complete Excel Elements Length {len(actual_excel_col1)},\n Web_Excel Elements Length {len(actual_webexcel_col1)}"
                                            f"Source Column Value is {source_col1},\n Actual excel file column value is {actual_excel_col1},\n Actual webexcel file column value is {actual_webexcel_col1}",
                                            pass_=False, log=True, screenshot=False)
                    raise Exception("Contents are not matching between Source template, Complete Excel and WebExcel Report")
        else:
            raise Exception("Contents in column 'Article Identifier(s)' are not in Sorted order")
        
        self.LogScreenshot.fLogScreenshot(message=f"*****Content validation between Extraction Template, Complete Excel Report and Complete Word Report for Utility Summary Completed*****",
                                          pass_=True, log=True, screenshot=False)
