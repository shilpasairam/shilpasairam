import math
import os
import time
import openpyxl
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from selenium.webdriver.support.wait import WebDriverWait

from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from utilities.readProperties import ReadConfig
from utilities.customLogger import LogGen
from utilities.logScreenshot import cLogScreenshot
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select
from operator import itemgetter

class ManageAbbreviations(Base):
    def __init__(self,driver,extra):
        super().__init__(driver,extra)
        self.extra = extra
        # Instantiate the Base class
        self.base = Base(self.driver, self.extra)
        # Creating object of ExtendedBase class
        self.exbase = ExtendedBase(self.driver, extra)        
        # Instantiate the logger class
        self.logger = LogGen.loggen()
        # Instantiate the logScreenshot class
        self.LogScreenshot = cLogScreenshot(self.driver, self.extra)
        # Instantiate webdriver wait class
        self.wait = WebDriverWait(driver, 10)

    def go_to_mng_abbreviation_page(self,env):
        self.base.click("manage_abbr_link",env)

    def verifyFileUploadStatus(self,statusMsg,expected_upload_status_text,locator,env):
        try:
            actual_upload_status_text = self.get_status_text(locator, env)
            time.sleep(1)
            if actual_upload_status_text == expected_upload_status_text:
                self.LogScreenshot.fLogScreenshot(message=f'{statusMsg}',
                                                      pass_=True, log=True, screenshot=True)
            else:
                self.LogScreenshot.fLogScreenshot(message=f'Unable to find status message while uploading '
                                                              f'Actual status message is {actual_upload_status_text} '
                                                              f'and Expected status message is '
                                                              f'{expected_upload_status_text}',
                                                      pass_=False, log=True, screenshot=True)
                raise Exception("Unable to find status message during Extraction file uploading")
            
        except Exception:
                raise Exception("Error while uploading")   

    def uploadFile(self,path,statusMsg,expected_text,env):
        self.base.input_text("mng_abbr_file_upload_loc", path,env)
        time.sleep(3)
        self.base.isenabled("mng_abbr_upload_abbr_btn",env)
        self.base.click("mng_abbr_upload_abbr_btn",env)
        self.verifyFileUploadStatus(statusMsg,expected_text,"file_status_popup_text",env)

    def verifyDataFrameUpdatedOrNot(self,count1,count2):
        if count2 > count1:
            self.LogScreenshot.fLogScreenshot(message=f'Before updating Abbreviation and Definition in template number of rows in file {count1}, after updating Abbreviation and Definition in file - number of rows in file {count2}',
                                                      pass_=True, log=True, screenshot=False)
        else:
            self.LogScreenshot.fLogScreenshot(message=f'Failed to add data to the dataframe{count1} and {count2}',
                                                      pass_=False, log=True, screenshot=False)
        

    def addingValuesToDownloadedFile(self,pathOfFile,fileName,Abbreviation,Definition):
        pathOfFile = os.getcwd()+f'\ActualOutputs\{fileName}'
        dataFrame = pd.read_excel(pathOfFile)
        dataFrame1 = pd.DataFrame(dataFrame)
        noOfRowsBeforeValsAdd = len(dataFrame1.index)
        newData = pd.Series((Abbreviation, Definition), index=dataFrame1.columns)
        updatedDF = dataFrame1.append(newData,ignore_index=True)
        noOfRowsAfterValsAdd = len(updatedDF.index)
        self.verifyDataFrameUpdatedOrNot(noOfRowsBeforeValsAdd,noOfRowsAfterValsAdd)
        self.LogScreenshot.fLogScreenshot(message=f"Data present in the updated Abbreviation template file {updatedDF}",
                                             pass_=True, log=True, screenshot=False)
        updatedDF.to_excel(pathOfFile,index=False)

    def readDataFromDataFile(self,sheet):
        # converting file data in to dictionary
        mngAbbrKeyVals = sheet.to_dict()
        # taking the values of populations key
        populations = mngAbbrKeyVals['populations'].values()
        populations = [txt for txt in populations if str(txt)!='nan']
        Study_Types = mngAbbrKeyVals['Study_Types'].values()
        Study_Types = [txt for txt in Study_Types if str(txt)!='nan']
        Abbreviation = mngAbbrKeyVals['Abbreviation'].values()
        Abbreviation = [txt for txt in Abbreviation if str(txt)!='nan']
        Definition = mngAbbrKeyVals['Definition'].values()
        Definition = [txt for txt in Definition if str(txt)!='nan']
        expectedFileName = mngAbbrKeyVals['expectedFileName'].values()
        expectedFileName = [txt for txt in expectedFileName if str(txt)!='nan']
        
        return populations,Study_Types,Abbreviation,Definition,expectedFileName
    
    
            

    
    



    
        
    

    