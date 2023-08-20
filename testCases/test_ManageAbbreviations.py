import os
import pytest
import pandas as pd
import time
from datetime import date,timedelta
from Pages.Base import Base
from Pages.ExtendedBasePage import ExtendedBase
from Pages.LoginPage import LoginPage
from Pages.OpenLiveSLRPage import LiveSLRPage
from Pages.ManageAbbreviationsPage import ManageAbbreviations
from utilities.logScreenshot import cLogScreenshot
from utilities.readProperties import ReadConfig
from Pages.SLRReportPage import SLRReport
from selenium.webdriver.common.by import By

@pytest.mark.usefixtures("init_driver")
class Test_ManageAbbreviationPage:
    username = ReadConfig.getUserName()
    password = ReadConfig.getPassword()

    @pytest.mark.C42203
    def test_manageAbbreviationPage_functionalities(self,extra,env,request,caseid):
        base_Url = ReadConfig.getPortalURL(env)
        basefile = ReadConfig.getnononcologybasefile("nononcology_basefile")
        base = Base(self.driver, extra)         
        # Instantiate the logScreenshot class
        LogScreenshot = cLogScreenshot(self.driver, extra)
        # Creating object of ExtendedBase class
        exbase = ExtendedBase(self.driver, extra)
        # Creating object of loginpage class
        loginPage = LoginPage(self.driver, extra)
        # Creating object of ManageAbbreviationsPage class
        mngAbbr = ManageAbbreviations(self.driver, extra) 
        # Creating object of slrreport class
        slrreport = SLRReport(self.driver, extra)
        request.node._tcid = caseid
        request.node._title = "Verifying the Manage Abbreviation page functionalities"
        loginPage.driver.get(base_Url)
        loginPage.complete_portal_login(self.username, self.password, "launch_live_slr", "Cytel LiveSLR", base_Url, env)
        base.presence_of_admin_page_option("manage_abbreviation", env)
        #base.presence_of_admin_page_option("manage_abbr_link", env)
        base.go_to_page("manage_abbreviation", env)
        base.isvisible("mng_abbr_home_page_lbl",env,"")
        upload_success_status_msg = "Abbreviations updated successfully"
        upload_failure_status_msg = "error: The file extension should belong to this list: [.xls, .xlsx]"
        filepath = exbase.get_testdata_filepath(basefile, "livehta_3750_manage_abbreviation")
        sheetsInFile = pd.ExcelFile(filepath)
        for sheet in range(0,len(sheetsInFile.sheet_names)):
            mngAbbrpopulationData = pd.read_excel(filepath, sheet_name=f'population{sheet}data')
            populations,Study_Types,Abbreviation,Definition,expectedFileName = mngAbbr.readDataFromDataFile(mngAbbrpopulationData)
            for p in range(len(populations)):
                time.sleep(1)
                populationVal = populations[p]
                populationTxt = base.selectbyvisibletext("mng_abbr_population_dd", populationVal, env)
                base.evidenceOfDataInUI(populationTxt)
                for s in range(0,len(Study_Types)):
                    time.sleep(1)
                    studyType = base.selectbyvisibletext("mng_abbr_Std_Type_dd", Study_Types[s], env)
                    base.evidenceOfDataInUI(studyType) 
                    base.isenabled("mng_abbr_download_btn",env)
                    slrreport.generate_download_report("mng_abbr_download_btn",env)
                    time.sleep(3)
                    fileName = slrreport.get_latest_filename()
                    slrreport.get_and_validate_downloaded_fileName(fileName,expectedFileName[s])
                    pathOfFile = os.getcwd()+f'\ActualOutputs\{fileName}'
                    mngAbbr.addingValuesToDownloadedFile(pathOfFile,fileName,Abbreviation,Definition)
                    time.sleep(2)
                    jscmd = ReadConfig.get_remove_att_JScommand(16, 'hidden')
                    base.jsclick_hide(jscmd)
                    # uploading updated excel file to portal
                    time.sleep(2)
                    mngAbbr.uploadFile(pathOfFile,"File upload is success ",upload_success_status_msg,env)
                    # uploading invalid excel file to portal
                    path=os.getcwd()+f"\\Testdata\\Non_Oncology\\DataFiles\\ManageAbbreviations\\textToUpload.txt"
                    mngAbbr.uploadFile(path,"File upload is Failed",upload_failure_status_msg,env)
                    base.refreshpage()
                    time.sleep(2)
                    base.presence_of_element("mng_abbr_population_dd",env)
                    base.selectbyvisibletext("mng_abbr_population_dd", populationVal, env)


                


                    



