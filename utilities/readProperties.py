import configparser
import os

config = configparser.RawConfigParser()
config.read(os.getcwd()+"\\Configurations\\config.ini")


class ReadConfig:

    # static method helps you read the function in another file without instantiating the class
    @staticmethod
    def getApplicationURL():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return f"https://pse-portal-testing.azurewebsites.net/"
        elif env == 'staging':
            return f"https://portal-stage.livehta.com"
        elif env == 'production':
            return f"https://portal.livehta.com/"

    @staticmethod
    def getappversionfilepath():
        appversion = config.get('commonInfo', 'App_version_data')
        return appversion
    
    @staticmethod
    def getUserName():
        username = config.get('commonInfo', 'username')
        return username

    @staticmethod
    def getPassword():
        password = config.get('commonInfo', 'password')
        return password

    @staticmethod
    def getORFilePath():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            OR = config.get('commonInfo', 'OR_testingconfig')
        elif env == 'staging':
            OR = config.get('commonInfo', 'OR_stagingconfig')
        elif env == 'production':
            OR = config.get('commonInfo', 'OR_prodconfig')
        return OR

    # @staticmethod
    # def getEnvironmenttype():
    #     environmenttype = config.get('commonInfo', 'environmenttype')
    #     return environmenttype

    # @staticmethod
    # def getEnvironment():
    #     environment = config.get('commonInfo', 'environment')
    #     return environment

    # # Read input advisor authentication data from config.ini
    # @staticmethod
    # def getAuthInput(authinput):
    #     input = config.get(ReadConfig.getEnvironmenttype(), authinput)
    #     return input

    # Get test data file path for LiveSLR
    @staticmethod
    def getslrtestdata():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            populationdata = config.get('commonInfo', 'slrpopulationdata_testing')
        elif env == 'staging':
            populationdata = config.get('commonInfo', 'slrpopulationdata_staging')
        elif env == 'production':
            populationdata = config.get('commonInfo', 'slrpopulationdata_prod')
        return populationdata

    # Get test data file path for LiveNMA
    @staticmethod
    def getnmatestdata():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            populationdata = config.get('commonInfo', 'testinglivenmadata')
        elif env == 'staging':
            populationdata = config.get('commonInfo', 'staginglivenmadata')
        return populationdata

    # Get file containing data for QOL Utility
    @staticmethod
    def getutilityoutcome_QOL_data():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return config.get('commonInfo', 'utilityoutcome_QOL_testing')
        elif env == 'staging':
            return config.get('commonInfo', 'utilityoutcome_QOL_staging')
        elif env == 'production':
            return config.get('commonInfo', 'utilityoutcome_QOL_prod')
    
    # Get file containing data for ECON Utility
    @staticmethod
    def getutilityoutcome_ECON_data():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return config.get('commonInfo', 'utilityoutcome_ECON_testing')
        elif env == 'staging':
            return config.get('commonInfo', 'utilityoutcome_ECON_staging')
    
    # Get file containing data for Admin Page actions
    @staticmethod
    def getimportpublicationsdata():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return config.get('commonInfo', 'importpublicationsdata_testing')
        elif env == 'staging':
            return config.get('commonInfo', 'importpublicationsdata_staging')
    
    # Get file containing data for Manage QA Data page actions
    @staticmethod
    def getmanageqadatapath():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return config.get('commonInfo', 'manageqadata_testing')
        elif env == 'staging':
            return config.get('commonInfo', 'manageqadata_staging')
        elif env == 'production':
            return config.get('commonInfo', 'manageqadata_prod')
    
    # Get data file for manage population page
    @staticmethod
    def getmanagepopdatafilepath():
        template = config.get('commonInfo', 'managepopulationdata')
        return template
    
    # Get file containing data for Manageupdates page actions
    @staticmethod
    def getmanageupdatesdata():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return config.get('commonInfo', 'manageupdatesdata_testing')
        elif env == 'staging':
            return config.get('commonInfo', 'manageupdatesdata_staging')

    # Get file containing data for LineofTherapy page actions
    @staticmethod
    def getmanagelotdata():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return config.get('commonInfo', 'managelotdata_testing')
        elif env == 'staging':
            return config.get('commonInfo', 'managelotdata_staging')

    # Get file containing data for PRISMAs page actions
    @staticmethod
    def getprismadata():
        env = config.get('commonInfo', 'environment')
        if env == 'test':
            return config.get('commonInfo', 'prismadata_testing')
        elif env == 'staging':
            return config.get('commonInfo', 'prismadata_staging')
    
    # Get JS command to hide
    @staticmethod
    def get_remove_att_JScommand(index, value):
        command = f"document.getElementsByTagName('input')[{index}].removeAttribute('{value}')"
        return command
    
    # Get JS command to hide
    @staticmethod
    def get_set_att_JScommand(index, name, value=None):
        command = f"document.getElementsByTagName('input')[{index}].setAttribute('{name}', '{value}')"
        return command

    # Get file containing data for Excluded Studies page actions
    @staticmethod
    def getexcludedstudiespath():
        env = config.get('commonInfo', 'environment')
        if env == 'production':
            return config.get('commonInfo', 'ExcludedStudiesdata_prod')
        else:
            return config.get('commonInfo', 'ExcludedStudiesdata')

    # Get file containing data for Excluded Studies - LiveSLR Tab actions
    @staticmethod
    def getexcludedstudiesliveslrpath():
        exstdy_liveslr = config.get('commonInfo', 'ExcludedStudies_liveSLR_data')
        return exstdy_liveslr

    @staticmethod
    def getTestdata(wrapperfile):
        TestData = config.get('commonInfo', wrapperfile)
        return TestData
