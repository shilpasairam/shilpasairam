import configparser

config = configparser.RawConfigParser()
config.read("D:\\VersionControl\\pse.autotest\\Configurations\\config.ini")


class ReadConfig:

    # static method helps you read the function in another file without instantiating the class
    @staticmethod
    def getApplicationURL():
        env = config.get('commonInfo','environment')
        if env == 'test':
            return f"https://pse-portal-testing.azurewebsites.net/"
        elif env == 'staging':
            return f"https://pse-portal-staging.azurewebsites.net/"

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
        env = config.get('commonInfo','environment')
        if env == 'test':
            OR = config.get('commonInfo', 'OR')
        elif env == 'staging':
            OR = config.get('commonInfo', 'OR_stagingconfig')
        return OR

    @staticmethod
    def getEnvironmenttype():
        environmenttype = config.get('commonInfo', 'environmenttype')
        return environmenttype

    @staticmethod
    def getEnvironment():
        environment = config.get('commonInfo', 'environment')
        return environment

    # # Read input advisor authentication data from config.ini
    # @staticmethod
    # def getAuthInput(authinput):
    #     input = config.get(ReadConfig.getEnvironmenttype(), authinput)
    #     return input

    # Get test data file path for LiveSLR
    @staticmethod
    def getslrtestdata():
        env = config.get('commonInfo','environment')
        if env == 'test':
            populationdata = config.get('commonInfo', 'slrpopulationdata')
        elif env == 'staging':
            populationdata = config.get('commonInfo', 'stagingslrpopdata')
        return populationdata

    # Get test data file path for LiveNMA
    @staticmethod
    def getnmatestdata():
        env = config.get('commonInfo','environment')
        if env == 'test':
            populationdata = config.get('commonInfo', 'testinglivenmadata')
        elif env == 'staging':
            populationdata = config.get('commonInfo', 'staginglivenmadata')
        return populationdata

    # Get file containing data for Admin Page actions
    @staticmethod
    def getadminpagedata():
        admindata = config.get('commonInfo', 'adminpagedata')
        return admindata
    
    # Get data file for manage population page
    @staticmethod
    def getmanagepopdatafilepath():
        template = config.get('commonInfo', 'managepopulationdata')
        return template
    
    # Get JS command to hide
    @staticmethod
    def getJScommand():
        command = "document.getElementsByTagName('input')[16].removeAttribute('hidden')"
        return command