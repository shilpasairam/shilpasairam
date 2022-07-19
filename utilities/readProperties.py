import configparser

config = configparser.RawConfigParser()
config.read("D:\\Work\\Cytel\\PytestFrameworkLiveSLR\\Configurations\\config.ini")


class ReadConfig:

    # static method helps you read the function in another file without instantiating the class
    @staticmethod
    def getApplicationURL():
        # return f"https://pse-liveslr-ng-testing.azurewebsites.net/"
        return f"https://slr-stage.livehta.com/"

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
        # OR = config.get('commonInfo', 'OR')
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

    # Read input advisor authentication data from config.ini
    @staticmethod
    def getAuthInput(authinput):
        input = config.get(ReadConfig.getEnvironmenttype(), authinput)
        return input

    # Get test data file path for LiveSLR
    @staticmethod
    def getslrtestdata():
        # populationdata = config.get('commonInfo', 'slrpopulationdata')
        populationdata = config.get('commonInfo', 'stagingslrpopdata')
        return populationdata

    # Get test data file path for LiveNMA
    @staticmethod
    def getnmatestdata():
        populationdata = config.get('commonInfo', 'staginglivenmadata')
        return populationdata

    # only used for default visual. for future development only using getwrapperTD instead
    @staticmethod
    def get_sheetname_as_per_slrtype(value):
        sheetnames = []
        if value == "Interventional":
            sheetnames = ['Interventional Report']
        elif value == "Economic":
            sheetnames = ['CEA CUA Report', 'BIM Report', 'CostHCRU Report', 'Other Report']
        elif value == "Quality of Life":
            sheetnames = ['QOL Report']
        elif value == "Real-world Evidence":
            sheetnames = ['RWE Report']
        return sheetnames
