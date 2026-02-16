import configparser

class confighandler():

    def __init__(self, configfile):
        self.configfile=configfile
        self.config = configparser.ConfigParser()


    def configreadout(self):
        self.config.read(self.configfile)
        #print(self.config.sections())
        return self.config

    def writevalue(self, config):
        with open(self.configfile, 'w') as configfile:
            config.write(configfile)