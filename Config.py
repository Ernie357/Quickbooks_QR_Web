import configparser
from werkzeug.datastructures import ImmutableMultiDict

class Config:
    def __init__(self, config_filename: str):
        config = configparser.ConfigParser()
        config.read(config_filename)
        self.excel = dict(config['Excel']) if 'Excel' in config else {}
        self.csv = dict(config['CSV']) if 'CSV' in config else {}

def update_config(section: str, form: ImmutableMultiDict, out_filename: str):
    config = configparser.ConfigParser()
    config[section] = form
    with open(out_filename, 'w') as configfile:
        config.write(configfile)