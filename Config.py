import configparser
from werkzeug.datastructures import ImmutableMultiDict

class Config:
    def __init__(self, config_filename: str):
        config = configparser.ConfigParser()
        config.read(config_filename)
        self.excel_worksheet_name = config.get('excel', 'worksheet_name')

def update_config(section: str, form: ImmutableMultiDict, out_filename: str):
    config = configparser.ConfigParser()
    config[section] = form
    with open(out_filename, 'w') as configfile:
        config.write(configfile)