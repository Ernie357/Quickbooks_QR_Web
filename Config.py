import configparser
from werkzeug.datastructures import ImmutableMultiDict
from typing import Literal

# The below field dicts outline the form input names and form labels as key, value respectively.
# The form input names also access the ini config data to ensure consistency.
# Essentially, these dicts are the single source of truth for how the data is structured and accessed.

ExcelFieldKey = Literal[
    'worksheet_name',
    'invoice_number_name',
    'estate_number_name',
    'bill_to_name',
    'address_1_name',
    'address_2_name',
    'estate_of_name',
    '1st_run_name',
    '2nd_run_name',
    '3rd_run_name',
    'price_name',
    'qr_image_name',
    'qr_link_name',
]

excel_fields: dict[ExcelFieldKey, str] = {
    'worksheet_name': 'Excel Worksheet Name',
    'invoice_number_name': 'Invoice # Column Letter',
    'estate_number_name': 'Estate # Column Letter',
    'bill_to_name': 'Bill_To Column Letter',
    'address_1_name': 'Address 1 Column Letter',
    'address_2_name': 'Address 2 Column Letter',
    'estate_of_name': 'Estate Of Column Letter',
    '1st_run_name': '1st Run Column Letter',
    '2nd_run_name': '2nd Run Column Letter',
    '3rd_run_name': '3rd Run Column Letter',
    'price_name': 'Price Column Letter',
    'qr_image_name': 'QR Image Column Letter',
    'qr_link_name': 'QR Link Column Letter'
}

CsvFieldKey = Literal[
    'invoice_number_name',
    'customer_name',
    'invoice_date_name',
    'due_date_name',
    'memo_name',
    'item_amount_name',
    'item_description_name',
]

csv_fields: dict[CsvFieldKey, str] = {
    'invoice_number_name': 'Inv # Column Name',
    'customer_name': 'Customer Column Name',
    'invoice_date_name': 'Inv Date Column Name',
    'due_date_name': 'Due Date Column Name',
    'memo_name': 'Memo Column Name',
    'item_amount_name': 'Item Amount Column Name',
    'item_description_name': 'Item Description Column Name'
}

# alt+F9 in the template docx to see the fields
MailMergeKey = Literal[
    'Inv_Nbr',
    'Estate_No',
    'Bill_To',
    'Address_1',
    'Address_2',
    'Estate_Of',
    'M_1st_Run',
    'M_2nd_Run',
    'M_3rd_Run',
    'price',
    'QR_Image',
    'QR_Link'
]

mail_merge_map: dict[ExcelFieldKey, MailMergeKey] = {
    'invoice_number_name': 'Inv_Nbr',
    'estate_number_name': 'Estate_No',
    'bill_to_name': 'Bill_To',
    'address_1_name': 'Address_1',
    'address_2_name': 'Address_2',
    'estate_of_name': 'Estate_Of',
    '1st_run_name': 'M_1st_Run',
    '2nd_run_name': 'M_2nd_Run',
    '3rd_run_name': 'M_3rd_Run',
    'price_name': 'price',
    'qr_image_name': 'QR_Image',
    'qr_link_name': 'QR_Link'
}

merge_id_key: MailMergeKey = 'Inv_Nbr'
qr_image_key: MailMergeKey = 'QR_Image'

''' 
    This class holds ini config data in Python memory, ensuring it matches the structures above.
    This allows for the identification of Excel and CSV columns to be changed programmatically.
'''
class Config:
    def __init__(self, config_filename: str):
        config = configparser.ConfigParser()
        config.read(config_filename)
        self.excel_config = dict(config['Excel']) if 'Excel' in config else {}
        self.csv_config = dict(config['CSV']) if 'CSV' in config else {}

    def check_required_keys_exist(self):
        for required_key in excel_fields:
            if required_key not in self.excel_config:
                raise Exception(
                    f"Missing Excel config key: {required_key}"
                )
        for required_key in csv_fields:
            if required_key not in self.csv_config:
                raise Exception(
                    f"Missing CSV config key: {required_key}"
                )
            
    def check_keys_match(self):
        for k in self.excel_config:
            v = self.excel_config[k]
            if k not in excel_fields:
                raise Exception('Config key name mismatch for Excel section of config.')
            input_label = excel_fields[k]
            if v is None or v == '':
                raise Exception('One or more settings is blank for Excel config.')
            if 'Letter' in input_label and (not v.isalpha() or len(v) > 1):
                raise Exception('A letter field for the Excel config is not a single letter A-Z.')
        for k in self.csv_config:
            v = self.csv_config[k]
            if k not in csv_fields:
                raise Exception('Config key name mismatch for CSV section of config.')
            if v is None or v == '':
                raise Exception('One or more settings is blank for CSV config.')

    def validate(self):
        if not self.excel_config or not self.csv_config:
            raise Exception('Excel or CSV sections in config are missing or empty.')
        self.check_required_keys_exist()
        self.check_keys_match()
    
    def get_excel_config(self, key: ExcelFieldKey):
        if key not in excel_fields:
            raise Exception(f"Error: {key} not found in Excel section of config.")
        result = self.excel_config.get(key, None)
        if result is None:
            raise Exception(f"Error: Invalid field - {key}")
        return result 
    
    def get_csv_config(self, key: CsvFieldKey):
        if key not in csv_fields:
            raise Exception(f"Error: {key} not found in CSV section of config.")
        result = self.csv_config.get(key, None)
        if result is None:
            raise Exception(f"Error: Invalid field - {key}")
        return result 

def update_config(section: str, form: ImmutableMultiDict, out_filename: str):
    config = configparser.ConfigParser()
    config.read(out_filename)
    config[section] = form
    with open(out_filename, 'w') as configfile:
        config.write(configfile)