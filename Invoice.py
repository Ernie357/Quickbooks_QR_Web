from Config import Config, CsvFieldKey, csv_fields, CsvFullKey, CsvFieldConfig
from typing import Any, assert_never
import datetime

DataValue = str | float | None

''' Represents a single invoice, verifying its structure and data integrity against the CSV file '''
class Invoice:
    def __init__(self, inv: dict[str | Any, str | Any], config: Config):
        self.inv = inv
        self.config = config
        self.data: dict[CsvFieldKey, DataValue] = { k: self.get_inv_field(k) for k in csv_fields.keys() }
    
    def get_inv_field(self, key: CsvFieldKey):
        csv_column_name = self.config.get_csv_config(key)
        result = self.inv.get(csv_column_name, None)
        props = csv_fields[key]
        if result is None and not props['optional']:
            raise Exception(f"CSV field {key} does not exist in uploaded file or required value for {key} is missing.")
        if result is None:
            return None
        return self.validate_and_parse(key, result, props)
    
    def validate_and_parse(self, key: str, value: str, props: CsvFieldConfig):
        match props['type']:
            case 'number':
                try:
                    return float(value)
                except:
                    raise Exception(f"CSV value for {key} is not a number.")
            case 'date':
                try:
                    return datetime.datetime.strptime(value, "%m/%d/%y").strftime("%Y-%m-%d")
                except:
                    raise Exception(f"CSV value for {key} is not a proper date.")
            case 'string':
                return value
            case other:
                assert_never(other)
    
class FullInvoice:
    def __init__(self, invoice: Invoice, invoice_id: int, customer_id: int):
        self.data: dict[CsvFullKey, DataValue] = { 
            **invoice.data, 
            'invoice_id': invoice_id, 
            'customer_id': customer_id 
        }