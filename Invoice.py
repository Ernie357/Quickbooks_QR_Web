from Config import Config, CsvFieldKey
from typing import Any

''' Represents a single invoice, verifying its structure and data integrity against the CSV file '''
class Invoice:
    def __init__(self, inv: dict[str | Any, str | Any], config: Config):
        self.inv = inv
        self.config = config
        self.invoice_number = self.get_inv_field('invoice_number_name')
        self.invoice_date = self.get_inv_field('invoice_date_name')
        self.customer_name = self.get_inv_field('customer_name')
        self.item_amount = float(self.get_inv_field('item_amount_name'))
        self.memo = self.get_inv_field('memo_name')
        self.due_date = self.get_inv_field('due_date_name')
        self.item_description = self.get_inv_field('item_description_name')
    
    def get_inv_field(self, key: CsvFieldKey):
        csv_column_name = self.config.get_csv_config(key)
        result = self.inv.get(csv_column_name, None)
        if result is None:
            raise Exception(f"CSV field {key} does not exist in uploaded file.")
        return result
    
class FullInvoice:
    def __init__(self, invoice: Invoice, invoice_id: int, customer_id: int):
        self.__dict__.update(invoice.__dict__)
        self.invoice_id = invoice_id
        self.customer_id = customer_id