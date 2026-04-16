import csv
import requests
import datetime
from Config import Config, CsvFieldKey
from Invoice import Invoice, FullInvoice

class UnauthorizedException(Exception):
    pass

''' 
    Takes the realm ID and relevant access token for an intuit developer account,
    gathers data from and interacts with the Quickbooks API
'''
class QuickbooksInvoiceHandler():
    def __init__(self, realm_id: str, access_token: str, is_prod: bool, config: Config):
        self.realm_id = realm_id
        self.access_token = access_token
        self.is_prod= is_prod
        self.config = config
        self.url_base = "https://quickbooks.api.intuit.com/v3/company/" if self.is_prod else "https://sandbox-quickbooks.api.intuit.com/v3/company/"
        self.invoices: list[FullInvoice] = []

    def send_invoice(self, invoice_id: str):
        print("Sending email for invoice ID", invoice_id)
        url = f"{self.url_base}{self.realm_id}/invoice/{invoice_id}/send?sendTo=johnmasreglia32@gmail.com&minorversion=65"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "Content-Type": "application/octet-stream"
        }
        response = requests.post(url=url, headers=headers)
        if response.status_code == 401:
            print("Refreshing Access Token...")
            raise UnauthorizedException("Missing Access Token")
        if response.status_code != 200:
            err = f"Error sending email for invoice ID {invoice_id}"
            print(err)
            raise Exception(err)

    ''' Takes customer DisplayName to check, returns their ID or -1 if not found '''
    def customer_exists(self, name: str) -> int:
        print("\nChecking to see if customer", name, "already exists...")
        url = f"{self.url_base}{self.realm_id}/query?minorversion=65"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "Content-Type": "application/text"
        }
        query = f"select * from Customer where DisplayName = '{name}'"
        response = requests.post(url, headers=headers, data=query)
        if response.status_code == 401:
            print("Refreshing Access Token...")
            raise UnauthorizedException("Missing Access Token")
        if(response.status_code == 200):
            data = response.json()
            customer = data.get("QueryResponse", {}).get("Customer", [])
            if len(customer) <= 0 or not customer[0]["Id"]:
                print("Customer", name, "does not exist.")
                return -1
            print("Customer", name, "already exists.")
            return int(customer[0]["Id"])
        print("Customer", name, "does not exist.")
        return -1

    ''' Uploads customer by DisplayName and returns the uploaded ID '''
    def upload_customer(self, name: str) -> int:
        print("\nUploading Customer: ", name)
        existing_customer_id = self.customer_exists(name)
        if existing_customer_id != -1:
            return existing_customer_id
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
        payload = {
            "DisplayName": name
        }
        url = f"{self.url_base}{self.realm_id}/customer?minorversion=65"
        response = requests.post(url=url, headers=headers, json=payload)
        if response.status_code == 401:
            print("Refreshing Access Token...")
            raise UnauthorizedException("Missing Access Token")
        if response.status_code in (200, 201):
            data = response.json()
            customer_id = data["Customer"]["Id"]
            print("Customer successfully added. ID =", customer_id)
            return int(customer_id)
        else:
            raise Exception("Error adding customer: ", response.status_code, response.text)
        
    ''' Takes invoice data and corresponding customer ID, returns uploaded invoice ID '''
    def upload_invoice(self, inv: Invoice, customer_id: int) -> int:
        print("\nUploading invoice with customer ID", customer_id)
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
        url = f"{self.url_base}{self.realm_id}/invoice?minorversion=65"
        payload = {
            "Line": [
                {
                    "DetailType": "SalesItemLineDetail", 
                    "Amount": inv.data['item_amount_name'],
                    "Description": inv.data['item_description_name'],
                    "SalesItemLineDetail": {
                        "ItemRef": {
                            "name": "Services", 
                            "value": "189" if self.is_prod else "1"
                        }
                    }
                }
            ],
            "DocNumber": inv.data['invoice_number_name'],
            "TxnDate": inv.data['invoice_date_name'],
            "CustomerRef": { "name": inv.data['customer_name'], "value": customer_id },
            "DueDate": inv.data['due_date_name'],
            "PrivateNote": inv.data['memo_name']
        }
        response = requests.post(url=url, headers=headers, json=payload)
        if response.status_code == 401:
            print("Refreshing Access Token...")
            raise UnauthorizedException("Missing Access Token")
        if response.status_code == 200:
            data = response.json()
            invoice_id = data["Invoice"]["Id"]
            self.send_invoice(invoice_id=invoice_id)
            print("Invoice successfully added. ID =", invoice_id)
            return int(invoice_id)
        else:
            raise Exception("Error adding invoice: ", response.status_code, response.text)

    ''' Loads invoice IDs and numbers into this object from {filename} CSV, validating CSV values first '''
    def import_csv(self, filename: str):
        print("Importing CSV data to QuickBooks...\n")
        with open(filename) as file:    
            reader = self.verify_import_structure(filename, file)
            invoices = [row for row in reader]
            partial_invoices: list[Invoice] = []
            for inv in invoices:
                structured_invoice = Invoice(inv, self.config)
                partial_invoices.append(structured_invoice)
            for inv in partial_invoices:
                customer_id = self.upload_customer(self.config.get_csv_config('customer_name'))
                if customer_id <= 0:
                    continue
                invoice_id = self.upload_invoice(structured_invoice, customer_id)
                if invoice_id <= 0:
                    continue
                full_invoice = FullInvoice(structured_invoice, invoice_id, customer_id)
                self.invoices.append(full_invoice)
                
    ''' Takes invoice ID, returns that invoices payment link from API '''
    def generate_invoice_link(self, id: int) -> str:
        print("\nGenerating invoice link for ID", id)
        url = f"{self.url_base}{self.realm_id}/invoice/{id}?include=invoiceLink&minorversion=65"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json"
        }
        response = requests.get(url=url, headers=headers)
        if response.status_code == 401:
            print("Refreshing Access Token...")
            raise UnauthorizedException("Missing Access Token")
        if response.status_code == 200:
            data = response.json()
            link = data.get("Invoice", {}).get("InvoiceLink", None)
            if link is None:
                raise Exception("Could not find invoice link for ID", id)
            print("Invoice Link:", link)
            return link
        else:
            raise Exception("Error generating invoice link for ID", id)
        
    def verify_import_structure(self, filename: str, file):
        if not file:
            raise Exception(f"CSV file {filename} not found.")
        reader = csv.DictReader(file)
        if reader is None or reader.fieldnames is None:
            raise Exception(f"CSV file {filename} is empty.")
        for v in self.config.csv_config.values():
            if v not in reader.fieldnames:
                raise Exception(f"Column name {v} does not exist in the config.")
        return reader