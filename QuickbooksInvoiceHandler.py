import csv
import requests
import datetime
from typing import List

''' 
    Takes the realm ID and relevant access token for an intuit developer account,
    gathers data from and interacts with the Quickbooks API
'''
class QuickbooksInvoiceHandler():
    def __init__(self, realm_id: str, access_token: str, is_prod: bool):
        self.realm_id = realm_id
        self.access_token = access_token
        self.url_base = "https://quickbooks.api.intuit.com/v3/company/" if is_prod else "https://sandbox-quickbooks.api.intuit.com/v3/company/"
        self.invoice_ids: List[int] = []
        self.invoice_urls: List[str] = []
        self.qr_code_paths: List[str] = []
        self.invoice_numbers: List[str] = []

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
        if response.status_code in (200, 201):
            data = response.json()
            customer_id = data["Customer"]["Id"]
            print("Customer successfully added. ID =", customer_id)
            return int(customer_id)
        else:
            print("Error adding customer: ", response.status_code, response.text)
            return -1
        
    ''' Takes invoice data and corresponding customer ID, returns uploaded invoice ID '''
    def upload_invoice(self, inv, customer_id: int) -> int:
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
                    "Amount": float(inv["*ItemAmount"]),
                    "Description": inv["ItemDescription"],
                    "SalesItemLineDetail": {
                        "ItemRef": {
                            "name": "Services", 
                            "value": "1"
                        }
                    }
                }
            ],
            "TxnDate": datetime.datetime.strptime(inv["*InvoiceDate"], "%m/%d/%y").strftime("%Y-%m-%d"),
            "CustomerRef": {"name": inv["*Customer"], "value": customer_id},
            "DueDate": datetime.datetime.strptime(inv["*DueDate"], "%m/%d/%y").strftime("%Y-%m-%d"),
            "PrivateNote": inv["Memo"]
        }
        response = requests.post(url=url, headers=headers, json=payload)
        if response.status_code == 200:
            data = response.json()
            invoice_id = data["Invoice"]["Id"]
            print("Invoice successfully added. ID =", invoice_id)
            return int(invoice_id)
        else:
            print(response.status_code)
            print("Error adding invoice: ", response.status_code, response.text)
            return -1

    ''' Loads invoice IDs and numbers into this object from {filename} CSV '''
    def import_csv(self, filename: str):
        print("Importing CSV data to QuickBooks...\n")
        with open(filename) as file:
            if not file:
                raise Exception("CSV file", filename, "not found.")
            reader = csv.DictReader(file)
            invoices = [row for row in reader]
            for inv in invoices:
                customer_id = self.upload_customer(inv["*Customer"])
                if customer_id <= 0:
                    continue
                invoice_id = self.upload_invoice(inv, customer_id)
                if invoice_id <= 0:
                    continue
                self.invoice_ids.append(invoice_id)
                self.invoice_numbers.append(inv["*InvoiceNo"])

    ''' Takes invoice ID, returns that invoices payment link from API '''
    def generate_invoice_link(self, id: int) -> str:
        print("\nGenerating invoice link for ID", id)
        url = f"{self.url_base}{self.realm_id}/invoice/{id}?include=invoiceLink&minorversion=65"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json"
        }
        response = requests.get(url=url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            link = data.get("Invoice", {}).get("InvoiceLink", None)
            if link is None:
                print("Could not find invoice link for ID", id)
                return ""
            print("Invoice Link:", link)
            return link
        else:
            print("Error generating invoice link for ID", id)
            return ""