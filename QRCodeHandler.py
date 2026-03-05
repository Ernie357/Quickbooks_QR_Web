import qrcode
import os
from typing import List, Callable
from PIL.Image import Image
from ExcelHandler import ExcelHandler
from ExcelHandler import CorrespondingData
from utils import get_full_script_dir

''' 
    Handles QRCode generation and holds program data about them and their links
'''
class QRCodeHandler:
    def __init__(self, is_prod: bool):
        self.code_paths: List[str] = []
        self.code_links: List[str] = []
        self.is_prod = is_prod

    ''' 
        Takes link to generate from and a filename to name the output,
        returns a tuple of Image file and non-full path filename
    '''
    def make_image_from_link(self, link: str, filename: str) -> tuple[Image, str]:
        print("\nCreating QR Code image", filename)
        img = qrcode.make(link)
        file = img.get_image()
        return (file, filename)
    
    ''' 
        Takes tuple of Image file and filename, 
        saves image to filename and loads filename into object's code_paths

        filename in details MUST be a full, absolute path
    '''
    def save_img(self, details: tuple[Image, str]):       
        (img, filename) = details
        print("\nSaving image", filename)
        img.save(filename)
        self.code_paths.append(filename)

    ''' 
        Returns list of QR codes going into directory target_dir, 
        using ids to generate file names.

        target_dir must just be a name with no slashes or dots, like "qr_codes"
    '''
    def generate_qr_codes(
            self, 
            target_dir: str, 
            ids: List[int], 
            prod_link_function: Callable[[int], str]
        ):
        print("Generating QR Codes...")
        if len(ids) <= 0:
            print("\nNo data found. No QR Codes to generate.")
        script_path = get_full_script_dir()
        full_target_dir = f"{script_path}\\{target_dir}"
        if not os.path.exists(full_target_dir):
            os.makedirs(full_target_dir)
        for id in ids:
            dev_link = f"https://app.qbo.intuit.com/app/invoice?txnId={id}"
            link = prod_link_function(id) if self.is_prod else dev_link
            print("QR Link:", link)
            filename = f"invoice_link_{id}.png"
            (img, _) = self.make_image_from_link(link, filename)
            save_path = f"{full_target_dir}\\{filename}"
            self.save_img((img, save_path))
            self.code_links.append(link)

    ''' 
        Takes ExcelHandler object to work with, the letter of the column that contains
        the invoice numbers, the relevant invoice numbers from QuickbooksInvoiceHandler,
        and lists of tuples that contain a column number and a list of data to add to
        that column number corresponding with the invoice numbers

        Returns the a list of merge info for all relavant rows hit by QR adding
    '''
    def add_qrs_excel(
        self, 
        excel: ExcelHandler, 
        invoice_num_col_name: str, 
        invoice_nums: List[str], 
        data_lists: List[CorrespondingData]
    ) -> List[dict[str, str | float | dict ]]:
        print("\nAdding QR Code paths to excel Workbook", excel.filename)
        if len(self.code_paths) <= 0:
            print("\nNo QR Codes found.")
        merge_data_list = []
        def bind(row, id, idx, row_num):
            excel.add_corresponding_data(row_num=row_num, idx=idx, data_lists=data_lists)
            merge_data = excel.get_merge_data_from_row(row_num=row_num)
            merge_data_list.append(merge_data)
        excel.iterate_rows_by_ids_bind(callback=bind, id_col_letter=invoice_num_col_name, ids=invoice_nums)
        excel.save_file_changes()
        return merge_data_list