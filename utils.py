import pathlib
from openpyxl.utils import column_index_from_string
from datetime import datetime
import os
import sys
from flask import request
import glob
import shutil

def get_full_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return pathlib.Path(__file__).parent.resolve()

merge_name_map = {
    "Inv_Nbr": column_index_from_string("G") - 1,
    "Estate_No": column_index_from_string("I") - 1,
    "Bill_To": column_index_from_string("J") - 1,
    "Address_1": column_index_from_string("K") - 1,
    "Address_2": column_index_from_string("L") - 1,
    "Estate_of": column_index_from_string("M") - 1,
    "M_1st_Run": column_index_from_string("O") - 1,
    "M_2nd_Run": column_index_from_string("P") - 1,
    "M_3rd_Run": column_index_from_string("Q") - 1, 
    "price": column_index_from_string("R") - 1,
    "QR_Image": column_index_from_string("X") - 1,
    "QR_Link": column_index_from_string("Y") - 1
}

merge_id_key = "Inv_Nbr"

def get_formatted_value(key: str, value: str):
    if key == "price" and value is not None:
        return float(value)
    if "Run" in key:
        if value == "None":
            return ""
        dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%#m/%#d/%Y")
    return value

def get_credentials(is_prod) -> tuple[str, str]:
    proper_realm_id = os.getenv("PROD_REALM_ID" if is_prod else "DEV_REALM_ID")
    access_token = request.cookies.get('access_token')
    realm_id = request.cookies.get('realm_id')
    if access_token is None or realm_id is None or proper_realm_id != realm_id:
        raise Exception("Unauthorized. If you are the intended user for this app, please contact the site admin.")
    return (access_token, realm_id)

def get_filename_with_ext(dir: str, ext: str):
    if not os.path.exists(dir):
        return "None"
    for filename in glob.iglob(f"{dir}/*{ext}"):
        return filename
    return "None"

def get_all_dir_files(dir: str) -> list[str]:
    return [f.path for f in os.scandir(dir) if f.is_file()]

def zip_all_dir_files(dir: str, output_zip: str):
    shutil.make_archive(output_zip, 'zip', dir)

def remove_files_with_ext(dir: str, ext: str):
    if not os.path.exists(dir) or len(os.listdir(dir)) == 0:
        return
    for filename in glob.iglob(f"{dir}/*{ext}"):
        os.remove(filename)

