from QuickbooksInvoiceHandler import QuickbooksInvoiceHandler
from QRCodeHandler import QRCodeHandler
from ExcelHandler import ExcelHandler
from ExcelHandler import CorrespondingData
from MailMergeHandler import MailMergeHandler
from AuthHandler import AuthHandler
from flask import Flask, request, render_template, redirect, make_response, session, send_file
from werkzeug.utils import secure_filename
import os
from utils import get_credentials, get_filename_with_ext, zip_all_dir_files, remove_files_with_ext, get_abs_path
from dotenv import load_dotenv

load_dotenv()
UPLOAD_FOLDER = 'uploads'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = get_abs_path(UPLOAD_FOLDER)
app.secret_key = os.getenv("APP_SECRET_KEY")
is_prod = True

def get_error_html(e: Exception) -> str:
    return f'''
        <p>{e.__str__()}</p>
        <a href="/">Return to Home</a>
    '''

def process_data(access_token: str, realm_id: str):
    try:
        get_credentials(is_prod=is_prod)
        excel_filename = get_filename_with_ext(dir=app.config["UPLOAD_FOLDER"], ext=".xlsx")
        docx_filename = get_filename_with_ext(dir=app.config["UPLOAD_FOLDER"], ext=".docx")
        csv_filename = get_filename_with_ext(dir=app.config["UPLOAD_FOLDER"], ext=".csv")
        if excel_filename == "None" or docx_filename == "None" or csv_filename == "None":
            raise Exception("No files to uploaded to process.")
        abs_qr_path = get_abs_path("qr_codes")
        abs_invoice_path = get_abs_path("invoice_mail")
        abs_zip_path = get_abs_path("invoice_mail.zip")
        remove_files_with_ext(dir=abs_qr_path, ext=".png")
        remove_files_with_ext(dir=abs_invoice_path, ext=".docx")
        if os.path.exists(abs_zip_path):
            os.remove(abs_zip_path)
        print("\n")
        qh = QuickbooksInvoiceHandler(realm_id=realm_id, access_token=access_token, is_prod=is_prod)
        qr = QRCodeHandler(is_prod=is_prod, out_dir=abs_qr_path)
        excel = ExcelHandler(
            filename=excel_filename,
            worksheet_name="Bill&Cert"
        )
        print("\n")
        mm = MailMergeHandler(template_filename=docx_filename, out_dir=abs_invoice_path)
        print("\n")
        qh.import_csv(filename=csv_filename)
        print(f"\nInvoice IDs: {qh.invoice_ids}")
        print(f"\nInvoice Numbers: {qh.invoice_numbers}")
        print("\n")
        qr.generate_qr_codes(ids=qh.invoice_ids, prod_link_function=qh.generate_invoice_link)
        qr_path_data_to_add = CorrespondingData(col_name="X", data=qr.code_paths)
        qr_link_data_to_add = CorrespondingData(col_name="Y", data=qr.code_links)
        merge_data_list = qr.add_qrs_excel(
            excel=excel,
            invoice_num_col_name="G",
            invoice_nums=qh.invoice_numbers,
            data_lists=[qr_path_data_to_add, qr_link_data_to_add]
        )
        mm.merge_multiple(merge_data_list=merge_data_list)
        print("\n")
        mm.close()
        print("\n")
        zip_all_dir_files("invoice_mail", "invoice_mail")
        session["process_message"] = "Data Successfully Processed."
    except Exception as e:
        session["process_message"] = str(e)

@app.errorhandler(404)
def page_not_found(_):
    return render_template("404.html"), 404

@app.errorhandler(500)
def internal_server_error(_):
    return render_template("500.html"), 500

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        auth_handler = AuthHandler(is_prod=is_prod)
        auth_url = auth_handler.get_auth_url()
        return redirect(auth_url, code=302)
    return render_template('index.html')

@app.route('/authorize', methods=['GET'])
def authorize():
    try:
        auth_handler = AuthHandler(is_prod=is_prod)
        code_param = request.args.get('code')
        realm_id_param = request.args.get('realmId')
        (access_token, realm_id) = auth_handler.get_auth_tokens_from_code(code=code_param, realm_id=realm_id_param)
        resp = make_response(redirect("/process", code=302))
        resp.set_cookie(key="access_token", value=access_token, secure=is_prod, httponly=True, samesite='Lax')
        resp.set_cookie(key="realm_id", value=realm_id, secure=is_prod, httponly=True, samesite='Lax')
        return resp
    except Exception as e:
        print(f"Error in /authorize: {e}")
        return get_error_html(e)

@app.route('/process', methods=['GET', 'POST'])
def process():
    try:
        (access_token, realm_id) = get_credentials(is_prod=is_prod)
        if request.method == 'POST':
            process_data(access_token=access_token, realm_id=realm_id)
            return redirect("/process", code=302)
        excel_filename = get_filename_with_ext(dir=app.config["UPLOAD_FOLDER"], ext=".xlsx", full=False)
        docx_filename = get_filename_with_ext(dir=app.config["UPLOAD_FOLDER"], ext=".docx", full=False)
        csv_filename = get_filename_with_ext(dir=app.config["UPLOAD_FOLDER"], ext=".csv", full=False)
        abs_zip_path = get_abs_path("invoice_mail.zip")
        invoice_mail_exists = os.path.exists(abs_zip_path)
        upload_message = session.pop("upload_message", None)
        process_message = session.pop("process_message", None)
        return render_template(
            'process.html', 
            excel_filename=excel_filename, 
            docx_filename=docx_filename, 
            csv_filename=csv_filename,
            upload_message=upload_message,
            process_message=process_message,
            mail_to_download=invoice_mail_exists
        )
    except Exception as e:
        print(f"Error in /process: {e}")
        return get_error_html(e)

@app.route('/upload_files', methods=['POST'])
def upload_files():
    try:
        get_credentials(is_prod=is_prod)
        error = ""
        empty_file_count = 0
        keys = request.files.keys()
        values = request.files.values()
        files = list(zip(keys, values))
        if len(files) == 0:
            session["upload_message"] = "No files selected."
            return redirect("/process", code=302)
        if not os.path.exists(app.config['UPLOAD_FOLDER']):
            os.makedirs(app.config['UPLOAD_FOLDER'])
        for (extension, file) in files:
            if not file or not file.filename:
                empty_file_count += 1
                continue
            elif file.filename[-(len(extension)):] != extension:
                error += f"Wrong file extension for {extension} "
            else:
                filename = secure_filename(file.filename)
                remove_files_with_ext(dir=app.config['UPLOAD_FOLDER'], ext=extension)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        if empty_file_count >= 3:
            session["upload_message"] = "No files selected."
            return redirect("/process", code=302)
        if error:
            session["upload_message"] = error
            return redirect("/process", code=302)
        session["upload_message"] = "Files Successfully Uploaded."
        return redirect("/process", code=302)
    except:
        session["upload_message"] = "Error uploading files."
        return redirect("/process", code=302)

@app.route('/download_files', methods=['POST'])
def download_files():
    try:
        get_credentials(is_prod=is_prod)
        download_folder = get_abs_path("invoice_mail.zip")
        return send_file(path_or_file=download_folder, as_attachment=True)
    except Exception as e:
        session["download_message"] = str(e)
        return redirect("/process", code=302)

if __name__ == '__main__':
    app.run(debug=True)
