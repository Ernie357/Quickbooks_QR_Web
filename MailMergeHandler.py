from mailmerge import MailMerge
from docx import Document
from utils import get_full_script_dir, merge_id_key
from typing import List
import os

'''
    Takes the filename of the docx template with MERGEFIELD names;
    Handles Mail Merge operations, docx outputs, and image replacements
'''
class MailMergeHandler:
    def __init__(self, template_filename: str):
        self.filename = template_filename
        print("Initializing Mail Merge Handler from file", self.filename)
        self.template_document = MailMerge(self.filename)
        self.no_doc_err = f"No template document found for {self.filename}"

    ''' Closes the template doc, DO THIS AFTER DONE WITH OBJECT '''
    def close(self):
        if self.template_document is None:
            raise Exception(self.no_doc_err)
        print("Closing Word Document", self.filename)
        self.template_document.close()
        self.template_document = None

    '''
        Helpful utility for viewing the MERGEFIELD names for this docx
    '''
    def print_merge_fields(self):
        if self.template_document is None:
            raise Exception(self.no_doc_err)
        print(self.template_document.get_merge_fields())

    '''  
        merge_data: Dictionary of merge data, where keys are mergefields from the word
            template docx and values are properly formatted; MUST match the field 
            name map in utils.py

        id: A string that is appended to each output file from the merge to 
            distinguish each result; MUST match one of the keys of the map
            from utils.py; Currently is Inv Nbr
    '''
    def initate_merge(self, merge_data, id: str):
        if self.template_document is None:
            raise Exception(self.no_doc_err)
        print("\nInitiating Merge for invoice #", id, "...")
        image = str(merge_data.pop("QR_Image"))
        self.template_document.merge(**merge_data)
        out_dir = "invoice_mail"
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)
        script_path = get_full_script_dir()
        full_target_dir = os.path.join(script_path, out_dir)
        out_file = os.path.join(full_target_dir, f"invoice_{id}.docx")
        self.write_document_out(out_file=out_file)
        image_pairs = [("media/image4", image)]
        self.replace_images(doc_filename=out_file, image_pairs=image_pairs)

    '''
        Takes a list of merge data dictionaries, all must be of the same structure,
        and performs a mail merge for each identified by an id (see function above)
    '''
    def merge_multiple(self, merge_data_list):
        if self.template_document is None:
            raise Exception(self.no_doc_err)
        for merge_data in merge_data_list:
            self.initate_merge(merge_data=merge_data, id=merge_data[merge_id_key])

    '''
        Writes the filled out template file results to out_file
    '''
    def write_document_out(self, out_file: str):
        print("Writing out to", out_file)
        if self.template_document is None:
            raise Exception(self.no_doc_err)
        self.template_document.write(out_file)

    '''
        Takes a list of tuples of format (original_image_name, new_image_name.png)
        Replaces the old images with the new and overwrites the document

        IMPORTANT NOTE: the original_image_name in each tuple MUST be of format
        "media/image{number}", currently the 1st dummy image (QR Code)
        starts at media/image4
    '''
    def replace_images(self, doc_filename: str, image_pairs: List[tuple[str, str]]):
        doc = Document(doc_filename)
        for pair in image_pairs:
            (dummy_image_basename, new_image_path) = pair
            print("Adding image", new_image_path, "to", doc_filename)
            for rel in doc.part.rels.values():
                if "image" in rel.reltype:
                    if dummy_image_basename in rel.target_ref:
                        img_part = rel.target_part
                        with open(new_image_path, "rb") as f:
                            img_part._blob = f.read()
                        break
        doc.save(doc_filename)