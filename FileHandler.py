import os
import configparser
import json
import re
from pathlib import Path
from zipfile import ZipFile
import shutil
import win32com.client
from win32com.client import constants
import pyminizip
import warnings
import re

dirname = os.path.dirname(__file__)

# Ignore deprecation Warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# Configuration
configfile_name = "config.ini"
configParser = configparser.RawConfigParser()
configParser.read(configfile_name, encoding='utf-8-sig')

# Parse Configuration
# Paths
src_dir = configParser.get("config", "path")
processed_file_path = configParser.get("config", "processed_files_path")
converted_file_path = configParser.get("config", "converted_files_path")
# If converted file path is empty, use processed filepath instead.
if not converted_file_path:
    converted_file_path = processed_file_path

# Zipfile
zipfile_name = configParser.get("config", "zipfile_name")
zipfile_password = configParser.get("config", "zipfile_password")
zipfile_path = processed_file_path

def load_as_json_if_not_empty(config_type, config_name):
    if configParser.get(config_type,config_name):
        config_as_json = json.loads(configParser.get(config_type,config_name))
        return config_as_json

# Filetypes
allowed = load_as_json_if_not_empty("filetypes","allowed")
# allowed = json.loads(configParser.get("filetypes","allowed"))
convert = load_as_json_if_not_empty("filetypes", "convert")
remove = load_as_json_if_not_empty("filetypes", "remove")
compress = load_as_json_if_not_empty("filetypes", "compress")

# Init for some variables and arrays
compress_files = []
remove_files = []


def create_dir(path):
    try:
        os.makedirs(path)
        #print("Directory" , processed_file_path , "Created ")
    except FileExistsError:
        print("INFO: Directory:" , path , "already exists")

def get_full_subdir(root, dir):
    subdir = root.replace(src_dir, '')
    subdir = os.path.join(subdir, dir)
    return subdir

def process_files(src_dir):
# traverse root directory, and list directories as dirs and files as files
    for root, dirs, files in os.walk(src_dir):
        for dir in dirs:
            #Create Subdirectory for processed files
            if get_full_subdir(root, dir):
                create_dir(os.path.join(processed_file_path,get_full_subdir(root, dir)))
            if not converted_file_path == processed_file_path:
                create_dir(os.path.join(converted_file_path,get_full_subdir(root, dir)))

        path = root.split('\\')
        for file in files:
            dir_name = os.path.dirname(get_full_subdir(root, file))
            suffix = (Path(file).suffix)
            
            # print("Current file suffix" , suffix)
            
            if suffix in convert:
                #print("FILE:" , root, file , "needs to be converted")
                convert_file(root, file, suffix)

            if suffix in compress:
                compress_files.append((os.path.join(root, file),dir_name))

            if suffix in remove:    
                remove_files.append(os.path.join(root, file))                 
                             
            if suffix in allowed:
                #("INFO: Allowed File found: " , os.path.join(root, file), " Copying...")
                #print ("Suffix Allowed Filepath:", dir_name)
                #print("FILE: Moving file:" , os.path.join(root, file) , "to processed folder")
                shutil.copy(os.path.join(root, file), os.path.join(processed_file_path,dir_name,file))
           
    zip_file(compress_files)
    remove_file(remove_files)

def remove_file(remove_files):
        for file in remove_files: 
            os.remove(file)
    
      
def zip_file(compress_files):
    file_paths = [x[0] for x in compress_files]
    sub_dirs = ["\\" if x[1] == '' else x[1] for x in compress_files]

    if file_paths:
        pyminizip.compress_multiple(file_paths, sub_dirs, os.path.join(zipfile_path, zipfile_name), zipfile_password, 5)

def convert_file(file_path, file, suffix):
    if suffix == ".doc" or suffix == ".docm":
        convert_doc2docx(file_path, file)
        
    if suffix == ".xls" or suffix == ".xlsm":
        convert_xls2xlsx(file_path, file)

def convert_doc2docx(file_path, file):
    try:
        dir_name = os.path.dirname(get_full_subdir(file_path, ''))
        file_path = os.path.abspath(os.path.join(file_path, file))
        
        # Opening MS Word
        word = win32com.client.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(file_path)
        doc.Activate ()

        # Rename processed path with .docx       
        new_file_abs = os.path.abspath(os.path.join(converted_file_path,dir_name,file))
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)
        word.Application.Quit()
    except:
        print("ERROR: Unable to open or convert '" + file + "'. Skipping file.")

def convert_xls2xlsx(file_path, file):
    try:
        dir_name = os.path.dirname(get_full_subdir(file_path, ''))
        file_path = os.path.abspath(os.path.join(file_path, file))
        
        # Opening Excel
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file_path)

        # Rename processed path with .xlsx
        new_file_abs = os.path.abspath(os.path.join(converted_file_path,dir_name,file))         
        new_file_abs = re.sub(r'\.\w+$', '.xlsx', new_file_abs)

        # Save and Close
        wb.SaveAs(new_file_abs, FileFormat = 51) #FileFormat = 51 is for .xlsx extension
        wb.Close(False)                          #FileFormat = 56 is for .xls extension
        excel.Application.Quit()
    except:
        print("ERROR: Unable to open or convert '" + file_path + "'. Skipping file.")


def main():
    # Create Subdirectory for processed and converted files
    create_dir(processed_file_path)
    if not converted_file_path == processed_file_path:
        create_dir(converted_file_path)

    
    process_files(src_dir)
    print()
    print("Finished!")
    input("Press enter to exit")


if __name__ == "__main__":
    main()
