# MacroRemover
A python script which can be used during or after cyber incidents to prepare filesystems to be moved to the green/clean zone.

It is capable of parsing recursively at a defined starting path. It will check and process all files according to a ruleset which is defined in the config.ini. Processed files are copied into a processed folder which is mirroring the original directory structure.

# Functionalities
- The script is capable of removing Office macros using the Win32 API.
  To do this, the filetype (".doc", ".docm", ".xls", ".xslm") has to be defined under convert in the config.ini

- It is capable of removing defined files according to a fileending.
  To do this, the file ending (e.g. ".exe") has to be defined under remove in the config.ini

- It is capable of just copying filetypes into the processed folder
  To do this, the file ending (e.g. ".dwg") has to be defined under allowed in the config.ini

- To create a backup copy of the original files before they are deleted or the macros removed, it is possible to collect them and save them in a archive file with a password.
  To do this, the file endings (e.g. ".exe", ".doc") have to be defined under compressed in the config.ini. The password and the filename can be defined there as well.

# Config File
The configuration of the tool should be very much self explanatory.

### [config]
**Path where the files to be searched are located**

path = C:\Users\IEUser\Desktop\Handler\TEST\

**Path where matched files are placed. (Rule: "allowed")**

processed_files_path = C:\Users\IEUser\Desktop\Handler\processed\

**Path where converted files should be placed.**

_If the path is empty, "processed_file_path" is used_

converted_files_path = 

**Filename of the zip archive for files to be compressed**

zipfile_name = archive.zip

**Password of the zip archive**

zipfile_password = 123456

### [filetypes]

**File extensions defined here are allowed and will be copied to the "processed" folder.**

allowed = [".docx"]

**File extensions defined here will be converted to a non-macro format.**

_Currently supported formats: xls, xlsm, docm, doc_

convert = [".doc"]

**File extensions defined here will be deleted by the script.**

remove = []

**File extensions defined here will be placed in a zip archive in the "processed" folder.**

compress = [".txt",".doc"]

## Required Libraries

- pywin32
- pyminizip

