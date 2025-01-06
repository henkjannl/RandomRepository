import os

# Configure search job by defining some constants
# this is done by commenting out (ctrl-/) or adding lines

# Directory
# DIR = os.path.dirname(os.path.abspath(__file__)) # Directory in which script resides
DIR = 'C:/Users/hjvanderpol/OneDrive - ASMPT Limited/Papers/_Papers'
# DIR = 'C:/Users/hjvanderpol/OneDrive - ASMPT Limited'

# Filetypes
# TYPE = '*.xlsx'
# TYPE = '*.png'
TYPE = '*.pdf'

# Search subdirectories
# RECURSIVE = False
RECURSIVE = True 

# Which file data to export
EXPORT = ['PATH', 'FILE', 'DATE']
# EXPORT = ['PATH', 'FILE', 'DATE', 'PDF_DATE']
# EXPORT = ['FILE', 'PDF_DATE']

# Which format to be used for date and time
# DATE_FMT = "%Y-%m-%d %H:%M:%S"
DATE_FMT = "%Y-%m-%d"

# ========================================================================
# Code below does not need to be changed 
# From here, data is retrieved from disk and copied to the clipboard
# ========================================================================
import glob
import pyperclip
from datetime import datetime, timedelta, timezone
from PyPDF2 import PdfReader

# Convert datetime record to string
format_date = lambda date: date.strftime(DATE_FMT)

# Various functions to produce output on per-file basis
get_file_path = lambda file: os.path.dirname(file)
get_file_name = lambda file: os.path.basename(file)
get_file_last_modified_date = lambda file: datetime.fromtimestamp(os.path.getmtime(file))

def parse_pdf_date(pdf_date):
    # Remove the "D:" prefix
    date_str = pdf_date[2:]
    
    # Extract the main date and time part
    dt_part = date_str[:14]
    
    # Parse the main part into a datetime object
    dt = datetime.strptime(dt_part, "%Y%m%d%H%M%S")
    
    # Handle timezone offset
    if len(date_str) > 14:
        offset_sign = date_str[14]
        offset_hours = int(date_str[15:17])
        offset_minutes = int(date_str[18:20])
        offset = timedelta(hours=offset_hours, minutes=offset_minutes)
        
        # Apply the offset
        if offset_sign == '-':
            dt = dt.replace(tzinfo=timezone(-offset))
        elif offset_sign == '+':
            dt = dt.replace(tzinfo=timezone(offset))
    
    return dt

def get_pdf_creation_date(file_path):
    reader = PdfReader(file_path)
    metadata = reader.metadata
    creation_date = parse_pdf_date(metadata.get('/CreationDate', 'Unknown'))
    return creation_date

# Read files from disk
files = glob.glob(os.path.join(DIR, '**', TYPE), recursive=RECURSIVE)

# Format data for output
export = ['\t'.join(EXPORT)]
for file in files:
    expt = []
    for elmnt in EXPORT:
        if elmnt=='PATH': 
            expt.append(get_file_path(file))
        elif elmnt=='FILE': 
            expt.append(get_file_name(file))
        elif elmnt=='DATE': 
            expt.append(format_date(get_file_last_modified_date(file)))
        elif elmnt=='PDF_DATE': 
            expt.append(format_date(get_pdf_creation_date(file)))
    export.append('\t'.join(expt))

pyperclip.copy('\n'.join(export))
print(f'Data from {len(export)} files copied to the clipboard')