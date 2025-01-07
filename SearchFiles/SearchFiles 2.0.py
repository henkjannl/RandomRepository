# This script retrieves file information from disk
# and copies it to the clipboard in an Excel-friendly (tab delimeted) format
# Configure search job by defining some constants
# this is done by commenting out (ctrl-/) or adding lines

# Directory
# DIR = False # Means to search the directory in which the script resides
# DIR = 'C:/Users/hjvanderpol/OneDrive - ASMPT Limited/Papers/_Papers'
# DIR = 'C:/Users/hjvanderpol/OneDrive - ASMPT Limited'
DIR = 'C:/Users/hjvanderpol/Downloads'

# Recursive search (subdirectories) or not
# RECURSIVE = False
RECURSIVE = True 

# Filetypes
TYPE = '*.*'
# TYPE = '*.xlsx'
# TYPE = '*.png'
# TYPE = '*.jpg'

# Which file data to export. The top line has all available fields. Case insensitive.
# EXPORT = ['FULLPATH', 'PATH', 'FILE', 'EXT', 'MODIFIED', 'ACCESSED', 'CREATED', 'SIZE', 'PDF_DATE', 'EXIFDATE', 'WIDTH', 'HEIGHT']
# EXPORT = ['PATH', 'FILE', 'SIZE']
EXPORT = ['PATH', 'FILE', 'SIZE', 'MODIFIED', 'ACCESSED', 'CREATED']
# EXPORT = ['PATH', 'FILE', 'SIZE', 'EXIFDATE', 'WIDTH', 'HEIGHT']

# Which format to be used for date and time
# DATE_FMT = "%Y-%m-%d %H:%M:%S"
DATE_FMT = "%Y-%m-%d"

# How to sort the results
SORT = False
# SORT = 'FULLPATH'
# SORT = 'FILE'
# SORT = 'CREATED'
# SORT = 'PDF_DATE'
# SORT = 'EXIFDATE'
# SORT = 'WIDTH'
# SORT = 'HEIGHT'

SORT_REVERSE = False
#SORT_REVERSE = True

# ========================================================================
# Version history
# V1.0: First working version
# V1.1: Image information (exif, width, height) added
# V1.2: PDF date added
# V2.0: All constants to the top, all code below
#
# To do:
#
# ========================================================================

# ========================================================================
# Code below does not need to be changed 
# From here, data is retrieved from disk and copied to the clipboard
# ========================================================================
def has_field(field):
    return field.upper() in [elem.upper() for elem in EXPORT]

import os
import glob
# pip install pyperclip
import pyperclip
from datetime import datetime, timedelta, timezone

# Optional libraries
try:
    # pip install pypdf2
    from PyPDF2 import PdfReader 
except:
    pass

try:
    # pip install pillow
    from PIL import Image 
except:
    pass

try:
    # pip install exifread
    import exifread   
except:
    pass

# Convert datetime record to string
format_date = lambda date: date.strftime(DATE_FMT)

# Various functions to produce output on per-file basis
# most exception handling is done when filling in the fields
# so only do what is needed to optimize output


def parse_pdf_date(pdf_date):
    date_str = pdf_date[2:]    # Remove the "D:" prefix
    dt_part = date_str[:14]    # Extract the main date and time part
    dt = datetime.strptime(dt_part, "%Y%m%d%H%M%S")  # Parse the main part into a datetime object
    
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


# Get date of photo from Exif
def getExifDate(file):

    # Return nothing if the path is a directory
    if os.path.isdir(file):
        return ''

    try:
        # First try exifread, for raw files
        with open(file, 'rb') as image_file:
            tags = exifread.process_file(image_file)
            result = datetime.strptime(tags['Image DateTime'].values, "%Y:%m:%d %H:%M:%S")
        return format_date(result)
    except:
        try:
            # Then try PIL
            result = datetime.strptime(Image.open(file)._getexif()[36867], "%Y:%m:%d %H:%M:%S")
            return format_date(result)
        except:
            # If that does not work, return an error
            return 'N/A'


# Get width of photo from Exif
def getImageWidth(path):

    # Return nothing if the path is a directory
    if os.path.isdir(path):
        return ''

    try:
        # First try exifread, for raw files
        with open(path, 'rb') as image_file:
            tags = exifread.process_file(image_file)
        return tags['Image ImageWidth'].printable
    except:
        try:
            # Then try PIL
            return str(Image.open(path).size[0])
        except:
            # If that does not work, return an error
            return 'Error'


# Get height of photo from Exif
def getImageHeight(path):

    # Return nothing if the path is a directory
    if os.path.isdir(path):
        return ''

    try:
        # First try PIL
        return str(Image.open(path).size[1])
    except:
        try:
            # If that does not work, try exifread (for raw files)
            # Beware: 'Image ImageLength' might return the original image size;
            # the size may have been changed after creation
            with open(path, 'rb') as image_file:
                tags = exifread.process_file(image_file)
            return tags['Image ImageLength'].printable
        except:
            # If that does not work, return an error
            return 'N/A'

FIELDS = {
    'FULLPATH': lambda file: file,
    'PATH':     lambda file: os.path.dirname(file),
    'FILE':     lambda file: os.path.basename(file),
    'EXT':      lambda file: os.path.splitext(file)[1],
	'SIZE':		lambda file: f'{os.path.getsize(file)}',
    'MODIFIED': lambda file: format_date(datetime.fromtimestamp(os.path.getmtime(file))),
    'ACCESSED': lambda file: format_date(datetime.fromtimestamp(os.path.getatime(file))),
    'CREATED':  lambda file: format_date(datetime.fromtimestamp(os.path.getctime(file))),
    'PDF_DATE': lambda file: format_date(get_pdf_creation_date(file)),
    'EXIFDATE': lambda file: getExifDate(file),   # Date taken of image
    'WIDTH':    lambda file: getImageWidth(file), # Width of image
    'HEIGHT':   lambda file: getImageHeight(file) # Height of image
}

# Actual code starts here

# If DIR is false, replace it by the directory in which this script resides
if not DIR:
    DIR = os.path.dirname(os.path.abspath(__file__)) 

# Read files from disk
files = glob.glob(os.path.join(DIR, '**', TYPE), recursive=RECURSIVE)
print(f'{len(files)} files found.', end=' ')

# Straighten forward slashes and backslashes
files = [os.path.normpath(file) for file in files]

fields_upper = [field.upper() for field in EXPORT]

# Gather requested fields for each file
export = []
for file in files:
    rec = {}
    for field in fields_upper:
        try:
            rec[field] = FIELDS[field](file)
        except:
            rec[field] = 'N/A'

    # Also include the sort field so we can sort the files    
    if SORT:
        SORT = SORT.upper()
        if not has_field(SORT):
            try:
                rec[SORT] = FIELDS[SORT](file)
            except:
                rec[SORT] = 'N/A'

    export.append(rec)

# Sort the files
if SORT:
    export = sorted(export, key=lambda field: field[SORT], reverse=SORT_REVERSE)

# Convert the dictionaries in tab separated lines
export = ['\t'.join([rec[field] for field in fields_upper]) for rec in export]

# Add a header to the beginner. Use the non-uppercase version here
export.insert(0, '\t'.join(EXPORT))

pyperclip.copy('\n'.join(export))
print(f'Data copied to the clipboard')
