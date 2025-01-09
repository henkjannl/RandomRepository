# This script retrieves file information from disk and
# copies it to the clipboard in an Excel-friendly (tab delimeted) format
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

# Include hidden files
# INCLUDE_HIDDEN = False
INCLUDE_HIDDEN = True

# Filetypes
# TYPES = ['*.*']
# TYPES = ['*.xlsx']
# TYPES = ['*.pdf']
TYPES = ['*.png', '*.tiff', '*.tif', '*.jpg', '*.jpeg', '*.cr2', '*.arw']
# TYPES = ['*.jpg']

# Which file data to export. The top line has all available fields. Case insensitive.
# EXPORT = ['FULLPATH', 'PATH', 'FILE', 'EXT', 'MODIFIED', 'ACCESSED', 'CREATED', 'SIZE', 'PDF_DATE', 'EXIFDATE', 'WIDTH', 'HEIGHT']
# EXPORT = ['PATH', 'FILE', 'SIZE']
# EXPORT = ['PATH', 'FILE', 'SIZE', 'MODIFIED', 'ACCESSED', 'CREATED']
EXPORT = ['Path', 'File', 'Size', 'Created', 'ExifDate', 'Width', 'Height']

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
# Code below typically does not need to be changed 
# From here, data is retrieved from disk and copied to the clipboard
#
# Version history
# V1.0: First working version
# V1.1: Image information (exif, width, height) added
# V1.2: PDF date added
# V2.0: All constants to the top, all code below
# V2.1: Support for hidden files added
#       Support for multiple file types added
#
# To do:
#    perhaps write as single function call to be used by overarching scripts
# ========================================================================
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

# ========================================================================
# Various functions to produce output on per-file basis
# most exception handling is done when filling in the fields
# so only do what is needed to optimize output
# ========================================================================
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
            return 'N/A'

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

def get_field(file, field):
    try:
        return FIELDS[field.upper()](file)
    except:
        return "N/A"

# ========================================================================
# Main starts here
# ========================================================================

# If DIR is false, replace it by the directory in which this script resides
if not DIR:
    DIR = os.path.dirname(os.path.abspath(__file__)) 

# Read files from disk
files = []
for type in TYPES:
    files.extend( glob.glob(os.path.join(DIR, '**', type), recursive=RECURSIVE, include_hidden=INCLUDE_HIDDEN) )
print(f'{len(files)} files found.', end=' ')

# Straighten forward slashes and backslashes
files = [os.path.normpath(file) for file in files]

# Create list of uppercase fields without modifying EXPORT
fields_upper = [field.upper() for field in EXPORT]

# Add sort field to the list if needed
if SORT:
    SORT = SORT.upper()
    if SORT not in fields_upper:
        fields_upper.append(SORT)

# Gather a dictionary of requested fields for each file
export_files = [{field: get_field(file, field) for field in fields_upper} for file in files]

# Sort the files
if SORT:
    export_files = sorted(export_files, key=lambda field: field[SORT], reverse=SORT_REVERSE)

# Convert the dictionaries in tab separated lines
export_files = ['\t'.join([file[field.upper()] for field in EXPORT]) for file in export_files]

# Insert header at the beginning
export_files.insert(0, '\t'.join(EXPORT))

pyperclip.copy('\n'.join(export_files))
print(f'Data copied to the clipboard')
