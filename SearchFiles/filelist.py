# Configure search job by defining some constants
# this is done by commenting out (ctrl-/) or adding lines

# Directory
# DIR = False # Means to search the directory in which the script resides
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
# EXPORT = ['PATH', 'FILE', 'EXT', 'DATE', 'PDF_DATE', 'EXIFDATE', 'WIDTH', 'HEIGHT']
EXPORT = ['FullPath', 'Path', 'File', 'Ext', 'Date', 'Pdf_date']
# EXPORT = ['PATH', 'FILE', 'DATE', 'PDF_DATE']
# EXPORT = ['FILE', 'PDF_DATE']


# Which format to be used for date and time
# DATE_FMT = "%Y-%m-%d %H:%M:%S"
DATE_FMT = "%Y-%m-%d"

# How to sort the results
# SORT = 'PATH'
# SORT = 'FILE'
# SORT = 'DATE'
SORT = 'PDF_DATE'
# SORT = 'EXIFDATE'
# SORT = 'WIDTH'
# SORT = 'HEIGHT'

SORT_REVERSE = False
#SORT_REVERSE = True

# ========================================================================
# To do:
# include extension
# include full path (path+filename)
# ========================================================================

# ========================================================================
# Code below does not need to be changed 
# From here, data is retrieved from disk and copied to the clipboard
# ========================================================================
def has_field(field):
    return field.upper() in [elem.upper() for elem in EXPORT]

import os
import glob
import pyperclip
from datetime import datetime, timedelta, timezone

# Optional libraries
try:
    from PyPDF2 import PdfReader
except:
    pass

try:
    from PIL import Image
except:
    pass

try:
    import exifread
except:
    pass

# Convert datetime record to string
format_date = lambda date: date.strftime(DATE_FMT)

# Various functions to produce output on per-file basis
# most exception handling is done when filling in the fields
# so only do what is needed to optimize output
get_file_extension = lambda file: os.path.splitext(file)[1]
get_file_path = lambda file: os.path.dirname(file)
get_file_name = lambda file: os.path.basename(file)
get_file_last_modified_date = lambda file: datetime.fromtimestamp(os.path.getmtime(file))


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
        return tags['Image DateTime'].values
    except:
        try:
            # Then try PIL
            return Image.open(file)._getexif()[36867]
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
    'PATH':     lambda file: get_file_path(file),
    'FILE':     lambda file: get_file_name(file),
    'EXT':      lambda file: get_file_extension(file),
    'DATE':     lambda file: format_date(get_file_last_modified_date(file)),
    'PDF_DATE': lambda file: format_date(get_pdf_creation_date(file)),
    'EXIFDATE': lambda file: getExifDate(file),
    'WIDTH':    lambda file: getImageWidth(file),
    'HEIGHT':   lambda file: getImageHeight(file)
}

# Actual code starts here

# If DIR is false, replace it by the directory in which this script resides
if not DIR:
    DIR = os.path.dirname(os.path.abspath(__file__)) 

# Read files from disk
files = glob.glob(os.path.join(DIR, '**', TYPE), recursive=RECURSIVE)

# Straighten forward slashes and backslashes
files = [os.path.normpath(file) for file in files]

fields_upper = [field.upper() for field in EXPORT]
SORT = SORT.upper()

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
    if not has_field(SORT):
        try:
            rec[SORT] = FIELDS[SORT](file)
        except:
            rec[SORT] = 'N/A'

    export.append(rec)

# Sort the files
export = sorted(export, key=lambda x: x[SORT], reverse=SORT_REVERSE)

# Convert the dictionaries in tab separated lines
export = ['\t'.join([rec[field] for field in fields_upper]) for rec in export]

# Add a header to the beginner. Use the non-uppercase version here
export.insert(0, '\t'.join(EXPORT))

pyperclip.copy('\n'.join(export))
print(f'Data from {len(export)} files copied to the clipboard')
