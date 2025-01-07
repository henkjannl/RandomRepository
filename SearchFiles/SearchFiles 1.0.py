import os
from datetime import datetime
from PIL import Image
# import exifread

# STEP 1: Where to start the search
#START_LOCATION = os.getcwd()
#START_LOCATION = "C:\\Users\\henkj\\OneDrive\\Documents\\eclipse"
# START_LOCATION = "C:\\Users\\henkj\\OneDrive\\Fotos"
START_LOCATION = "C:/Users/henkj/OneDrive/01 Gezamenlijk/03 Huis/03 Zonstraat - Hengelo/2024-12-02 Schutting bokwiel/Stuklijst/STEP en PDF"
#START_LOCATION = r"C:\Users\henkj\OneDrive\Fotos\2005"
#START_LOCATION = r"C:\Users\henkj\OneDrive\Fotos\Amber21"
#START_LOCATION = "C:\\"

# STEP 2: How deep do we want the search to be?
MAX_DEPTH = 0
#MAX_DEPTH = 1
# MAX_DEPTH = 999

# STEP 3: Filter the files that were found
def Requirement(direntry):
	return True
	#return direntry.name.lower().endswith('cr2')

	# return direntry.name.lower().endswith('cr2')  or \
    #        direntry.name.lower().endswith('arw')  or \
    #        direntry.name.lower().endswith('jpeg') or \
    #        direntry.name.lower().endswith('jpg')

	# return direntry.name.lower().endswith('cr2')  or \
    #        direntry.name.lower().endswith('jpeg') or \
    #        direntry.name.lower().endswith('jpg')  or \
    #        direntry.name.lower().endswith('tif')  or \
    #        direntry.name.lower().endswith('tiff') or \
    #        direntry.name.lower().endswith('png')
	#return direntry.is_dir()
	#return ('arduino' in direntry.name.lower()) and direntry.name.lower().endswith('svg')
	#return direntry.name.lower().endswith('h')
	#return ('linux' in direntry.name.lower()) and direntry.name.lower().endswith('fb.h')
	#return 'stratasys' in direntry.name.lower()


# STEP 4: Sort the files in a certain order
SORT = 'path'
#SORT = 'exifdate'
#SORT = 'date'
SORT_REVERSE = False
#SORT_REVERSE = True


# STEP 4: Define which COLUMNS in the report, in what order
COLUMNS = ['name']
#COLUMNS = ['name', 'path', 'size', 'created', 'modified', 'accessed']
#COLUMNS = ['name', 'path', 'created']
#COLUMNS = ['name', 'exifdate', 'width', 'height']
#COLUMNS = ['name', 'exifdate', 'size', 'width', 'height', 'path']

# Fields to choose from in the COLUMNS (can be extended)
TAGS = {
	'Name': 	lambda de: de.name,
	'Path': 	lambda de: de.path,
	'Size':		lambda de: '%d' % de.stat().st_size,
	'Created':	lambda de: datetime.fromtimestamp(de.stat().st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
	'Modified':	lambda de: datetime.fromtimestamp(de.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
	'Accessed':	lambda de: datetime.fromtimestamp(de.stat().st_atime).strftime('%Y-%m-%d %H:%M:%S'),
    'Exifdate':	lambda de: getExifDate(de.path),
    'Width':	lambda de: getImageWidth(de.path),
    'Height':	lambda de: getImageHeight(de.path)
}

# Get date of photo from Exif
def getExifDate(path):

    # Return nothing if the path is a directory
    if os.path.isdir(path):
        return ''

    try:
        # First try exifread, for raw files
        with open(path, 'rb') as image_file:
            tags = exifread.process_file(image_file)
        return tags['Image DateTime'].values
    except:
        try:
            # Then try PIL
            return Image.open(path)._getexif()[36867]
        except:
            # If that does not work, return an error
            return 'Error'

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
            return 'Error'



# Recursive function that scans the disk and add relevant files to the
# selection
def AddToSelection(baseDir, selection, requirement, level):

	try:
		files = os.scandir(baseDir)

		for de in files:

			if requirement(de):
				selection.append(de)

			if de.is_dir(follow_symlinks=False) and level<MAX_DEPTH:
				AddToSelection(de.path, selection, requirement, level+1)
	except:
		pass


# Retrieve files
selection = []
AddToSelection(START_LOCATION, selection, Requirement, level=MAX_DEPTH)

# for yr in range(2000, 2023):
#     loc = '{sl}\{yr}'.format(sl=START_LOCATION, yr=yr)
#     AddToSelection(loc, selection, Requirement, level=0)
#     print(loc, len(selection))

# This way of sorting allows to sort for fields that do not end up in search results
# However, it may require some fields to be retrieved twice, which is slower

# Sort the files
if SORT.lower()=='name':
	selection.sort( key= lambda de: de.name )
elif SORT.lower()=='path':
	selection.sort( key= lambda de: de.path )
elif SORT.lower()=='date':
	selection.sort( key= lambda de: de.stat().st_ctime )
elif SORT.lower()=='size':
	selection.sort( key= lambda de: de.stat().st_size )
elif SORT.lower()=='exifdate':
	selection.sort( key= lambda de: getExifDate(de.path) )
elif SORT.lower()=='width':
	selection.sort( key= lambda de: getImageWidth(de.path) )
elif SORT.lower()=='height':
	selection.sort( key= lambda de: getImageHeight(de.path) )

if SORT_REVERSE:
	selection = list(reversed(selection))


# Prepare the headers
results=''
for i,k in enumerate(COLUMNS):

	for l in TAGS.keys():

		if k.strip().upper() == l.strip().upper():
			results+=l+'\t'

			# Correct spelling mistakes in COLUMNS list
			COLUMNS[i]=l
results+='\n'

# Export file results
lines=[]
for i,de in enumerate(selection):

	lines.append('\t'.join( [TAGS[k](de) for k in COLUMNS]))

	# Only print the first 20 lines
	if i<20: print(de.name)


if len(selection)>=20:
	print(':')

results+='\n'.join(lines)

print('\nTotal of {cnt:d} files'.format(cnt=len(selection)))

import pyperclip
pyperclip.copy(results)
print('Results copied to clipboard')
