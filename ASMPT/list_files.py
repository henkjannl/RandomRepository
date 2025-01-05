import glob
import os
import pyperclip
DIR = "C:/Users/henkj/OneDrive/03 HenkJan/_ASM-PT/"
files =glob.glob(f'{DIR}/**/*.pdf', recursive=True)
filenames = [os.path.basename(file) for file in files]
pyperclip.copy('\n'.join(filenames))
print('\n'.join(filenames))