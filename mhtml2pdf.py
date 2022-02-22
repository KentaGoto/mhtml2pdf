# coding: utf-8
import os
import shutil
import win32com
from win32com.client import *


def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)


def mhtml2pdf(mhtml_fullpath, word):
    mhtml_fullpath = mhtml_fullpath.replace("/", "\\")
    print(mhtml_fullpath)

    dirname = os.path.dirname(mhtml_fullpath)
    current_file = os.path.basename(mhtml_fullpath)
    fname, ext = os.path.splitext(current_file)
    doc = word.Documents.Open(mhtml_fullpath)  # Open the mhtml (or mht) in Word
    # Save as PDF file
    doc.SaveAs(dirname + '/' + fname + '.pdf', FileFormat=17)
    doc.Close()


if __name__ == '__main__':
    s = input("Dir: ")
    root_dir = s.strip('\"')
    root_dir_copy = root_dir + '__copy'
    shutil.copytree(root_dir, root_dir_copy)

    # Com object
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    print('Processing...')

    for i in all_files(root_dir_copy):
        dirname = os.path.dirname(i)
        current_file = os.path.basename(i)
        fname, ext = os.path.splitext(current_file)
        
        if ext == '.mht' or ext == '.mhtml':
            try:
                # Convert mhtml to pdf
                mhtml2pdf(dirname + '/' + current_file, word)
            except:
                print('Error: ' + i)

    word.Quit()

    print('')
    print('Done!')
    print('Enter to exit.')
    os.system("pause > nul")
