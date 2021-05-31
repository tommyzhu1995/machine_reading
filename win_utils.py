import os
import shutil
from win32com import client
import pythoncom


def doc2docx(fpath, tpath):
    '''
    Args:
        fpath and tpath both need to be absolute paths, so they
        are preprocessed firstly to fullfill the requirements.
    '''
    fpath = os.path.abspath(fpath)
    tpath = os.path.abspath(tpath)
    try:
        pythoncom.CoInitialize()
        word = client.DispatchEx('word.Application') # Independent process
        word.Visible = 0 # Do not display the word GUI
        word.DisplayAlerts = 0 # Do not display warnings.
    except Exception as e:
        print(e)
        if 'word' in locals():
            shutil.copyfile(fpath, tpath)
            word.Quit()

        return

    try:
        doc = word.Documents.Open(fpath)
        doc.SaveAs(tpath, 12) # param 16 is for saving as doc, and 12 is for saving as docx
    except Exception as e:
        print(e)
        shutil.copyfile(fpath, tpath)
    finally:
        if 'doc' in locals():
            doc.Close()

        word.Quit()


if __name__ == "__main__":
    fpath = r'C:\Users\aaronnli\Desktop\cimc_test\DFLNGHT-0003（20210316终版）.doc'
    tpath = r'C:\Users\aaronnli\Desktop\cimc_test\DFLNGHT-0003（20210316终版）.docx'
    # fpath = r'C:\Users\aaronnli\Desktop\cimc_test\test.doc'
    # tpath = r'C:\Users\aaronnli\Desktop\cimc_test\test.docx'
    doc2docx(fpath, tpath)