# coding: utf-8

import os
import shutil
import win32com
from win32com.client import *

def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)

def doc2docx(doc_fullpath):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc_fullpath = doc_fullpath.replace("/", "\\")
    print(doc_fullpath)
    dirname = os.path.dirname(doc_fullpath)
    current_file = os.path.basename(doc_fullpath)
    fname, ext = os.path.splitext(current_file)
    os.chdir(dirname)
    doc = word.Documents.Open(doc_fullpath)
    doc.SaveAs(dirname + '/' + fname + '.docx', FileFormat=16)
    doc.Close()
    return dirname + '/' + fname + '.docx'

def ppt2pptx(ppt_fullpath):
    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    powerpoint.DisplayAlerts = 0
    ppt_fullpath = ppt_fullpath.replace("/", "\\")
    print(ppt_fullpath)
    dirname = os.path.dirname(ppt_fullpath)
    current_file = os.path.basename(ppt_fullpath)
    fname, ext = os.path.splitext(current_file)
    ppt = powerpoint.Presentations.Open(ppt_fullpath, False, False, False)
    ppt.SaveAs(dirname + '/' + fname + '.pptx')
    ppt.Close()
    return dirname + '/' + fname + '.pptx'

def xls2xlsx(xls_fullpath):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = 0
    xls_fullpath = xls_fullpath.replace("/", "\\")
    print(xls_fullpath)
    dirname = os.path.dirname(xls_fullpath)
    current_file = os.path.basename(xls_fullpath)
    fname, ext = os.path.splitext(current_file)
    xls = excel.Workbooks.Open(xls_fullpath)
    xls.SaveAs(dirname + '/' + fname + '.xlsx', FileFormat=51)
    xls.Close()
    return dirname + '/' + fname + '.xlsx'

if __name__ == '__main__':
    s = input("Dir: ")
    root_dir = s.strip('\"')
    root_dir_copy = root_dir + '__copy'
    shutil.copytree(root_dir, root_dir_copy)

    # Convert doc, ppt, xls to docx, pptx, xlsx.
    print('Processing...')
    for i in all_files(root_dir_copy):
        dirname = os.path.dirname(i)
        current_file = os.path.basename(i)
        fname, ext = os.path.splitext(current_file)
        os.chdir(dirname)
        if ext == '.doc':
            docx = doc2docx(dirname + '/' + current_file)
            os.remove(dirname + '/' + current_file)
        elif ext == '.ppt':
            pptx = ppt2pptx(dirname + '/' + current_file)
            os.remove(dirname + '/' + current_file)
        elif ext == '.xls':
            xlsx = xls2xlsx(dirname + '/' + current_file)
            os.remove(dirname + '/' + current_file)
    
    print('Done!')  