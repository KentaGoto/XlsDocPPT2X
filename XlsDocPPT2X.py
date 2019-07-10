# coding: utf-8

import os
import shutil
import pathlib
import glob
import re
import pprint
import subprocess
import win32com
from win32com.client import *
import time

def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)

def whitespaceRename(root_dir, spc_flag, dir_flag):
    p_root_dir = pathlib.Path(root_dir)

    pattern = re.compile(r'\s') # マッチさせたい正規表現パターン
    replace = '_' # 置き換える文字

    # ファイルを処理する
    for full_path in list(p_root_dir.glob('**/*')):
        # フルパス、親ディレクトリ、ファイル名およびディレクトリ名を文字列（型）にする
        full_path_str = full_path.__str__()
        parentDir = full_path.parent.__str__()
        fileName = full_path.name.__str__()
        matchObj = re.search(pattern, fileName) # マッチオブジェクト
        if matchObj:
            # ディレクトリの場合はフラグを付けて次にスキップ
            if os.path.isdir(full_path_str):
                dir_flag += 1
                continue

            fileName_r = re.sub(pattern, replace, fileName)
            os.rename(parentDir + '/' + fileName, parentDir + '/' + fileName_r)
            spc_flag += 1

    # ディレクトリを処理する
    if dir_flag > 0:
        for full_path in list(p_root_dir.glob('**/*')):
            # フルパス、親ディレクトリ、ファイル名およびディレクトリ名を文字列（型）にする
            full_path_str = full_path.__str__()
            parentDir = full_path.parent.__str__()
            fileName = full_path.name.__str__()
            matchObj = re.search(pattern, fileName) # マッチオブジェクト
            if matchObj:
                fileName_r = re.sub(pattern, replace, fileName)
                os.rename(parentDir + '/' + fileName, parentDir + '/' + fileName_r)
                spc_flag += 1
    
    return spc_flag

def doc2docx(doc_fullpath):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc_fullpath = doc_fullpath.replace("\\", "/")
    print(doc_fullpath)
    dirname = os.path.dirname(doc_fullpath)
    current_file = os.path.basename(doc_fullpath)
    fname, ext = os.path.splitext(current_file)
    os.chdir(dirname)
    doc = word.Documents.Open(doc_fullpath)
    # docxに変換する
    # FileFormatのValue値は「https://msdn.microsoft.com/VBA/Word-VBA/articles/wdsaveformat-enumeration-word?f=255&MSPPError=-2147217396」を参照
    doc.SaveAs(dirname + '/' + fname + '.docx', FileFormat=16)
    doc.Close()
    # word.Quit()  # releases Word object from memory
    return dirname + '/' + fname + '.docx'

def ppt2pptx(ppt_fullpath):
    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    powerpoint.DisplayAlerts = 0
    ppt_fullpath = ppt_fullpath.replace("\\", "/")
    print(ppt_fullpath)
    dirname = os.path.dirname(ppt_fullpath)
    current_file = os.path.basename(ppt_fullpath)
    fname, ext = os.path.splitext(current_file)
    ppt = powerpoint.Presentations.Open(ppt_fullpath, False, False, False)
    # pptxに変換する
    ppt.SaveAs(dirname + '/' + fname + '.pptx')
    ppt.Close()
    # powerpoint.Quit()
    return dirname + '/' + fname + '.pptx'

def xls2xlsx(xls_fullpath):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = 0
    # xls_fullpath = xls_fullpath.replace("\\", "/")
    xls_fullpath = xls_fullpath.replace("/", "\\")  # Excel操作するときは\じゃないとダメみたい
    print(xls_fullpath)
    dirname = os.path.dirname(xls_fullpath)
    current_file = os.path.basename(xls_fullpath)
    fname, ext = os.path.splitext(current_file)
    xls = excel.Workbooks.Open(xls_fullpath)
    # xlsxに変換する
    xls.SaveAs(dirname + '/' + fname + '.xlsx', FileFormat=51)
    xls.Close()
    # excel.Quit()
    return dirname + '/' + fname + '.xlsx'

if __name__ == '__main__':
    root_dir = input("Dir: ")
    root_dir_copy = root_dir + '__copy'
    shutil.copytree(root_dir, root_dir_copy)

    spc_flag = 0
    dir_flag = 0

    # 前処理: 空白文字を置換
    spc_flag = whitespaceRename(root_dir_copy, spc_flag, dir_flag)

    # doc、ppt、xlsをそれぞれdocx、pptx、xlsxに変換する
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

    print('')
    # フォルダ/ファイルに空白があった場合
    if spc_flag > 0:
        print('Because there was a space in the folder / file, it was replaced with underscores.')
    
    print('Done!')