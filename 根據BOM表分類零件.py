import os
from pathlib import Path
import numpy as np
import argparse
import shutil
import tkinter as tk
from tkinter import simpledialog
from tkinter import filedialog
import pandas as pd
import xml.etree.ElementTree as ET

def getfolderpath(type): 
    root = tk.Tk()
    root.withdraw()
    if type == 1: # 文字式輸入讀資料夾
        folder_path = simpledialog.askstring("輸入資料夾路徑", "請輸入資料夾完整路徑：")
    elif type == 2: # GUI式選擇讀資料夾
        folder_path = filedialog.askdirectory(title="請選擇資料夾")
    root.destroy()
    return folder_path

def getfiles(type):
    root = tk.Tk()
    root.withdraw()
    if type == 1:
        ftype = ("excel files", "*.xlsx")
    elif type == 2:
        ftype = ("text files", "*.txt")
    elif type == 3:
        ftype = ("all files", "*.*")
    file_path = filedialog.askopenfilename(
        title="請選擇檔案",
        filetypes=[ftype],
    )
    if not file_path:
        print("未選取任何檔案!")
        return
    return file_path

def get_filenames(filespath, files_type):
    allfiles = os.listdir(filespath)
    filtered_files = [file for file in allfiles if file.endswith(files_type)]
    print("你輸入的檔案有：", filtered_files)
    stripped_files = [os.path.splitext(file)[0] for file in filtered_files]
    return stripped_files

def get_excel_fieldslist(path, *cols):
    df = pd.read_excel(path)
    
    missing_cols = [c for c in cols if c not in df.columns]
    if missing_cols:
        raise KeyError(f"欄位不存在: {missing_cols}")    
    selected = df[list(cols)]
    values = selected.values  
    return values

cols = ["零件名稱", "加工方法"]  # 欄位可自由增減
bom_path = getfiles(1)
folder_path = getfolderpath(2)
df = get_excel_fieldslist(bom_path, "零件名稱", "加工方法")
files_type = ".sldprt"
namelist = get_filenames(folder_path, files_type)
class1 = folder_path + "板金件/"
class2 = folder_path + "車床件/"
class3 = folder_path + "焊接件/"
for i in df:
    part, method = i
    for name in namelist:
        if part == name:
            if method == "S":
                os.makedirs(class1, exist_ok=True)
                shutil.move(folder_path + name + files_type, class1)
                print(name + "已分配至板金件")
            elif method == "T":
                os.makedirs(class2, exist_ok=True)
                shutil.move(folder_path + name + files_type, class2)
                print(name + "已分配至車床件")
            elif method == "C":
                os.makedirs(class3, exist_ok=True)
                shutil.move(folder_path + name + files_type, class3)
                print(name + "已分配至焊接件")
            else:
                print(name + "是未知類別")
                continue
        else:

            continue
