import pandas as pd
import numpy as np
import os
import tkinter.messagebox
from tkinter import *
from tkinter import filedialog

dictionaryPath = './ExcelConverterForPriceSheets/Dictionary.xlsx'

def SearchInDictionary(excel):
    dictionary = pd.read_excel(dictionaryPath, usecols="A,B")
    counter = 0
    for dictionaryIndex in dictionary.index:
        for excelIndex in excel.index:
            if(isinstance(excel['Kod towaru'][excelIndex], str) and isinstance(dictionary['Zle'][dictionaryIndex], str) and excel['Kod towaru'][excelIndex] == dictionary['Zle'][dictionaryIndex]):
                print(excel['Kod towaru'][excelIndex] + " jest zly, zamieniam na: " + dictionary['Dobre'][dictionaryIndex])
                excel['Kod towaru'][excelIndex] = dictionary['Dobre'][dictionaryIndex]
                counter +=1
    print(counter)
                    
            

def Clips_clamps_export(path, output_file):
    excel = pd.read_excel(path, usecols = "C,E", header=None)
    excel.drop(excel.index[0:5], inplace=True)
    excel[2].replace('', np.nan, inplace=True)
    excel[2].replace('`', np.nan, inplace=True)
    excel.dropna(axis=0, inplace=True)

    excel.rename(columns={2: "Kod towaru", 4: "Cena"}, inplace=True)
    excel.insert(1,"Kod u kontrahenta","")
    excel.insert(2,"Kontrahent", "")

    SearchInDictionary(excel)
    excel.to_excel(output_file, index=False, sheet_name='Cennik')
    print(excel)

def Fabricated_parts_export(path, output_file):
    excel = pd.read_excel(path, usecols = "D,I", header=None)
    excel.drop(excel.index[0], inplace=True)
    excel[3].replace('', np.nan, inplace=True)
    excel[3].replace('`', np.nan, inplace=True)
    excel.dropna(subset=[3], inplace=True)

    excel.rename(columns={3: "Kod towaru", 8: "Cena"}, inplace=True)
    excel.insert(1,"Kod u kontrahenta","")
    excel.insert(2,"Kontrahent", "")

    SearchInDictionary(excel)
    excel.to_excel(output_file, index=False, sheet_name='Cennik')
    print(excel)

def Leg_and_Anchor_export(path, output_file):
    sheet0 = pd.read_excel(path,header=None, sheet_name=0)
    sheet0 = pd.concat([sheet0.iloc[[1]].T, sheet0.iloc[[42]].T], axis=1)
    sheet0.drop(sheet0.index[0:3], inplace=True)
    sheet0[1].replace('', np.nan, inplace=True)
    sheet0[1].replace('`', np.nan, inplace=True)
    sheet0.dropna(subset=[1], inplace=True)
    sheet0.rename(columns={1: "Kod towaru", 42: "Cena"}, inplace=True)

    sheet1 = pd.read_excel(path, header=None, sheet_name=1)
    sheet1 = pd.concat([sheet1.iloc[[2]].T, sheet1.iloc[[44]].T], axis=1)
    sheet1.drop(sheet1.index[0:1], inplace=True)
    sheet1[2].replace('', np.nan, inplace=True)
    sheet1[2].replace('`', np.nan, inplace=True)
    sheet1.dropna(subset=[2], inplace=True)
    sheet1.rename(columns={2: "Kod towaru", 44: "Cena"}, inplace=True)

    excel = pd.concat([sheet0,sheet1])
    excel.insert(1,"Kod u kontrahenta","")
    excel.insert(2,"Kontrahent", "")

    SearchInDictionary(excel)
    excel.to_excel(output_file, index=False, sheet_name='Cennik')
    print(excel)

def Struts_cut_lengths_export(path, output_file):
    excel = pd.read_excel(path, usecols = "A,H", header=None)
    excel.drop(excel.index[0:1], inplace=True)
    excel[0].replace('', np.nan, inplace=True)
    excel[0].replace('`', np.nan, inplace=True)
    excel.dropna(axis=0, inplace=True)

    excel.rename(columns={0: "Kod towaru", 7: "Cena"}, inplace=True)
    excel.insert(1,"Kod u kontrahenta","")
    excel.insert(2,"Kontrahent", "")

    SearchInDictionary(excel)
    excel.to_excel(output_file, index=False, sheet_name='Cennik')
    print(excel)

def CheckIfContains(string, containsArray):
    if(isinstance(containsArray, list)):
        checkedCounter = 0
        for x in containsArray:
            if(x in string):
                checkedCounter +=1
        if(checkedCounter == len(containsArray)):
            return True
        else:
            return False
    else:
        return False

def AddExportedStringToPath(path):
    return os.path.splitext(path)[0]+" Exported.xlsx"

def ShowMessage(path):
    tkinter.messagebox.showinfo(title=None, message="Wyexportowano plik: " + path)

def DetermineExcelTypeAndExport(path):
    if(CheckIfContains(path,["Clips", "clamps", "price", "template"])):
        Clips_clamps_export(path, AddExportedStringToPath(path))
        ShowMessage(path)
        return 0
    elif(CheckIfContains(path,["Fabricated", "Parts", "Prices"])):
        Fabricated_parts_export(path, AddExportedStringToPath(path))
        ShowMessage(path)
        return 1
    elif(CheckIfContains(path,["Leg", "and", "anchor", "template"])):
        Leg_and_Anchor_export(path, AddExportedStringToPath(path))
        ShowMessage(path)
        return 2
    elif(CheckIfContains(path,["Struts", "cut", "lenghts", "roller", "tubes", "prices"])):
        Struts_cut_lengths_export(path, AddExportedStringToPath(path))
        ShowMessage(path)
        return 3
    else:
        tkinter.messagebox.showerror(title="Blad", message="Nie znaleziono template'a")

def main(path):
    DetermineExcelTypeAndExport(path)


def OpenFileDialog():
    filepath = Tk()
    filepath.title("Export excela")
    filepath.name = filedialog.askopenfilename(initialdir="/", title="Select A File", filetypes=(("xlsx files", "*.xlsx"),("all files", "*.*")))
    main(filepath.name)
 
OpenFileDialog()