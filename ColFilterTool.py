# -*- coding: utf-8 -*-
"""
Created on Mon Mar 12 10:01:01 2019

@author: Girvan Tse

"""
import re
from PySimpleGUI import Text, FileBrowse, Input, Window, Popup, Submit, Cancel, Checkbox, Button, Column
from tkinter import TclError
from pandas import ExcelWriter, DataFrame, read_excel
from xlrd import XLRDError

layout =   [[Text('File to Query')],

            [Input('[Path to Excel Workbook]',
                  key = 'path'),
             FileBrowse(file_types=(("Excel Workbook",
                                     "*.xlsx"),
                                    ("All Files",
                                     "*.*")
                                    )),],
            [Input('[Sheet Name]',
                   key = 'sheet',
                   size = (40, 0)), 
             Input('[Skip # Rows]',
                  key = 'rows',
                  size = (12, 0))],
            [Submit(key = 'next'), Cancel(key = 'exit')]]

layout2 =  [[Text('Filter which Columns')]]

window = Window('ColFilterTool' ).Layout(layout)

def validate(file):
    try:
        _testParam = read_excel(file[0], 
                                sheet_name = file[1])
    except FileNotFoundError:
        return 0
    except XLRDError:
        return 0
    return 1

RunTool = True
while RunTool:
    try:
        event, values = window.Read()
    except TclError:
        pass
    if (event is None or event == 'exit'):
        RunTool = False
    if (event is 'next' and validate([values['path'], values['sheet']])):
        window.Close()
        try:
            skiprows = int(values['rows'])
        except:
            skiprows = 0
        queryFrame = read_excel(values['path'],
                                sheet_name = values['sheet'],
                                skiprows = skiprows)
        dropCols = list()
        for column in queryFrame.columns:
            if (column.startswith('Unnamed: ')):
                dropCols.append(column)
        for colName in dropCols:
            queryFrame.drop(colName, axis=1, inplace=True)
        PATH = values['path']
        headerList = list()
        checkList = list()
        for header in queryFrame.columns:
            headerList.append(header)
            checkList.append([Checkbox(header)])
        layout2.append([Column(checkList, size = (400, 500), scrollable = True)])
        layout2.append([Submit(key = 'next1'), Cancel(key = 'exit')])
        window = Window('ColFilterTool').Layout(layout2)
    if (event is 'next1'):
        RunTool = False
        window.Close()
        for i in range(0, len(values)):
            if (values[i] == False):
                queryFrame.drop(headerList[i], axis=1, inplace=True)
        writer = ExcelWriter(PATH[:-5] + " OUTPUT.xlsx",
                             engine = 'xlsxwriter')
        queryFrame.to_excel(writer, sheet_name = 'Output', index = False)        
        writer.save()
        writer.close()
        Popup('Successful Execution!')
