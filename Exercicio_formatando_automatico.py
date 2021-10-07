#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import xlsxwriter as opcoesDOXL

import os

nomeArquivo = 'C:\\Users\\User\\Desktop\\RPA1\\xlsx\\MergeCells.xlsx'
workbook = opcoesDOXL.Workbook(nomeArquivo)

sheetPadrao = workbook.add_worksheet()

add_merge_celulas = workbook.add_format({
    'bold': True, 
    'border': 6,
    'valign': 'vcenter',
    'size': 30,
    'fg_color': 'blue',
    'font_color': 'white',
    
})

sheetPadrao.merge_range('B3:I5', 'Merge Cells',add_merge_celulas)
workbook.close()

os.startfile(nomeArquivo)

