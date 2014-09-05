#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd

w = xlrd.open_workbook(u'2.xlsm')
data_sheet = w.sheet_by_name(u'data_table')

def get_column_names(xl_sheet, name):
    row = xl_sheet.row(0) # 1st row
    for idx, cell_obj in enumerate(row):
        if cell_obj.value == name:
            return idx

gender_col = get_column_names(data_sheet, 'gender') # 1/2
age_col = get_column_names(data_sheet, 'age') # 1/2
cid_col    = get_column_names(data_sheet, 'con_3')
tnb_col = get_column_names(data_sheet, 'ques_7_2') # 0/1
smoke_col = get_column_names(data_sheet, 'ques_9_1') # 0/1
danguchun_col = get_column_names(data_sheet, 'ques_10_2')
sbp_col = get_column_names(data_sheet, 'ques_10_1_1')

layer = {}
#gender age tnb smoke
layer['1']['4']['1']['0'] = []

def get(row):
    data['gender'] = data_sheet.cell_value(row, gender_col)
    data['age'] = data_sheet.cell_value(row, age_col)
    data['tnb'] = data_sheet.cell_value(row, tnb_col)
    data['smoke'] = data_sheet.cell_value(row, smoke_col)
    data['danguchun'] = data_sheet.cell_value(row, danguchun_col)
    data['sbp'] = data_sheet.cell_value(row, sbp_col)

    for k,v in data:
        if v == 'missing':
            return None


    pass

if __name__ == '__main__':
    get(1)


                