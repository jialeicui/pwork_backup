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

layer = {'1':{}, '2':{}, '3':{}, '4':{}}
#gender age tnb smoke
layer['1']['4'] = [[2,3,4,4,4,0,1,1,2,4,0,0,0,0,1,0,0,0,0,1,],[4,4,4,4,4,1,2,3,3,4,0,0,0,1,2,0,0,0,0,1,],[1,2,3,4,4,0,1,1,1,2,0,0,0,0,1,0,0,0,0,0,],[3,4,4,4,4,0,1,1,2,4,0,0,0,0,2,0,0,0,0,0,],]
layer['1']['5'] = [[3,4,4,4,4,1,1,1,3,4,0,0,0,1,1,0,0,0,0,1,],[4,4,4,4,4,2,3,4,4,4,0,1,1,2,3,0,0,0,0,1,],[2,2,3,4,4,0,0,1,1,3,0,0,0,0,1,0,0,0,0,0,],[4,4,4,4,4,1,1,2,3,4,0,0,0,1,2,0,0,0,0,1,],]
layer['1']['6'] = [[4,4,4,4,4,3,4,4,4,4,2,2,3,3,4,1,1,1,2,2,],[4,4,4,4,4,4,4,4,4,4,3,4,4,4,4,1,2,2,3,4,],[3,4,4,4,4,1,1,2,2,3,0,0,0,1,1,0,0,0,0,0,],[4,4,4,4,4,2,2,3,4,4,1,1,1,2,3,0,0,0,1,1,],]
layer['1']['7'] = [[4,4,4,4,4,2,2,3,4,4,1,1,2,2,3,0,0,1,1,1,],[4,4,4,4,4,3,4,4,4,4,2,2,3,4,4,1,1,2,2,2,],[3,4,4,4,4,1,2,2,3,3,1,1,1,1,2,0,0,0,1,1,],[4,4,4,4,4,2,3,4,4,4,1,2,3,3,3,0,1,1,1,2,],]
layer['2']['4'] = [[2,4,4,4,4,0,0,1,2,3,0,0,0,0,2,0,0,0,0,0,],[4,4,4,4,4,0,2,2,4,4,0,0,0,1,3,0,0,0,0,0,],[1,2,3,4,4,0,0,0,1,2,0,0,0,0,1,0,0,0,0,0,],[2,3,4,4,4,0,0,1,2,4,0,0,0,0,3,0,0,0,0,0,],]
layer['2']['5'] = [[3,4,4,4,4,0,1,1,2,4,0,0,0,1,2,0,0,0,0,0,],[4,4,4,4,4,1,2,3,4,4,0,0,1,2,3,0,0,0,0,1,],[1,2,3,4,4,0,0,0,1,2,0,0,0,0,1,0,0,0,0,0,],[3,4,4,4,4,0,1,1,2,4,0,0,0,1,2,0,0,0,0,0,],]
layer['2']['6'] = [[4,4,4,4,4,2,2,3,4,4,1,1,1,2,3,0,0,0,1,1,],[4,4,4,4,4,2,3,4,4,4,1,1,2,2,4,0,0,1,1,2,],[2,2,3,4,4,0,0,1,1,2,0,0,0,0,1,0,0,0,0,0,],[3,4,4,4,4,1,1,2,2,4,0,0,1,1,2,0,0,0,0,1,],]
layer['2']['7'] = [[4,4,4,4,4,2,2,3,4,4,1,1,2,2,3,0,0,1,1,1,],[4,4,4,4,4,3,4,4,4,4,2,2,3,4,4,1,1,2,2,2,],[2,3,3,4,4,1,1,1,1,2,0,0,0,1,1,0,0,0,0,0,],[3,4,4,4,4,1,2,2,3,4,1,1,1,1,2,0,0,0,1,1,],]

def get_layer(d):
    age = str((int(d['age']))/10)
    group = layer[str(int(d['gender']))][age]
    print d
    if d['tnb']:
        if not d['smoke']:
            g = group[0]
        else:
            g = group[1]
    else:
        if not d['smoke']:
            g = group[2]
        else:
            g = group[3]
    pass
    
    print g

def get(row):
    data = {}
    data['gender'] = data_sheet.cell_value(row, gender_col)
    data['age'] = data_sheet.cell_value(row, age_col)
    data['tnb'] = data_sheet.cell_value(row, tnb_col)
    data['smoke'] = data_sheet.cell_value(row, smoke_col)
    data['danguchun'] = data_sheet.cell_value(row, danguchun_col)
    data['sbp'] = data_sheet.cell_value(row, sbp_col)

    for k,v in data.items():
        if v == u'missing':
            return None

    return get_layer(data)
    pass

if __name__ == '__main__':
    for x in xrange(1,10):
        get(x)
        pass
