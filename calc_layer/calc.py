# -*- coding: utf-8 -*-

import xlrd

w = xlrd.open_workbook(u'3.xlsm')
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
layer_col = get_column_names(data_sheet, 'ques_10_16')
jiangya_col = get_column_names(data_sheet, 'ques_7_6_1')
jiangzhi_col = get_column_names(data_sheet, 'ques_7_6_2')

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

def get_danguchun(d):
    for x in xrange(5,9):
        if d < x:
            return x - 5
        pass
    return 4
    pass

def get_sbp(d):
    if d <= 139:
        return 0
    elif d <= 159:
        return 1
    elif d <= 179:
        return 2
    else:
        return 3
    pass

def get_in_group(group, d):
    dgc = get_danguchun(d['danguchun'])
    sbp = get_sbp(d['sbp'])
    # print dgc, sbp
    return group[(3 - sbp) * 5 + dgc]
    pass

def get_layer(d, row):
    age = str((int(d['age']))/10)
    if int(age) > 7:
        age = '7'
    if int(age) < 4:
        age = '4'
    group = layer[str(int(d['gender']))][age]
    # print d
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
    # print g
    return get_in_group(g, d)

def get(row):
    # data = {'gender':2,'age':48, 'smoke':0, 'tnb':0, 'danguchun':7, 'sbp':181}
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
    if data['gender'] == '' or data['age'] == '':
            return None
    layer = get_layer(data, row)
    jy = data_sheet.cell_value(row, jiangya_col)
    jz = data_sheet.cell_value(row, jiangzhi_col)
    if jy or jz:
        layer = layer + 1
        if layer > 4:
            layer = 4

    return layer
    pass


if __name__ == '__main__':
    f = open('list.txt', 'w')
    for x in xrange(1,102054):
        res = get(x)
        if res != None:
            res = res + 1
        else:
            res = ''
        f.write(str(res) + '\n')
        pass
    f.close()
