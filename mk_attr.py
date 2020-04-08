import sys
import os
import re
import xlrd
import itertools


from tkinter import filedialog as tkd
file_path_string = tkd.askopenfilename()
book = xlrd.open_workbook(file_path_string)

def xls2lists(book_argument):
    sheets = book_argument.sheet_names()
    base_sheet_variants = ['iis', 'main']
    coding_sheet_variants = ['кодировка']
    for sheet in sheets:
        if any(x in str(sheet).strip().lower() for x in base_sheet_variants):
            base_sheet = sheet
        if any(y in str(sheet).strip().lower() for y in coding_sheet_variants):
            coding_sheet = sheet
    base_heading = book.sheet_by_name(base_sheet).row_values(0)
    codings = [book.sheet_by_name(coding_sheet).row_values(row_index) for row_index in range(book.sheet_by_name(coding_sheet).nrows)]
    return base_heading, codings

def makecat(inputarr):
    result = {}
    for x in inputarr:
        if x != '':
            try:
                code = int(x)
                if 'code' not in result:
                    result.update({'code':code})
            except:
                label = str(x)
                if 'label' not in result:
                    result.update({'label':label})
    return result if ('code' in result and 'label' in result) else None   

bh, cds = xls2lists(book)
print(bh,'\n',cds)
borders = {cds.index(element):element[0] for element in cds if element[0]!=''}

bh_lower = [bh_elem.lower() for bh_elem in bh]
index_of_split = bh_lower.index('sex')
vars_cat = bh[index_of_split:]
vars_oe = bh[1:index_of_split]

print('\n',borders,'\n\n',borders.values(),'\n\n',vars_cat,'\n\n')

with open(file_path_string[:file_path_string.rfind('/')+1]+"attributs.txt", "w", encoding="utf-8") as f:
    default_dict = [{'code':'0','label':'Нет'},{'code':'1','label':'Да'}] 
    for x in vars_cat: 
        if x.strip().lower() in borders.values() or x in borders.values():
            try:
                startindex = list(borders.keys())[list(borders.values()).index(x.strip().lower())]
                tmpx = x.strip().lower()
            except:
                startindex = list(borders.keys())[list(borders.values()).index(x)]
                tmpx = x
            try:
                lastindex = list(borders.keys())[list(borders.values()).index(tmpx) + 1]
                custom_attr = cds[startindex:lastindex]
            except:
                lastindex = None
                custom_attr = cds[startindex:]
            print(x, ', ', startindex,', ',lastindex)
            print([makecat(y[1:]) for y in custom_attr])
            f_output = '\t' + x.strip().lower() + ' \"\"\n\tcategorical[1..1]\n\t{\n'+',\n'.join(['\t\t_' + str(z['code']) + ' \"' + str(z['label']) + '\"' for z in [makecat(y[1:]) for y in custom_attr if makecat(y[1:])!=None]])+'\n\t};\n\n'
            f.write(f_output)
        else:
            f_output = '\t' + x.strip().lower() + ' \"\"\n\tcategorical[1..1]\n\t{\n'+',\n'.join(['\t\t_' + str(z['code']) + ' \"' + str(z['label']) + '\"' for z in default_dict])+'\n\t};\n\n'
            f.write(f_output)
    
    for oe_x in vars_oe:
        f_output = '\t\t\tif len(iom.SampleRecord[\"'+ str(oe_x)+'\"])>0 then\n\t\t\t\tGetOpenFromSample(IOM,\"'+ str(oe_x) +'\")\n\t\t\tend if\n\n'
        f.write(f_output)
    
    vars_cat.append('qfr')
    for cat_x in vars_cat:
        f_output = '\t\t\tif len(iom.SampleRecord[\"'+ str(cat_x)+'\"])>0 then\n\t\t\t\tGetCategoricalFromSample(IOM,\"'+ str(cat_x) +'\", \"_\")\n\t\t\tend if\n\n'
        f.write(f_output)

    vars_cat.remove('qfr')
    for cat_x in vars_cat:
        f_output = '\t\t\t' + str(cat_x).strip().lower() + '.Banners.AddNew("pictureTop","<span style=\'color:red;margin-left:25px;\'>Screen shown only for test IDs</span>")\n\t\t\t' + str(cat_x).strip().lower() +'.Ask()\n\n'
        f.write(f_output)