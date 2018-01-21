import openpyxl
import os
import datetime
import time


wb = openpyxl.load_workbook('UR.xlsx')

# print(type(wb))
# print(os.getcwd())
# print(wb.get_sheet_names())

sheet = wb.get_sheet_by_name('Sheet1')

# print(sheet)

print(sheet.title)


def get_sec(time_str):
    try:
        h, m, s = time_str.split(':')
        return int(h) * 3600 + int(m) * 60 + int(s)
    except ValueError:
        pass


def breaklunch(arg1):
    
    # 1st day
    
    if arg1 < 4:
        return 0
    elif arg1 < 6:
        return 0.35
    elif arg1 < 8:
        return 0.85
    elif arg1 < 12:
        return 1.1
    
    # 2nd day
    
    elif arg1 < 14:
        return 1.45
    elif arg1 < 16:
        return 1.95
    elif arg1 < 20:
        return 2.2
    elif arg1 < 22:
        return 2.55
    
    #3rd day
    
    elif arg1 < 24:
        return 3.05
    elif arg1 < 28:
        return 3.3
    elif arg1 < 30:
        return 3.65
    elif arg1 < 32:
        return 4.15


'''available_start = str(sheet.cell(row=9, column=16).value)
inbound_start = str(sheet.cell(row=9, column=17).value)
outbound_start = str(sheet.cell(row=9, column=18).value)
hold_start = str(sheet.cell(row=9, column=20).value)
total_time = str(sheet.cell(row=9, column=21).value)'''


holder = {}


for i in range(1, 77):
    name = str(sheet.cell(row = i + 9, column = 3).value)
    a = str(sheet.cell(row = i + 9, column = 16).value)
    b = str(sheet.cell(row = i + 9, column = 17).value)
    c = str(sheet.cell(row = i + 9, column = 18).value)
    d = str(sheet.cell(row = i + 9, column = 20).value)
    e = str(sheet.cell(row = i + 9, column = 21).value)
    aout = get_sec(a)
    bout = get_sec(b)
    cout = get_sec(c)
    dout = get_sec(d)
    eout = get_sec(e)
    try:
        result = round(((aout + bout + cout + dout) / (eout - (breaklunch(eout/3600)*3600))) * 100, 2)
    except TypeError:
        result = None
    print(result)
    holder[name] = result
    
print(holder)