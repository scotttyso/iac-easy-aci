#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
# from classes import access, admin, fabric, site_policies, system_settings
from class_tenants import tenants
from ordered_set import OrderedSet
import re

#======================================================
# Log Level - 0 = None, 1 = Class only, 2 = Line
#======================================================
log_level = 2

#=====================================================================
# Note: This is simply to make it so the classes don't appear Unused.
#=====================================================================
# class_list = [access, admin, fabric, site_policies, system_settings, tenants]

#======================================================
# Function to Count the Number of Keys/Columns
#======================================================
def countKeys(ws, func):
    count = 0
    for i in ws.rows:
        if any(i):
            if str(i[0].value) == func:
                count += 1
    return count

#======================================================
# Function to find the Keys for each Worksheet
#======================================================
def findKeys(ws, func_regex):
    func_list = OrderedSet()
    for i in ws.rows:
        if any(i):
            if re.search(func_regex, str(i[0].value)):
                func_list.add(str(i[0].value))
    return func_list

#======================================================
# Function to Create Terraform auto.tfvars files
#======================================================
def findVars(ws, func, rows, count):
    var_list = []
    var_dict = {}
    for i in range(1, rows + 1):
        if (ws.cell(row=i, column=1)).value == func:
            try:
                for x in range(2, 34):
                    if (ws.cell(row=i - 1, column=x)).value:
                        var_list.append(str(ws.cell(row=i - 1, column=x).value))
                    else:
                        x += 1
            except Exception as e:
                e = e
                pass
            break
    vcount = 1
    while vcount <= count:
        var_dict[vcount] = {}
        var_count = 0
        for z in var_list:
            var_dict[vcount][z] = ws.cell(row=i + vcount - 1, column=2 + var_count).value
            var_count += 1
        var_dict[vcount]['row'] = i + vcount - 1
        vcount += 1
    return var_dict

#======================================================
# Function to Read the Worksheet and Create Templates
#======================================================
def read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws):
    rows = ws.max_row
    func_list = findKeys(ws, func_regex)
    stdout_log(ws, None, 'begin')
    for func in func_list:
        count = countKeys(ws, func)
        var_dict = findVars(ws, func, rows, count)
        for pos in var_dict:
            row_num = var_dict[pos]['row']
            del var_dict[pos]['row']
            for x in list(var_dict[pos].keys()):
                if var_dict[pos][x] == '':
                    del var_dict[pos][x]
            stdout_log(ws, row_num, 'begin')
            var_dict[pos].update(
                {
                    'class_folder':class_folder,
                    'easyDict':easyDict,
                    'easy_jsonData':easy_jsonData,
                    'row_num':row_num,
                    'wb':wb,
                    'ws':ws
                }
            )
            easyDict = eval(f"{class_init}(class_folder).{func}(**var_dict[pos])")
    
    stdout_log(ws, row_num, 'end')
    # Return the easyDict
    return easyDict

#======================================================
# Function to Define stdout_log output
#======================================================
def stdout_log(ws, row_num, spot):
    if log_level == 0:
        return
    elif ((log_level == (1) or log_level == (2)) and
            (ws) and (row_num is None)) and spot == 'begin':
        print(f'-----------------------------------------------------------------------------\n')
        print(f'   Begin Worksheet "{ws.title}" evaluation...')
        print(f'\n-----------------------------------------------------------------------------\n')
    elif (log_level == (1) or log_level == (2)) and spot == 'end':
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Completed Worksheet "{ws.title}" evaluation...')
        print(f'\n-----------------------------------------------------------------------------')
    elif log_level == (2) and (ws) and (row_num is not None):
        if re.fullmatch('[0-9]', str(row_num)):
            print(f'    - Evaluating Row   {row_num}...')
        elif re.fullmatch('[0-9][0-9]',  str(row_num)):
            print(f'    - Evaluating Row  {row_num}...')
        elif re.fullmatch('[0-9][0-9][0-9]',  str(row_num)):
            print(f'    - Evaluating Row {row_num}...')
    else:
        return

