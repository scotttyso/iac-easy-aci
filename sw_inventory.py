#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from copy import deepcopy
from dotmap import DotMap
from openpyxl import Workbook
import argparse
import classes
import easy_functions
import json
import platform
import os
import sys
import re

#======================================================
# Workbook Styles
#======================================================
wbstyles = easy_functions.workbook_styles()

#======================================================
# Function to Create Switch Dictionary
#======================================================
def create_switch_dictionary(args, nodeData, topData):
    swDict = {}
    for i in topData['imdata']:
        for key, value in i.items():
            a = DotMap(value['attributes'])
            newdn = (a.dn).strip('/sys')
            swDict.update({newdn:dict(
                intfs   = {},
                model   = "",
                name    = a.name,
                pod     = a.podId,
                role    = a.role,
                serial  = a.serial,
                version = a.version
            )})
            intfDict = {}
            if value.get('children'):
                for c in value['children']:
                    for k, v in c.items():
                        aa = DotMap(v['attributes'])
                        admin_status = aa.adminSt
                        if args.function_type == 'api':
                            intf = re.search('\\[(.*)\\]', aa.rn).group(1)
                        else:
                            intf = re.search('\\[(.*)\\]', aa.dn).group(1)
                        if re.search('/(\\d)$', intf):
                            intf = intf.split('/')[0] + '/' + '0' + re.search('/(\\d)$', intf).group(1)
                        for y, z in v['children'][0].items():
                            aaa = DotMap(z['attributes'])
                            duplex = aaa.operDuplex
                            speed  = aaa.operSpeed
                            state  = aaa.operSt
                            optic  = aaa.operStQual
                            if not optic == 'sfp-missing':
                                for w, x in z['children'][0].items():
                                    aaaa = DotMap(x['attributes'])
                                    optic = aaaa.typeName
                        intfDict.update(deepcopy({intf:{
                            'admin_status': admin_status,
                            'duplex': duplex,
                            'optic': optic,
                            'speed': speed,
                            'state': state,
                        }}))
                swDict[newdn]['intfs'].update(intfDict)
    for i in nodeData['imdata']:
        for k, v in i.items():
            a = DotMap(v['attributes'])
            newdn = (a.dn).strip('/sys')
            swDict[newdn].update(deepcopy({'model':a.model}))
    swDict = dict(sorted(swDict.items()))
    for k, v in swDict.items():
        v['intfs'] = dict(sorted(v['intfs'].items()))
    #print(json.dumps(swDict, indent=4))
    return(swDict)


#======================================================
# Function to Create the Workbook
#======================================================
def create_workbook_from_data(swDict):
    dest_file = 'switch_inventory.xlsx'
    wbstyles = easy_functions.workbook_styles()
    wb = Workbook()
    wb.add_named_style(wbstyles.wsh1)
    wb.add_named_style(wbstyles.wsh2)
    wb.add_named_style(wbstyles.ws_odd)
    wb.add_named_style(wbstyles.ws_even)
    ws1 = wb.active
    ws1.title = 'Switches'
    ws2 = wb.create_sheet(title = 'Interfaces')
    
    # Populate Switch Worksheet
    data = ['Pod', 'Switch Name','Model','Serial', 'Role', 'Version']
    for x in range(0,len(data)):
        ltr = chr(ord('@')+(x+1))
        ws1.column_dimensions[ltr].width = 30
    ws1.append(data)
    for cell in ws1["1:1"]: cell.style = 'wsh2'

    ws_count = 2
    for k, v in swDict.items():
        v = deepcopy(DotMap(v))
        if not v.role == 'controller':
            data = [v.pod, v.name, v.model, v.serial, v.role, v.version]
            ws1.append(data)
            rc = '%s:%s' % (ws_count, ws_count)
            for cell in ws1[rc]:
                if ws_count % 2 == 0: cell.style = 'ws_even'
                else: cell.style = 'ws_odd'
            ws_count += 1

    # Populate Interface Worksheet
    data = ['Pod', 'Switch Name','Interface','Optic', 'Admin Status', 'State', 'Speed', 'Duplex']
    for x in range(0,len(data)):
        ltr = chr(ord('@')+(x+1))
        ws2.column_dimensions[ltr].width = 30
    ws2.append(data)
    for cell in ws2["1:1"]: cell.style = 'wsh2'

    ws_count = 2
    for k, v in swDict.items():
        v = deepcopy(DotMap(v))
        if not v.role == 'controller':
            for kk, vv in v.intfs.items():
                a = deepcopy(DotMap(vv))
                data = [v.pod, v.name, kk, a.optic, a.admin_status, a.state, a.speed, a.duplex]
                ws2.append(data)
                rc = '%s:%s' % (ws_count, ws_count)
                for cell in ws2[rc]:
                    if ws_count % 2 == 0: cell.style = 'ws_even'
                    else: cell.style = 'ws_odd'
                ws_count += 1

    wb.save(dest_file)

#======================================================
# Function to Create Switch Dictionary
#======================================================
def define_arguments():
    # User Input Arguments
    Parser = argparse.ArgumentParser(description='Configuration Migration')
    Parser.add_argument('-f', '--function_type',
                        required=True,
                        help = 'Options are: '\
                            '1. "api" or '\
                            '2. "files".'
    )
    Parser.add_argument('-n', '--node_input',
                        action='store',
                        required='files' in sys.argv,
                        help = 'The Source file for fabricNode dump.'
    )
    Parser.add_argument('-s', '--server',
                        action='store',
                        required='api' in sys.argv,
                        help = 'The Server Name for the APIC.'
    )
    Parser.add_argument('-t', '--top_system_input',
                        action='store',
                        required='files' in sys.argv,
                        help = 'The Source file for topSystem dump.'
    )
    Parser.add_argument('-u', '--username',
                        action='store',
                        default = 'admin',
                        required='api' in sys.argv,
                        help = 'The Login Username for the APIC Host.'
    )

    args = Parser.parse_args()
    return args

#======================================================
# Main Module
#======================================================
def main():
    args = define_arguments()

    opSystem = platform.system()
    kwargs = {}
    kwargs['args'] = args
    script_path = os.path.dirname(os.path.realpath(sys.argv[0]))
    if opSystem == 'Windows': path_sep = '\\'
    else: path_sep = '/'
    jsonFile = f'{script_path}{path_sep}templates{path_sep}variables{path_sep}easy_variables.json'
    jsonOpen = open(jsonFile, 'r')
    easy_jsonData = json.load(jsonOpen)
    jsonOpen.close()
    easy_jsonData = easy_jsonData['components']['schemas']
    easyDict = {}
    kwargs = {
        'args':args,
        'easyDict':easyDict,
        'easy_jsonData':easy_jsonData,
        'remove_default_args':False,
    }
    #======================================================
    # If User Inputs Server Argument, Get Data From API
    #======================================================
    if not args.server == None:
        kwargs['jsonData'] = kwargs['easy_jsonData']['site.Identifiers']['allOf'][1]['properties']
        kwargs["Variable"] = 'apicPass'
        args.apic_pass = easy_functions.sensitive_var_value(**kwargs)
        fablogin = classes.apicLogin(args.server, args.username, args.apic_pass)
        cookies = fablogin.login()

        # Switch Dictionaries
        swDict      = {}
        tempfile    = 'dummy.json'
        uri         = 'api/node/class/fabricNode'
        nodeResponse= easy_functions.apic_api(args.server, 'get', {}, cookies, uri, tempfile)
        uri         = 'api/node/class/topSystem'
        uriFilter   = "rsp-subtree=full&rsp-subtree-class=fabricNode&rsp-subtree-class=l1PhysIf&rsp-subtree-class=ethpmPhysIf&rsp-subtree-class=ethpmFcot"
        topResponse = easy_functions.apic_api_with_filter(args.server, cookies, uri, uriFilter)
        nodeData    = nodeResponse.json()
        topData     = topResponse.json()

        swDict = create_switch_dictionary(args, nodeData, topData)
        create_workbook_from_data(swDict)


    #======================================================
    # If User Does Not Assign Server Argument use Files
    #======================================================
    elif not args.node_input == None and not args.top_system_input == None:
        #file = open(args.node_input, 'r')
        topData = json.load(open(args.top_system_input, 'r'))
        nodeData = json.load(open(args.node_input, 'r'))

        swDict = create_switch_dictionary(args, nodeData, topData)
        create_workbook_from_data(swDict)


if __name__ == '__main__':
    main()
