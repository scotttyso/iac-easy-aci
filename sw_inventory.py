#!/usr/bin/env python3

from copy import deepcopy
from dotmap import DotMap
from openpyxl import Workbook
import argparse
import classes
import easy_functions
import json
import platform
import os
import re

wbstyles = easy_functions.workbook_styles()

def main():
    # User Input Arguments
    Parser = argparse.ArgumentParser(description='Configuration Migration')
    Parser.add_argument('-n', '--node_input',
                        action='store',
                        required=False,
                        help = 'The Login Name for the APIC Host.'
    )
    Parser.add_argument('-t', '--top_system_input',
                        action='store',
                        required=False,
                        help = 'The Login Name for the APIC Host.'
    )
    Parser.add_argument('-s', '--server',
                        action='store',
                        required=False,
                        help = 'The Server Name for the APIC.'
    )
    Parser.add_argument('-u', '--username',
                        action='store',
                        default = 'admin',
                        required=False,
                        help = 'The Login Name for the APIC Host.'
    )
    args = Parser.parse_args()
    polVars = {}
    if not args.server == None:
        polVars["Variable"] = 'apicPass'
        args.apic_pass = easy_functions.sensitive_var_value(**polVars)
        fablogin = classes.apicLogin(args.server, args.username, args.apic_pass)
        cookies = fablogin.login()

        # Switch Dictionaries
        swDict = {}
        tempfile = 'dummy.json'
        uri   = 'api/node/class/fabricNode'
        uri_response = easy_functions.apic_api(args.server, 'get', {}, cookies, uri, tempfile)
        nodeData = uri_response.json()
        for item in nodeData['imdata']:
            item = DotMap(item)
            i = item.fabricNode.attributes
            if not i.role == 'controller':
                swDict.update({i.dn:dict(
                    dn      = i.dn,
                    intfs   = {},
                    model   = i.model,
                    name    = i.name,
                    role    = i.role,
                    serial  = i.serial,
                    version = i.version
                )})
                uri = f'api/mo/{i.dn}'
                uriFilter = f'query-target=subtree&target-subtree-class=ethpmFcot'
                apiResult = easy_functions.apic_api_with_filter(args.server, cookies, uri, uriFilter, tempfile)
                ethpmFcot = apiResult.json()
                print(json.dumps(ethpmFcot.json(), indent=4))
                uriFilter = f'query-target=subtree&target-subtree-class=ethpmPhysIf'
                apiResult = easy_functions.apic_api_with_filter(args.server, cookies, uri, uriFilter, tempfile)
                ethpmPhysIf = apiResult.json()
                print(json.dumps(ethpmPhysIf.json(), indent=4))

        # Interface SFP/QSFP Data
        uri = 'api/node/class/ethpmFcot'
        uri_response = easy_functions.apic_api(args.server, 'get', {}, cookies, uri, tempfile)
        opticData = {}
        uriData = uri_response.json()
        print(json.dumps(uriData['imdata'], indent=4))
        for item in uriData['imdata']:
            item = DotMap(item)
            i = item.ethpmFcot.attributes
            dnnew = i.dn.replace('/phys/fcot', '')
            opticData.update({dnnew:dict(
                dn = i.dn,
                type = i.typeName
            )})
        # Interface Status
        statusData   = {}
        uri          = '/api/node/class/ethpmPhysIf'
        uri_response = easy_functions.apic_api(args.server, 'get', {}, cookies, uri, tempfile)
        uriData      = uri_response.json()
        #print(json.dumps(uriData['imdata'], indent=4))
        for item in uriData['imdata']:
            item = DotMap(item)
            i = item.ethpmPhysIf.attributes
            dnnew = i.dn.replace('/phys', '')
            statusData.update({dnnew:dict(
                dn = i.dn,
                accessVlan = i.accessVlan,
                allowedVlans = i.allowedVlans,
                operDuplex   = i.operDuplex,
                operState    = i.operSt,
                operReason   = i.operStQual,
                operSpeed    = i.operSpeed,
            )})
        #print(json.dumps(statusData, indent=4))
        for key, value in swDict.items():
            for k, v in statusData.items():
                i = DotMap(i)
                x = k.split('/')
                newk = x[0] + '/' + x[1] + '/' + x[2]
                if key == newk:
                    v = DotMap(v)
                    intf = re.search('\\[(eth[\\d/]+)\\]', v.dn).group(1)
                    if re.search('eth1/(\\d)$', intf):
                        intf = 'eth1/0{}'.format(re.search('eth1/(\\d)$', intf).group(1))
                    idict = dict(
                        acessVlan         = v.accessVlan,
                        allowedVlans      = v.allowedVlans,
                        operationalDuplex = v.operDuplex,
                        operationalSpeed  = v.operSpeed,
                        operationalState  = v.operState,
                    )
                    if re.search('(link-not-connected|none)', v.operReason):
                        newkey = k.replace('/sys-', '/sys/phys-')
                        idict.update(deepcopy({'optic':opticData[newkey]['type']}))
                    else: idict.update(deepcopy({'optic':v.operReason}))
                    swDict[key]['intfs'].update(deepcopy({intf:idict}))
        #swDict = dict(sorted(swDict))
        for k, v in swDict.items():
            v['intfs'] = dict(sorted(v['intfs'].items()))
        swDict = dict(sorted(swDict.items()))
        print(json.dumps(swDict, indent=4))
        mainDict = {}
        for key, value in swDict.items():
            upcount = 0
            downcount = 0
            for k, v in value['intfs'].items():
                if v['operationalState'] == 'up': upcount += 1
                else: downcount += 1
            mainDict.update({value['name']:{
                'node_id': key.split('-')[2],
                'interfaces_down': downcount,
                'interfaces_up': upcount,
            }})
        print(json.dumps(mainDict, indent=4))
        opticDict = {}
        for key, value in swDict.items():
            upcount = 0
            downcount = 0
            for k, v in value['intfs'].items():
                if not v['optic'] == 'sfp-missing':
                    if not opticDict.get(v['optic']):
                        opticDict[v['optic']] = []
                    opticDict[v['optic']].append(1)
        opticDict = dict(sorted(opticDict.items()))
        optics = {}
        for k, v in opticDict.items():
            optics[k] = {'quantity':len(v)}
        print(json.dumps(optics, indent=4))
    elif not args.node_input == None and not args.top_system_input == None:
        #file = open(args.node_input, 'r')
        topData = json.load(open(args.top_system_input, 'r'))
        nodeData = json.load(open(args.node_input, 'r'))
        print(json.dumps(nodeData, indent=4))
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
                swDict[a.dn].update(deepcopy({'model':a.model}))
        swDict = dict(sorted(swDict.items()))
        for k, v in swDict.items():
            v['intfs'] = dict(sorted(v['intfs'].items()))
        #print(json.dumps(swDict, indent=4))
        
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

if __name__ == '__main__':
    main()
