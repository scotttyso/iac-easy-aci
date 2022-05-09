#!/usr/bin/env python3
"""ACI/NDO IAC - 
This Script is to Create Terraform HCL configuration from an Excel Spreadsheet.
It uses argparse to take in the following CLI arguments:
    d or dir:           Base Directory to use for creation of the HCL Configuration Files
    w or workbook:   Name of Excel Workbook file for the Data Source
"""

#======================================================
# Source Modules
#======================================================
from class_tenants import tenants
from classes import access, admin, fabric, site_policies, switches, system_settings
from easy_functions import apply_aci_terraform, check_git_status
from easy_functions import countKeys, findKeys, findVars, get_user_pass
from easy_functions import get_latest_versions, merge_easy_aci_repository
from easy_functions import read_easy_jsonData, read_in
from easy_functions import stdout_log
from pathlib import Path
import argparse
import json
import os
import re

#=====================================================================
# Note: This is simply to make it so the classes don't appear Unused.
#=====================================================================
class_list = [access, admin, fabric, site_policies, switches, system_settings, tenants]

#======================================================
# Global Variables
#======================================================
excel_workbook = None
home = Path.home()
workspace_dict = {}

#======================================================
# Regular Expressions to Control wich rows in the
# Worksheet should be processed.
#======================================================
a1 = 'aep_profile|cdp|(fibre|port)_(channel|security)|interface_policy|l2_interface|(phys|l3)_domain|'
a2 = '(leaf|spine)_pg|link_level|lldp|mcp|pg_(access|breakout|bundle|spine)|stp|vlan_pool'
access_regex = f'^({a1}{a2})$'
admin_regex = '^(auth|(export|mg)_policy|maint_group|radius|remote_host|security|tacacs)$'
bridge_domains_regex = '^add_bd$'
contracts_regex = '(^(contract|filter|subject)_(add|entry|to_epg)$)'
epgs_regex = '^((app|epg)_add)$'

f1 = 'date_time|dns_profile|ntp(_key)?|smart_(callhome|destinations|smtp_server)|'
f2 = 'snmp_(clgrp|community|destinations|policy|user)|syslog(_destinations)?'
fabric_regex = f'^({f1}{f2})$'
l3out_regex = '^(add_l3out|ext_epg|node_(prof|intf|path)|bgp_peer)$'
mgmt_tenant_regex = '^(add_bd|mgmt_epg|oob_ext_epg)$'
sites_regex = '^(site_id|group_id)$'
switch_regex = '^(sw_modules|switch)$'
system_settings_regex = '^(apic_preference|bgp_(asn|rr)|global_aes)$'
tenants_regex = '^((tenant|vrf)_add|vrf_community)$'
tenants_regex = '^((tenant)_block)$'
virtual_regex = '^(vmm_(controllers|creds|domain|elagp|vswitch))$'

#======================================================
# Function to Read the Access Worksheet
#======================================================
def process_access(easyDict, easy_jsonData, wb):
    # Evaluate Access Worksheet
    class_init = 'access'
    class_folder = 'access'
    func_regex = access_regex
    ws = wb['Access']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Admin Worksheet
#======================================================
def process_admin(easyDict, easy_jsonData, wb):
    # Evaluate Admin Worksheet
    class_init = 'admin'
    class_folder = 'admin'
    func_regex = admin_regex
    ws = wb['Admin']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Bridge Domains Worksheet
#======================================================
def process_bridge_domains(easyDict, easy_jsonData, wb):
    # Evaluate Bridge_Domains Worksheet
    class_init = 'tenants'
    class_folder = 'tenants'
    func_regex = bridge_domains_regex
    ws = wb['Bridge_Domains']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Contracts Worksheet
#======================================================
def process_contracts(easyDict, easy_jsonData, wb):
    # Evaluate Contracts Worksheet
    class_init = 'tenants'
    class_folder = 'tenants'
    func_regex = contracts_regex
    ws = wb['Contracts']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the EPGs Worksheet
#======================================================
def process_epgs(easyDict, easy_jsonData, wb):
    # Evaluate EPGs Worksheet
    class_init = 'tenants'
    class_folder = 'tenants'
    func_regex = epgs_regex
    ws = wb['EPGs']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Fabric Worksheet
#======================================================
def process_fabric(easyDict, easy_jsonData, wb):
    # Evaluate Fabric Worksheet
    class_init = 'fabric'
    class_folder = 'fabric'
    func_regex = fabric_regex
    ws = wb['Fabric']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the L3Out Worksheet
#======================================================
def process_l3out(easyDict, easy_jsonData, wb):
    # Evaluate L3Out Worksheet
    class_init = 'tenants'
    class_folder = 'tenants'
    func_regex = l3out_regex
    ws = wb['L3Out']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Sites Worksheet
#======================================================
def process_sites(easyDict, easy_jsonData, wb):
    # Evaluate Sites Worksheet
    class_init = 'site_policies'
    class_folder = 'sites'
    func_regex = sites_regex
    ws = wb['Sites']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Fabric Worksheet
#======================================================
def process_switches(easyDict, easy_jsonData, wb):
    # Evaluate Inventory Worksheet
    class_init = 'switches'
    class_folder = 'switches'
    func_regex = switch_regex
    ws = wb['Switch Profiles']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the System Settings Worksheet
#======================================================
def process_system_settings(easyDict, easy_jsonData, wb):
    # Evaluate System_Settings Worksheet
    class_init = 'system_settings'
    class_folder = 'system_settings'
    func_regex = system_settings_regex
    ws = wb['System Settings']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Tenants Worksheet
#======================================================
def process_tenants(easyDict, easy_jsonData, wb):
    # Evaluate Tenants Worksheet
    class_init = 'tenants'
    class_folder = 'tenants'
    func_regex = tenants_regex
    ws = wb['Tenants']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#======================================================
# Function to Read the Virtual Networking Worksheet
#======================================================
def process_virtual_networking(easyDict, easy_jsonData, wb):
    # Evaluate Tenants Worksheet
    class_init = 'access'
    class_folder = 'access'
    func_regex = virtual_regex
    ws = wb['Virtual Networking']
    easyDict = read_worksheet(class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

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
# The Main Module
#======================================================
def main():
    Parser = argparse.ArgumentParser(description='IaC Easy ACI Deployment Module')
    Parser.add_argument('-d', '--dir',
        default = 'ACI',
        help = 'The Directory to use for the Creation of the Terraform Files.'
    )
    Parser.add_argument('-wb', '--workbook',
        default = 'ACI_Base_Workbookv2.xlsx',
        help = 'The source Workbook.'
    )
    Parser.add_argument('-ws', '--worksheet', 
        default = None,
        help = 'Only evaluate this single worksheet. Worksheet values are:\
            1. access - for Access\
            2. admin: for Admin\
            3. bridge_domains: for Bridge Domains\
            4. contracts: for Contracts\
            5. epgs: for EPGs\
            6. fabric: for Fabric\
            7. l3out: for L3Out\
            8. sites: for Sites\
            9. switches: for Switch Profiles\
            10. system_settings: for System Settings\
            11. tenants: for Tenants\
            12. virtual_networking: for Virtual Networking'
    )
    args = Parser.parse_args()

    jsonFile = 'templates/variables/easy_variables.json'
    jsonOpen = open(jsonFile, 'r')
    easy_jsonData = json.load(jsonOpen)
    jsonOpen.close()

    destdirCheck = False
    while destdirCheck == False:
        splitDir = args.dir.split("/")
        for folder in splitDir:
            if folder == '':
                folderCount = 0
            elif not re.search(r'^[\w\-\.\:\/\\]+$', folder):
                print(folder)
                print(f'\n-------------------------------------------------------------------------------------------\n')
                print(f'  !!ERROR!!')
                print(f'  The Directory structure can only contain the following characters:')
                print(f'  letters(a-z, A-Z), numbers(0-9), hyphen(-), period(.), colon(:), or and underscore(-).')
                print(f'  It can be a short path or a fully qualified path.')
                print(f'\n-------------------------------------------------------------------------------------------\n')
                exit()
        os.environ['TF_DEST_DIR'] = '%s' % (args.dir)
        destdirCheck = True

    # Ask user for required Information: ACI_DEPLOY_FILE
    if os.path.isfile(args.workbook):
        excel_workbook = args.workbook
    else:
        print('\nWorkbook not Found.  Please enter a valid /path/filename for the source workbook you will be using.')
        while True:
            print('Please enter a valid /path/filename for the source you will be using.')
            excel_workbook = input('/Path/Filename: ')
            if os.path.isfile(excel_workbook):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   {excel_workbook} exists.  Will Now Check for API Variables...')
                print(f'\n-----------------------------------------------------------------------------\n')
                break
            else:
                print('\nWorkbook not Found.  Please enter a valid /path/filename for the source you will be using.')

    # Load Workbook
    wb = read_in(excel_workbook)

    # Create Dictionary for Worksheets in the Workbook
    easyDict = easy_jsonData['components']['schemas']['easy_aci']['allOf'][1]['properties']['easyDict']

    # Obtain the Latest Provider Releases
    # easyDict = get_latest_versions(easyDict)
    ndoVersions = [
        "3.7.1g","3.7.1d","3.5.2f","3.5.1e","3.4.1i","3.4.1a","3.3.1e","3.2.1f","3.1.1l",
        "3.1.1i","3.1.1h","3.1.1g","3.0.3m","3.0.3l","3.0.3i","3.0.2k","3.0.2j"
    ]
    easyDict['latest_versions']['aci_provider_version'] = "2.1.0"
    easyDict['latest_versions']['ndo_provider_version'] = "0.6.0"
    easyDict['latest_versions']['ndo_versions']['enum'] = ndoVersions
    easyDict['latest_versions']['ndo_versions']['default'] = ndoVersions[0]
    easyDict['latest_versions']['terraform_version'] = "1.1.9"
    # print(json.dumps(easyDict, indent=4))
    # exit()

    # Run Proceedures for Worksheets in the Workbook
    easyDict['wb'] = wb
    easyDict = process_sites(easyDict, easy_jsonData, wb)

    # Either Run All Remaining Proceedures or Just Specific based on sys.argv[2:]
    if not args.worksheet == None:
        ws_regex = '^(access|admin|bridge_domains|contracts|epgs|fabric|inventory|l3out|sites|switch|system_settings|tenants)$'
        if re.search(ws_regex, str(args.worksheet)):
            process_type = f'process_{args.worksheet}'
            eval(f"{process_type}(easyDict, easy_jsonData, wb)")
        else:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   ERROR: "{args.worksheet}" is not a valid worksheet.  If you are trying ')
            print(f'   to run a single worksheet please re-enter the -ws argument.')
            print(f'   Exiting...')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()
    else:
        process_list = easy_jsonData['components']['schemas']['easy_aci']['allOf'][1]['properties']['processes']['enum']
        for x in process_list:
            process_type = f'process_{x}'
            eval(f"{process_type}(easyDict, easy_jsonData, wb)")

    easyDict.pop('wb')
    # print(json.dumps(easyDict['access']['virtual_networking'], indent = 4))
    # exit()

    # Begin Proceedures to Create files
    easyDict['wb'] = wb
    read_easy_jsonData(easy_jsonData, **easyDict)
    merge_easy_aci_repository(easy_jsonData)

    folders = check_git_status()
    get_user_pass()
    apply_aci_terraform(folders)
    # else:
    #     print('hello')
    #     path = './'
    #     repo = Repo.init(path)

    #     index = Repo.init(path.index)

    #     index.commit('Testing Commit')

    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'  Proceedures Complete!!! Closing Environment and Exiting Script.')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

if __name__ == '__main__':
    main()
