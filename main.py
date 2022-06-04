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
from classes import access, admin, fabric, site_policies, switches, system_settings, tenants
from easy_functions import apply_terraform
from easy_functions import git_base_repo, git_check_status
from easy_functions import countKeys, findKeys, findVars
from easy_functions import get_latest_versions, merge_easy_aci_repository
from easy_functions import read_easy_jsonData, read_in
from easy_functions import stdout_log
import argparse
import json
import os
import platform
import re

#=====================================================================
# Note: This is simply to make it so the classes don't appear Unused.
#=====================================================================
class_list = [access, admin, fabric, site_policies, switches, system_settings, tenants]

#======================================================
# Regular Expressions to Control wich rows in the
# Worksheet should be processed.
#======================================================
a1 = '(domains_(l3|phys)|global_aaep|interface_policy|pg_(access|breakout|bundle|spine)'
a2 = '(leaf|spine)_pg|pol_(cdp|fc|l2|link_level|lldp|mcp|port_(ch|sec)|stp)|pools_vlan)'
access_regex = f'^({a1}|{a2})$'

admin_regex = '^(auth|(export|mg)_policy|maint_group|radius|remote_host|security|tacacs)$'
apps_epgs_regex = '^((app|epg|vmm)_(add|(vmm_)?policy))$'
bds_regex = '^((bd)_(add|general|l3|subnet))$'
contracts_regex = '(^(assign_contract|(contract|filter|subject)_(add|entry))$)'

f1 = 'date_time|dns_profile|ntp(_key)?|smart_(callhome|destinations|smtp_server)'
f2 = 'snmp_(clgrp|community|destinations|policy|user)|syslog(_destinations)?'
fabric_regex = f'^({f1}|{f2})$'

l3out1 = '((bgp|eigrp|ospf)_(peer|policy|profile|routing)|ext_epg(_policy|_sub)?)'
l3out2 = 'l3out_(add|policy)|node_(interface|intf_(cfg|policy)|profile)?'
l3out_regex = f'^({l3out1}|{l3out2})$'

port_convert_regex = '^port_cnvt$'
sites_regex = '^(site_id|group_id)$'
switch_regex = '^(sw_modules|switch)$'
system_settings_regex = '^(apic_preference|bgp_(asn|rr)|global_aes)$'
tenants_regex = '^(tenant_(add|site)|vrf_(add|community|policy))$'
tenant_pol_regex = '^(apic_inb|bgp_pfx|(eigrp|ospf)_interface)$'
virtual_regex = '^(vmm_(controllers|creds|domain|elagp|vswitch))$'

#=================================================================
# Function to Read the Access Worksheet
#=================================================================
def process_access(args, easyDict, easy_jsonData, wb):
    # Evaluate Access Worksheet
    class_init = 'access'
    class_folder = 'access'
    func_regex = access_regex
    ws = wb['Access']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Function to process the Admin Worksheet
#=================================================================
def process_admin(args, easyDict, easy_jsonData, wb):
    # Evaluate Admin Worksheet
    class_init = 'admin'
    class_folder = 'admin'
    func_regex = admin_regex
    ws = wb['Admin']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Function to process the Fabric Worksheet
#=================================================================
def process_fabric(args, easyDict, easy_jsonData, wb):
    # Evaluate Fabric Worksheet
    class_init = 'fabric'
    class_folder = 'fabric'
    func_regex = fabric_regex
    ws = wb['Fabric']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Process the Port Conversions from the Switch Profiles worksheet.
#=================================================================
def process_port_convert(args, easyDict, easy_jsonData, wb):
    # Evaluate Inventory Worksheet
    class_init = 'switches'
    class_folder = 'switches'
    func_regex = port_convert_regex
    ws = wb['Switch Profiles']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Function to Setup Terraform Run Location
#=================================================================
def process_site_settings(args, easyDict, easy_jsonData, wb):
    kwargs = {
        'args':args,
        'easyDict':easyDict,
        'easy_jsonData':easy_jsonData,
        'row_num':0,
        'wb':wb,
        'ws': wb['Sites']
    }
    easyDict = site_policies('site_settings').site_settings(**kwargs)
    return easyDict

#=================================================================
# Function to process the Sites Worksheet
#=================================================================
def process_sites(args, easyDict, easy_jsonData, wb):
    # Evaluate Sites Worksheet
    class_init = 'site_policies'
    class_folder = 'sites'
    func_regex = sites_regex
    ws = wb['Sites']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Function to process the Fabric Worksheet
#=================================================================
def process_switches(args, easyDict, easy_jsonData, wb):
    # Evaluate Switches Worksheet
    class_init = 'switches'
    class_folder = 'switches'
    func_regex = switch_regex
    ws = wb['Switch Profiles']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Function to process the System Settings Worksheet
#=================================================================
def process_system_settings(args, easyDict, easy_jsonData, wb):
    # Evaluate System_Settings Worksheet
    class_init = 'system_settings'
    class_folder = 'system_settings'
    func_regex = system_settings_regex
    ws = wb['System Settings']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Function to process the Tenants Worksheet
#=================================================================
def process_tenants(args, easyDict, easy_jsonData, wb):
    class_init = 'tenants'
    class_folder = 'tenants'

    # Evaluate the Tenants Worksheet
    func_regex = tenants_regex
    ws = wb['Tenants']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the Tenant Policies Worksheet
    func_regex = tenant_pol_regex
    ws = wb['Tenant Policies']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the Bridge Domains Worksheet
    func_regex = bds_regex
    ws = wb['Bridge Domains']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the Apps and EPGs Worksheet
    func_regex = apps_epgs_regex
    ws = wb['Apps and EPGs']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # # Evaluate the L3Out Worksheet
    func_regex = l3out_regex
    ws = wb['L3Out']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # # Evaluate the Contracts Worksheet
    func_regex = contracts_regex
    ws = wb['Contracts']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    return easyDict

#=================================================================
# Function to process the Virtual Networking Worksheet
#=================================================================
def process_virtual_networking(args, easyDict, easy_jsonData, wb):
    # Evaluate Tenants Worksheet
    class_init = 'access'
    class_folder = 'access'
    func_regex = virtual_regex
    ws = wb['Virtual Networking']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    return easyDict

#=================================================================
# Function to process the Worksheets and Create Terraform Files
#=================================================================
def read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws):
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
                    'args':args,
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

#=================================================================
# The Main Module
#=================================================================
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
            8. port_convert: for Uplink to Download Conversion\
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

    opSystem = platform.system()
    if opSystem == 'Windows': path_sep = '\\'
    else: path_sep = '/'
    destdirCheck = False
    while destdirCheck == False:
        splitDir = args.dir.split(path_sep)
        splitDir = [i for i in splitDir if i]
        for folder in splitDir:
            if not re.search(r'^[\w\-\.\:\/\\]+$', folder):
                print(folder)
                print(f'\n-------------------------------------------------------------------------------------------\n')
                print(f'  !!ERROR!!')
                print(f'  The Directory structure can only contain the following characters:')
                print(f'  letters(a-z, A-Z), numbers(0-9), hyphen(-), period(.), colon(:), or and underscore(_).')
                print(f'  It can be a short path or a fully qualified path. {folder} failed this check.')
                print(f'\n-------------------------------------------------------------------------------------------\n')
                exit()
        destdirCheck = True

    # Set the Source Workbook
    if os.path.isfile(args.workbook):
        excel_workbook = args.workbook
    else:
        print(f'\n-------------------------------------------------------------------------------------------\n')
        print( '\nWorkbook not Found.  Please enter a valid /path/filename for the source workbook.')
        print(f'\n-------------------------------------------------------------------------------------------\n')
        while True:
            print('Please enter a valid /path/filename for the source you will be using.')
            excel_workbook = input('/Path/Filename: ')
            if os.path.isfile(excel_workbook):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   {excel_workbook} exists.  Will Now begin collecting variables...')
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

    # Initialize the Base Repo/Terraform Working Directory
    if not os.path.isdir(args.dir):
        os.mkdir(args.dir)
    # baseRepo = git_base_repo(args, wb)

    # Process the Sites Worksheet
    easyDict['wb'] = wb
    easyDict = process_sites(args, easyDict, easy_jsonData, wb)

    # Process Individual Worksheets if specified in args or Process All by Default
    if not args.worksheet == None:
        r1 = 'access|admin|bridge_domains|contracts|epgs|fabric|inventory'
        r2 = 'l3out|port_convert|sites|switch|system_settings|tenants'
        ws_regex = f'^({r1}|{r2})$'
        if re.search(ws_regex, str(args.worksheet)):
            process_type = f'process_{args.worksheet}'
            eval(f"{process_type}(args, easyDict, easy_jsonData, wb)")
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
            easyDict = eval(f"{process_type}(args, easyDict, easy_jsonData, wb)")


    # Begin Proceedures to Create files
    read_easy_jsonData(args, easy_jsonData, **easyDict)
    easyDict = process_site_settings(args, easyDict, easy_jsonData, wb)
    merge_easy_aci_repository(args, easy_jsonData, **easyDict)
    exit()
    uncommitted_folders = git_check_status(args)
    apply_terraform(args, uncommitted_folders)

    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'  Proceedures Complete!!! Closing Environment and Exiting Script.')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

if __name__ == '__main__':
    main()
