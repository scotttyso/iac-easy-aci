#!/usr/bin/env python3
"""ACI/NDO IaC - 
This Script is to Create Terraform HCL configuration from an Excel Spreadsheet.
It uses argparse to take in the following CLI arguments:
    d or dir:        Base Directory to use for creation of the HCL Configuration Files
    g or git-check:  By default the script will not use git to check the destination for git status.  Include this flag to perform git check.
    s or skip-version-check: Setting this to "True" will disable the login to the controllers to determine the running version.
    w or workbook:   Name of Excel Workbook file for the Data Source
"""

#======================================================
# Source Modules
#======================================================
import argparse
import classes
import easy_functions
import json
import os
import platform
import re
import sys

#======================================================
# Regular Expressions to Control wich rows in the
# Worksheet should be processed.
#======================================================
ac1 = '(l3|phys)_domains|global_aaep|interface_policy|pg_(access|breakout|bundle|spine|template)'
ac2 = '(leaf|spine)_pg|pol_(cdp|fc|l2|link_level|lldp|mcp|port_(ch|sec)|stp)|pre_built|vlan_pools'
access_regex = f'^({ac1}|{ac2})$'

ad1 = 'auth|(export|mg)_policy|maint_group|radius|recommended_settings|remote_host|security|tacacs'
ad2 = 'smart_(callhome|destinations|smtp_server)|syslog(_destinations)?'
admin_regex = f'^({ad1}|{ad2})$'

apps_epgs_regex = '^(app|epg)_(add|template|vmm_temp)$'
bds_regex = '^(bd|subnet)_(add|template)$'
contracts_regex = '^(contract|filter|subject)_(add|assign|entry|filters)$'

fa1 = 'date_time|dns_profile|ntp(_key)?|recommended_settings'
fa2 = 'snmp_(clgrp|community|destinations|policy|user)'
fabric_regex = f'^({fa1}|{fa2})$'

l31 = '(bgp|eigrp|ospf)_(peer|template|profile|routing)|ext_epg(_temp|_sub)?'
l32 = 'l3out_(add|template)|node_(interface|intf_(cfg|temp)|profile)?'
l3out_regex = f'^({l31}|{l32})$'

port_convert_regex = '^port_cnvt$'
sites_regex = '^(site_id|group_id)$'
switch_regex = '^(sw_modules|switch)$'
system_settings_regex = '^(apic_preference|bgp_(asn|rr)|recommended_settings)$'
tenants_regex = '^(ndo_schema|(template|tenant)_(add|site)|vrf_(add|community|template))$'
tenant_pol_regex = '^(apic_inb|bgp_pfx|dhcp_relay|(eigrp|ospf)_interface)$'
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
    easyDict['remove_default_args'] = True
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
        'remove_default_args':False,
        'row_num':0,
        'wb':wb,
        'ws': wb['Sites']
    }
    easyDict = classes.site_policies('site_settings').site_settings(**kwargs)
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
    easyDict['remove_default_args'] = False
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
    easyDict['remove_default_args'] = False
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)
    easyDict['remove_default_args'] = True
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
    easyDict['remove_default_args'] = True
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the Tenant Policies Worksheet
    func_regex = tenant_pol_regex
    ws = wb['Tenant Policies']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the Apps and EPGs Worksheet
    func_regex = apps_epgs_regex
    ws = wb['Apps and EPGs']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the Bridge Domains Worksheet
    func_regex = bds_regex
    ws = wb['Bridge Domains']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the L3Out Worksheet
    func_regex = l3out_regex
    ws = wb['L3Out']
    easyDict = read_worksheet(args, class_init, class_folder, easyDict, easy_jsonData, func_regex, wb, ws)

    # Evaluate the Contracts Worksheet
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
    func_list = easy_functions.findKeys(ws, func_regex)
    easy_functions.stdout_log(ws, None, 'begin')
    for func in func_list:
        count = easy_functions.countKeys(ws, func)
        var_dict = easy_functions.findVars(ws, func, rows, count)
        for pos in var_dict:
            row_num = var_dict[pos]['row']
            del var_dict[pos]['row']
            for x in list(var_dict[pos].keys()):
                if var_dict[pos][x] == '':
                    del var_dict[pos][x]
            easy_functions.stdout_log(ws, row_num, 'begin')
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
            easyDict = eval(f"classes.{class_init}(class_folder).{func}(**var_dict[pos])")
    
    easy_functions.stdout_log(ws, None, 'end')
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
    Parser.add_argument(
        '-g', '--git-check', action='store_true',
        help = 'By default the script will not use git to check the destination for git status.  Include this flag to perform git check.'
    )
    Parser.add_argument(
        '-s', '--skip-version-check', action='store_true',
        help = 'Flag to Skip the APIC and NDO Version Check.'
    )
    Parser.add_argument('-wb', '--workbook',
        default = 'ACI_Base_Workbookv3.xlsx',
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

    # Determine the Operating System
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
    wb = easy_functions.read_in(excel_workbook)

    # Create Dictionary for Worksheets in the Workbook
    easy_jsonData = easy_jsonData['components']['schemas']
    easyDict = {}
    easyDict['latest_versions'] = easy_jsonData['easy_aci']['allOf'][1]['properties']['latest_versions']

    # Obtain the Latest Provider Releases
    easyDict = easy_functions.get_latest_versions(easyDict)

    # Initialize the Base Repo/Terraform Working Directory
    if not os.path.isdir(args.dir):
        os.mkdir(args.dir)
    # baseRepo = easy_functions.git_base_repo(args, wb)

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
        process_list = easy_jsonData['easy_aci']['allOf'][1]['properties']['processes']['enum']
        for x in process_list:
            process_type = f'process_{x}'
            easyDict = eval(f"{process_type}(args, easyDict, easy_jsonData, wb)")

    # Begin Proceedures to Create files
    easy_functions.create_yaml(args, easy_jsonData, **easyDict)
    easyDict = process_site_settings(args, easyDict, easy_jsonData, wb)
    site_names, site_directories = easy_functions.merge_easy_aci_repository(args, easy_jsonData, **easyDict)
    changed_folders = []
    if args.git_check == False:
        changed_folders = easy_functions.git_check_status(args, site_names, site_directories)
    else:
        changed_folders = site_directories
    easyDict['changed_folders'] = changed_folders
    easyDict['site_names'] = site_names
    easy_functions.apply_terraform(args, path_sep, **easyDict)

    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'  Proceedures Complete!!! Closing Environment and Exiting Script.')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

if __name__ == '__main__':
    main()
