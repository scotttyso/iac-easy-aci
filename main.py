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
from classes import access, admin, fabric, site_policies, system_settings
from easy_functions import countKeys, findKeys, findVars
from easy_functions import read_easy_jsonData, read_in
from easy_functions import stdout_log
from pathlib import Path
import argparse
import json
import os
import platform
import re
import requests
import stdiomask
import subprocess
import time

#=====================================================================
# Note: This is simply to make it so the classes don't appear Unused.
#=====================================================================
class_list = [access, admin, fabric, site_policies, system_settings, tenants]

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
access_regex = re.compile('^(aep_profile|bpdu|cdp|(fibre|port)_(channel|security)|l2_interface|l3_domain|(leaf|spine)_pg|link_level|lldp|mcp|pg_(access|breakout|bundle|spine)|phys_dom|stp|vlan_pool)$')
admin_regex = re.compile('^(auth|export_policy|radius|remote_host|security|tacacs)$')
admin_regex = re.compile('^(auth|export_policy|remote_host|security)$')
system_settings_regex = re.compile('^(apic_preference|bgp_(asn|rr)|global_aes)$')
bridge_domains_regex = re.compile('^add_bd$')
contracts_regex = re.compile('(^(contract|filter|subject)_(add|entry|to_epg)$)')
epgs_regex = re.compile('^((app|epg)_add)$')
fabric_regex = '^(date_time|dns_profile|ntp(_key)?|smart_(callhome|destinations|smtp_server)|snmp_(clgrp|community|destinations|policy|user)|syslog(_destinations)?)$'
inventory_regex = re.compile('^(apic_inb|switch|vpc_pair)$')
l3out_regex = re.compile('^(add_l3out|ext_epg|node_(prof|intf|path)|bgp_peer)$')
mgmt_tenant_regex = re.compile('^(add_bd|mgmt_epg|oob_ext_epg)$')
sites_regex = re.compile('^(site_id|group_id)$')
tenants_regex = re.compile('^((tenant|vrf)_add|vrf_community)$')

#======================================================
# Function to run 'terraform plan' and
# 'terraform apply' in the each folder of the
# Destination Directory.
#======================================================
def apply_aci_terraform(folders):

    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'  Found the Followng Folders with uncommitted changes:\n')
    for folder in folders:
        print(f'  - {folder}')
    print(f'\n  Beginning Terraform Plan and Apply in each folder.')
    print(f'\n-----------------------------------------------------------------------------\n')

    time.sleep(7)

    response_p = ''
    response_a = ''
    for folder in folders:
        path = './%s' % (folder)
        lock_count = 0
        p = subprocess.Popen(['terraform', 'init', '-plugin-dir=../../../terraform-plugins/providers/'],
                             cwd=path,
                             stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT)
        for line in iter(p.stdout.readline, b''):
            print(line)
            if re.search('does not match configured version', line.decode("utf-8")):
                lock_count =+ 1

        if lock_count > 0:
            p = subprocess.Popen(['terraform', 'init', '-upgrade', '-plugin-dir=../../../terraform-plugins/providers/'], cwd=path)
            p.wait()
        p = subprocess.Popen(['terraform', 'plan', '-out=main.plan'], cwd=path)
        p.wait()
        while True:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'  Terraform Plan Complete.  Please Review the Plan and confirm if you want')
            print(f'  to move forward.  "A" to Apply the Plan. "S" to Skip.  "Q" to Quit.')
            print(f'  Current Working Directory: {folder}')
            print(f'\n-----------------------------------------------------------------------------\n')
            response_p = input('  Please Enter ["A", "S" or "Q"]: ')
            if re.search('^(A|S)$', response_p):
                break
            elif response_p == 'Q':
                exit()
            else:
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'  A Valid Response is either "A", "S" or "Q"...')
                print(f'\n-----------------------------------------------------------------------------\n')

        if response_p == 'A':
            p = subprocess.Popen(['terraform', 'apply', '-parallelism=1', 'main.plan'], cwd=path)
            p.wait()

        while True:
            if response_p == 'A':
                response_p = ''
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'  Terraform Apply Complete.  Please Review for any errors and confirm if you')
                print(f'  want to move forward.  "M" to Move to the Next Section. "Q" to Quit..')
                print(f'\n-----------------------------------------------------------------------------\n')
                response_a = input('  Please Enter ["M" or "Q"]: ')
            elif response_p == 'S':
                break
            if response_a == 'M':
                break
            elif response_a == 'Q':
                exit()
            else:
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'  A Valid Response is either "M" or "Q"...')
                print(f'\n-----------------------------------------------------------------------------\n')

#======================================================
# Function to Check the Git Status of the Folders
#======================================================
def check_git_status():
    random_folders = []
    git_path = './'
    result = subprocess.Popen(['python3', '-m', 'git_status_checker'], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    while(True):
        # returns None while subprocess is running
        retcode = result.poll()
        line = result.stdout.readline()
        line = line.decode('utf-8')
        if re.search(r'M (.*/).*.tf\n', line):
            folder = re.search(r'M (.*/).*.tf\n', line).group(1)
            if not re.search(r'ACI.templates', folder):
                if not folder in random_folders:
                    random_folders.append(folder)
        elif re.search(r'\?\? (.*/).*.tf\n', line):
            folder = re.search(r'\?\? (.*/).*.tf\n', line).group(1)
            if not re.search(r'ACI.templates', folder):
                if not folder in random_folders:
                    random_folders.append(folder)
        elif re.search(r'\?\? (ACI/.*/)\n', line):
            folder = re.search(r'\?\? (ACI/.*/)\n', line).group(1)
            if not (re.search(r'ACI.templates', folder) or re.search(r'\.terraform', folder)):
                if os.path.isdir(folder):
                    folder = [folder]
                    random_folders = random_folders + folder
                else:
                    group_x = [os.path.join(folder, o) for o in os.listdir(folder) if os.path.isdir(os.path.join(folder,o))]
                    random_folders = random_folders + group_x
        if retcode is not None:
            break

    if not random_folders:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   There were no uncommitted changes in the environment.')
        print(f'   Proceedures Complete!!! Closing Environment and Exiting Script.')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

    strict_folders = []
    folder_order = ['Access', 'System', 'Tenant_common', 'Tenant_infra', 'Tenant_mgmt', 'Fabric', 'Admin', 'VLANs', 'Tenant_infra',]
    for folder in folder_order:
        for fx in random_folders:
            if folder in fx:
                if 'ACI' in folder:
                    strict_folders.append(fx)
    for folder in strict_folders:
        if folder in random_folders:
            random_folders.remove(folder)
    for folder in random_folders:
        if 'ACI' in folder:
            strict_folders.append(folder)

    # print(strict_folders)
    return strict_folders

#======================================================
# Function to Get User Password
#======================================================
def get_user_pass():
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Beginning Proceedures to Apply Terraform Resources to the environment')
    print(f'\n-----------------------------------------------------------------------------\n')

    user = input('Enter APIC username: ')
    while True:
        try:
            password = stdiomask.getpass(prompt='Enter APIC password: ')
            break
        except Exception as e:
            print('Something went wrong. Error received: {}'.format(e))

    os.environ['TF_VAR_aciUser'] = '%s' % (user)
    os.environ['TF_VAR_aciPass'] = '%s' % (password)

def merge_easy_aci_repository(easy_jsonData):
    jsonData = easy_jsonData['components']['schemas']['easy_aci']['allOf'][1]['properties']
    
    # Obtain Operating System and Get TF_DEST_DIR variable from Environment
    opSystem = platform.system()
    if os.environ.get('TF_DEST_DIR') is None:
        tfDir = 'Intersight'
    else:
        tfDir = os.environ.get('TF_DEST_DIR')

    # Get All sub-folders from the TF_DEST_DIR
    folders = []
    for root, dirs, files in os.walk(tfDir):
        for name in dirs:
            # print(os.path.join(root, name))
            folders.append(os.path.join(root, name))
    folders.sort()

    # Remove the First Level Folders from the List
    for folder in folders:
        print(f'folder is {folder}')
        if '/' in folder:
            x = folder.split('/')
            if len(x) == 2:
                folders.remove(folder)

    # Get the Latest Release Tag for the terraform-intersight-imm repository
    url = f'https://github.com/terraform-cisco-modules/terraform-easy-aci/tags/'
    r = requests.get(url, stream=True)
    repoVer = 'BLANK'
    stringMatch = False
    while stringMatch == False:
        for line in r.iter_lines():
            toString = line.decode("utf-8")
            if re.search('/releases/tag/(\d+\.\d+\.\d+)', toString):
                repoVer = re.search('/releases/tag/(\d+\.\d+\.\d+)', toString).group(1)
                break
        stringMatch = True

    for folder in folders:
        folderVer = "0.0.0"
        # Get the version.txt file if exist to compare to the latest Git Release of the repository
        if opSystem == 'Windows':
            if os.path.isfile(f'{folder}\\version.txt'):
                with open(f'{folder}\\version.txt') as f:
                    folderVer = f.readline().rstrip()
        else:
            if os.path.isfile(f'{folder}/version.txt'):
                with open(f'{folder}/version.txt') as f:
                    folderVer = f.readline().rstrip()
        
        # Determine the Type of Folder. i.e. Is this for Access Policies
        if os.path.isdir(folder):
            if opSystem == 'Windows':
                folder_length = len(folder.split('\\'))
                folder_type = folder.split('\\')[folder_length -1]
            else:
                folder_length = len(folder.split('/'))
                folder_type = folder.split('/')[folder_length -1]
            if re.search('^tenant_', folder_type):
                folder_type = 'tenant'
            
            # Get List of Files to download from jsonData
            files = jsonData['files'][folder_type]
            print(f'\n-------------------------------------------------------------------------------------------\n')
            print(f'\n  Beginning Easy ACI Module Downloads for "{folder}"\n')

            # Run the process to check for the files in the folder and if it doesn't exist download from Github
            for file in files:
                git_url = 'https://raw.github.com/terraform-cisco-modules/terraform-easy-aci/master/modules'
                if opSystem == 'Windows':
                    dest_file = f'{folder}\\{file}'
                else:
                    dest_file = f'{folder}/{file}'
                if not os.path.isfile(dest_file):
                    print(f'  Downloading "{file}"')
                    url = f'{git_url}/{folder_type}/{file}'
                    r = requests.get(url)
                    open(dest_file, 'wb').write(r.content)
                    print(f'  "{file}" Download Complete!\n')
                else:
                    if opSystem == 'Windows':
                        if not os.path.isfile(f'{folder}\\version.txt'):
                            print(f'  Downloading "{file}"')
                            url = f'{git_url}/{folder_type}/{file}'
                            r = requests.get(url)
                            open(dest_file, 'wb').write(r.content)
                            print(f'  "{file}" Download Complete!\n')
                        elif not os.path.isfile(f'{folder}\\version.txt'):
                            print(f'  Downloading "{file}"')
                            url = f'{git_url}/{folder_type}/{file}'
                            r = requests.get(url)
                            open(dest_file, 'wb').write(r.content)
                            print(f'  "{file}" Download Complete!\n')
                        elif os.path.isfile(f'{folder}\\version.txt'):
                            if not folderVer == repoVer:
                                print(f'  Downloading "{file}"')
                                url = f'{git_url}/{folder_type}/{file}'
                                r = requests.get(url)
                                open(dest_file, 'wb').write(r.content)
                                print(f'  "{file}" Download Complete!\n')
                    else:
                        if not os.path.isfile(f'{folder}/version.txt'):
                            print(f'  Downloading "{file}"')
                            url = f'{git_url}/{folder_type}/{file}'
                            r = requests.get(url)
                            open(dest_file, 'wb').write(r.content)
                            print(f'  "{file}" Download Complete!\n')
                        elif not os.path.isfile(f'{folder}/version.txt'):
                            print(f'  Downloading "{file}"')
                            url = f'{git_url}/{folder_type}/{file}'
                            r = requests.get(url)
                            open(dest_file, 'wb').write(r.content)
                            print(f'  "{file}" Download Complete!\n')
                        elif os.path.isfile(f'{folder}/version.txt'):
                            if not folderVer == repoVer:
                                print(f'  Downloading "{file}"')
                                url = f'{git_url}/{folder_type}/{file}'
                                r = requests.get(url)
                                open(dest_file, 'wb').write(r.content)
                                print(f'  "{file}" Download Complete!\n')

            # Create the version.txt file to prevent redundant downloads for the same Github release
            if not os.path.isfile(f'{folder}/version.txt'):
                print(f'* Creating the repo "terraform-easy-aci" version check file\n "{folder}/version.txt"')
                open(f'{folder}/version.txt', 'w').write('%s\n' % (repoVer))
            elif not folderVer == repoVer:
                print(f'* Updating the repo "terraform-easy-aci" version check file\n "{folder}/version.txt"')
                open(f'{folder}/version.txt', 'w').write('%s\n' % (repoVer))

            print(f'\n  Completed Easy IMM Module Downloads for "{folder}"')
            print(f'\n-------------------------------------------------------------------------------------------\n')

    # Loop over the folder list again and create blank auto.tfvars files for anything that doesn't already exist
    for folder in folders:
        if os.path.isdir(folder):
            if opSystem == 'Windows':
                folder_length = len(folder.split('\\'))
                folder_type = folder.split('\\')[folder_length -1]
            else:
                folder_length = len(folder.split('/'))
                folder_type = folder.split('/')[folder_length -1]
            if re.search('^tenant_', folder_type):
                folder_type = 'tenant'
            files = jsonData['files'][folder_type]
            removeList = jsonData['remove_files']
            for xRemove in removeList:
                if xRemove in files:
                    files.remove(xRemove)
            for file in files:
                varFiles = f"{file.split('.')[0]}.auto.tfvars"
                if opSystem == 'Windows':
                    dest_file = f'{folder}\\{varFiles}'
                else:    
                    dest_file = f'{folder}/{varFiles}'
                if not os.path.isfile(dest_file):
                    wr_file = open(dest_file, 'w')
                    x = file.split('.')
                    x2 = x[0].split('_')
                    varList = []
                    for var in x2:
                        var = var.capitalize()
                        varList.append(var)
                    varDescr = ' '.join(varList)
                    varDescr = varDescr + '- Variables'

                    wrString = f'#______________________________________________\n#\n# {varDescr}\n'\
                        '#______________________________________________\n'\
                        '\n%s = {\n}\n' % (file.split('.')[0])
                    wr_file.write(wrString)
                    wr_file.close()

            # Run terraform fmt to cleanup the formating for all of the auto.tfvar files and tf files if needed
            print(f'\n-------------------------------------------------------------------------------------------\n')
            print(f'  Running "terraform fmt" in folder "{folder}",')
            print(f'  to correct variable formatting!')
            print(f'\n-------------------------------------------------------------------------------------------\n')
            p = subprocess.Popen(
                ['terraform', 'fmt', folder],
                stdout = subprocess.PIPE,
                stderr = subprocess.PIPE
            )
            print('Format updated for the following Files:')
            for line in iter(p.stdout.readline, b''):
                line = line.decode("utf-8")
                line = line.strip()
                print(f'- {line}')

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
# Function to Read the Fabric Worksheet
#======================================================
def process_inventory(easyDict, easy_jsonData, wb):
    # Evaluate Inventory Worksheet
    class_init = 'access'
    class_folder = 'access'
    func_regex = inventory_regex
    ws = wb['Inventory']
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
            3. bridge_domains: for Bridge_Domains\
            4. contracts: for Contracts\
            5. epgs: for EPGs\
            6. fabric: for Fabric\
            7. inventory: for Inventory\
            8. l3out: for L3Out\
            9. sites: for Sites\
            10. system_settings: for System_Settings\
            11. tenants: for Tenants'
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
    easyDict = {
        'access':{
            'domains':[],
            'firmware':[],
            'global_policies':[],
            'interface_policies':[],
            'inventory':[],
            'leaf_interface_policy_groups':[],
            'leaf_policy_groups':[],
            'leaf_profiles':[],
            'spine_interface_policy_groups':[],
            'spine_policy_groups':[],
            'spine_profiles':[],
            'vmm':[],
            'vpc_domains':[],
        },
        'admin':{
            'authentication':[],
            'configuration_backups':[],
            'global_security':[],
            'radius':[],
            'tacacs':[],
        },
        'fabric':{
            'date_and_time':[],
            'dns_profiles':[],
            'smart_callhome':[],
            'snmp_policies':[],
            'syslog':[],
        },
        'inventory':{},
        'sites':{},
        'system_settings':{
            'apic_connectivity_preference':[],
            'bgp_asn':[],
            'bgp_rr':[],
            'global_aes_encryption_settings':[]
        },
        'tenants':{
            'application_epgs':[],
            'application_profiles':[],
            'bfd_interface_policies':[],
            'bgp_policies':[],
            'bridge_domains':[],
            'dhcp_option_policies':[],
            'dhcp_relay_policies':[],
            'endpoint_retention_policies':[],
            'filters':[],
            'hsrp_policies':[],
            'l3out_floating_svi':[],
            'l3out_hsrp':[],
            'l3out_static_route':[],
            'l3outs':[],
            'ospf_policies':[],
            'route_map_match_rules':[],
            'route_map_set_rules':[],
            'route_maps_for_route_control':[],
            'schemas':[],
            'tenants':[],
            'vrfs':[],
        },
        'wb':wb
    }

    # Run Proceedures for Worksheets in the Workbook
    easyDict = process_sites(easyDict, easy_jsonData, wb)

    # Either Run All Remaining Proceedures or Just Specific based on sys.argv[2:]
    if not args.worksheet == None:
        ws_regex = '^(access|admin|bridge_domains|contracts|epgs|fabric|inventory|l3out|sites|system_settings|tenants)$'
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
        easyDict = process_system_settings(easyDict, easy_jsonData, wb)
        easyDict = process_fabric(easyDict, easy_jsonData, wb)
        easyDict = process_admin(easyDict, easy_jsonData, wb)
        # easyDict = process_tenants(easyDict, easy_jsonData, wb)
        # easyDict = process_epgs(easyDict, easy_jsonData, wb)
        # easyDict = process_bridge_domains(easyDict, easy_jsonData, wb)
        # easyDict = process_access(easyDict, easy_jsonData, wb)
        # easyDict = process_inventory(easyDict, easy_jsonData, wb)
        # easyDict = process_l3out(easyDict, easy_jsonData, wb)
        # easyDict = process_contracts(easyDict, easy_jsonData, wb)

    # Begin Proceedures to Create files
    read_easy_jsonData(easy_jsonData, **easyDict)
    merge_easy_aci_repository(easy_jsonData)
    easyDict.pop('wb')
    # print(json.dumps(easyDict, indent = 4))

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
