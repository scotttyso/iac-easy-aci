#!/usr/bin/env python3

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
from git import Repo
import getpass
import lib_aci
import os, re, sys
import subprocess
import time

# Global Variables
excel_workbook = None
home = Path.home()
workspace_dict = {}

Access_regex = re.compile('(^aep_profile|bpdu|cdp|(fibre|port)_(channel|security)|l2_interface|l3_domain|(leaf|spine)_pg|link_level|lldp|mcp|pg_(access|breakout|bundle|spine)|phys_dom|stp|vlan_pool$)')
Admin_regex = re.compile('(^export_policy|firmware|maint_group|radius|realm|remote_host|security|tacacs|tacacs_acct$)')
Best_Practices_regex = re.compile('(^bgp_(asn|rr)|ep_controls|error_recovery|fabric_settings|fabric_wide|isis_policy|mcp_policy$)')
Bridge_Domains_regex = re.compile('(^add_bd$)')
Contracts_regex = re.compile('(^(contract|filter|subject)_(add|entry|to_epg)$)')
DHCP_regex = re.compile('(^dhcp_add$)')
EPGs_regex = re.compile('(^add_epg$)')
Fabric_regex = re.compile('(^date_time|dns|dns_profile|domain|pod_policy|ntp|sch_dstgrp|sch_receiver|snmp_(client|clgrp|comm|policy|trap|user)|syslog_(dg|rmt)|trap_groups$)')
Inventory_regex = re.compile('(^apic_inb|switch|vpc_pair$)')
L3Out_regex = re.compile('(^add_l3out|ext_epg|node_(prof|intf|path)|bgp_peer$)')
Mgmt_Tenant_regex = re.compile('(^add_bd|mgmt_epg|oob_ext_epg$)')
Sites_regex = re.compile('(^site_id|group_id$)')
Tenant_regex = re.compile('(^add_tenant$)')
VRF_regex = re.compile('(^add_vrf|ctx_common$)')
VMM_regex = re.compile('(^add_vrf|ctx_common$)')

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

def get_user_pass():
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Beginning Proceedures to Apply Terraform Resources to the environment')
    print(f'\n-----------------------------------------------------------------------------\n')

    user = input('Enter APIC username: ')
    while True:
        try:
            password = getpass.getpass(prompt='Enter APIC password: ')
            break
        except Exception as e:
            print('Something went wrong. Error received: {}'.format(e))

    os.environ['TF_VAR_aciUser'] = '%s' % (user)
    os.environ['TF_VAR_aciPass'] = '%s' % (password)

def process_Access(wb):
    # Evaluate Access Worksheet
    lib_aci_ref = 'lib_aci.Access_Policies'
    func_regex = Access_regex
    ws = wb['Access']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Admin(wb):
    # Evaluate Admin Worksheet
    lib_aci_ref = 'lib_aci.Admin_Policies'
    func_regex = Admin_regex
    ws = wb['Admin']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Best_Practices(wb):
    # Evaluate Best_Practices Worksheet
    lib_aci_ref = 'lib_aci.Best_Practices'
    func_regex = Best_Practices_regex
    ws = wb['Best_Practices']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Bridge_Domains(wb):
    # Evaluate Bridge_Domains Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    func_regex = Bridge_Domains_regex
    ws = wb['Bridge_Domains']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Contracts(wb):
    # Evaluate Contracts Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    func_regex = Contracts_regex
    ws = wb['Contracts']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_DHCP_Relay(wb):
    # Evaluate DHCP Relay Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    func_regex = DHCP_regex
    ws = wb['DHCP Relay']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_EPGs(wb):
    # Evaluate EPGs Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    func_regex = EPGs_regex
    ws = wb['EPGs']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Fabric(wb):
    # Evaluate Fabric Worksheet
    lib_aci_ref = 'lib_aci.Fabric_Policies'
    func_regex = Fabric_regex
    ws = wb['Fabric']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Inventory(wb):
    # Evaluate Inventory Worksheet
    lib_aci_ref = 'lib_aci.Access_Policies'
    func_regex = Inventory_regex
    ws = wb['Inventory']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_L3Out(wb):
    # Evaluate L3Out Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    func_regex = L3Out_regex
    ws = wb['L3Out']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Mgmt_Tenant(wb):
    # Evaluate Mgmt_Tenant Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    ws = wb['Mgmt_Tenant']
    func_regex = Mgmt_Tenant_regex
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Sites(wb):
    # Evaluate Sites Worksheet
    lib_aci_ref = 'lib_aci.Site_Policies'
    func_regex = Sites_regex
    ws = wb['Sites']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_Tenants(wb):
    # Evaluate Tenants Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    func_regex = Tenant_regex
    ws = wb['Tenants']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_VRF(wb):
    # Evaluate VRF Worksheet
    lib_aci_ref = 'lib_aci.Tenant_Policies'
    func_regex = VRF_regex
    ws = wb['VRF']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def process_VMM(wb):
    # Evaluate Sites Worksheet
    lib_aci_ref = 'lib_aci.VMM_Policies'
    func_regex = VMM_regex
    ws = wb['VMM']
    read_worksheet(wb, ws, lib_aci_ref, func_regex)

def read_worksheet(wb, ws, lib_aci_ref, func_regex):
    rows = ws.max_row
    func_list = lib_aci.findKeys(ws, func_regex)
    class_init = '%s(ws)' % (lib_aci_ref)
    lib_aci.stdout_log(ws, None)
    for func in func_list:
        count = lib_aci.countKeys(ws, func)
        var_dict = lib_aci.findVars(ws, func, rows, count)
        for pos in var_dict:
            row_num = var_dict[pos]['row']
            del var_dict[pos]['row']
            for x in list(var_dict[pos].keys()):
                if var_dict[pos][x] == '':
                    del var_dict[pos][x]
            lib_aci.stdout_log(ws, row_num)
            eval("%s.%s(wb, ws, row_num, **var_dict[pos])" % (class_init, func))

def wb_update(wr_ws, status, i):
    # build green and red style sheets for excel
    bd1 = Side(style="thick", color="8EA9DB")
    bd2 = Side(style="medium", color="8EA9DB")
    wsh1 = NamedStyle(name="wsh1")
    wsh1.alignment = Alignment(horizontal="center", vertical="center", wrap_text="True")
    wsh1.border = Border(left=bd1, top=bd1, right=bd1, bottom=bd1)
    wsh1.font = Font(bold=True, size=15, color="FFFFFF")
    wsh2 = NamedStyle(name="wsh2")
    wsh2.alignment = Alignment(horizontal="center", vertical="center", wrap_text="True")
    wsh2.border = Border(left=bd2, top=bd2, right=bd2, bottom=bd2)
    wsh2.fill = PatternFill("solid", fgColor="305496")
    wsh2.font = Font(bold=True, size=15, color="FFFFFF")
    green_st = NamedStyle(name="ws_odd")
    green_st.alignment = Alignment(horizontal="center", vertical="center")
    green_st.border = Border(left=bd2, top=bd2, right=bd2, bottom=bd2)
    green_st.fill = PatternFill("solid", fgColor="D9E1F2")
    green_st.font = Font(bold=False, size=12, color="44546A")
    red_st = NamedStyle(name="ws_even")
    red_st.alignment = Alignment(horizontal="center", vertical="center")
    red_st.border = Border(left=bd2, top=bd2, right=bd2, bottom=bd2)
    red_st.font = Font(bold=False, size=12, color="44546A")
    yellow_st = NamedStyle(name="ws_even")
    yellow_st.alignment = Alignment(horizontal="center", vertical="center")
    yellow_st.border = Border(left=bd2, top=bd2, right=bd2, bottom=bd2)
    yellow_st.font = Font(bold=False, size=12, color="44546A")
    # green_st = xlwt.easyxf('pattern: pattern solid;')
    # green_st.pattern.pattern_fore_colour = 3
    # red_st = xlwt.easyxf('pattern: pattern solid;')
    # red_st.pattern.pattern_fore_colour = 2
    # yellow_st = xlwt.easyxf('pattern: pattern solid;')
    # yellow_st.pattern.pattern_fore_colour = 5
    # if stanzas to catch the status code from the request
    # and then input the appropriate information in the workbook
    # this then writes the changes to the doc
    if status == 200:
        wr_ws.write(i, 1, 'Success (200)', green_st)
    if status == 400:
        print("Error 400 - Bad Request - ABORT!")
        print("Probably have a bad URL or payload")
        wr_ws.write(i, 1, 'Bad Request (400)', red_st)
        pass
    if status == 401:
        print("Error 401 - Unauthorized - ABORT!")
        print("Probably have incorrect credentials")
        wr_ws.write(i, 1, 'Unauthorized (401)', red_st)
        pass
    if status == 403:
        print("Error 403 - Forbidden - ABORT!")
        print("Server refuses to handle your request")
        wr_ws.write(i, 1, 'Forbidden (403)', red_st)
        pass
    if status == 404:
        print("Error 404 - Not Found - ABORT!")
        print("Seems like you're trying to POST to a page that doesn't"
              " exist.")
        wr_ws.write(i, 1, 'Not Found (400)', red_st)
        pass
    if status == 666:
        print("Error - Something failed!")
        print("The POST failed, see stdout for the exception.")
        wr_ws.write(i, 1, 'Unkown Failure', yellow_st)
        pass
    if status == 667:
        print("Error - Invalid Input!")
        print("Invalid integer or other input.")
        wr_ws.write(i, 1, 'Unkown Failure', yellow_st)
        pass

def main():
    # Ask user for required Information: ACI_DEPLOY_FILE
    if sys.argv[1:]:
        if os.path.isfile(sys.argv[1]):
            excel_workbook = sys.argv[1]
        else:
            print('\nWorkbook not Found.  Please enter a valid /path/filename for the source you will be using.')
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
    else:
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
    wb = lib_aci.read_in(excel_workbook)

    # Run Proceedures for Worksheets in the Workbook
    process_Sites(wb)

    # Either Run All Remaining Proceedures or Just Specific based on sys.argv[2:]
    if sys.argv[2:]:
        if re.search('site', str(sys.argv[2:])):
            process_Sites(wb)
        elif re.search('access', str(sys.argv[2:])):
            process_Access(wb)
        elif re.search('inventory', str(sys.argv[2:])):
            process_Inventory(wb)
        elif re.search('admin', str(sys.argv[2:])):
            process_Admin(wb)
        elif re.search('best', str(sys.argv[2:])):
            process_Best_Practices(wb)
        elif re.search('fabric', str(sys.argv[2:])):
            process_Fabric(wb)
        elif re.search('tenant', str(sys.argv[2:])):
            process_Tenants(wb)
        elif re.search('vrf', str(sys.argv[2:])):
            process_VRF(wb)
        elif re.search('contract', str(sys.argv[2:])):
            process_Contracts(wb)
        elif re.search('l3out', str(sys.argv[2:])):
            process_L3Out(wb)
        elif re.search('mgmt', str(sys.argv[2:])):
            process_Mgmt_Tenant(wb)
        elif re.search('bd', str(sys.argv[2:])):
            process_Bridge_Domains(wb)
        elif re.search('dhcp', str(sys.argv[2:])):
            process_DHCP_Relay(wb)
        elif re.search('epg', str(sys.argv[2:])):
            process_EPGs(wb)
        elif re.search('vmm', str(sys.argv[2:])):
            process_VMM(wb)
        else:
            process_Best_Practices(wb)
            process_Fabric(wb)
            process_Admin(wb)
            process_Access(wb)
            process_Inventory(wb)
            process_Tenants(wb)
            process_L3Out(wb)
            process_Contracts(wb)
            process_Mgmt_Tenant(wb)
            process_VRF(wb)
            process_Bridge_Domains(wb)
            # process_DHCP_Relay(wb)
            process_EPGs(wb)
            # process_VMM(wb)
    else:
        process_Best_Practices(wb)
        process_Fabric(wb)
        process_Admin(wb)
        process_Access(wb)
        process_Inventory(wb)
        process_Tenants(wb)
        process_L3Out(wb)
        process_Contracts(wb)
        process_Mgmt_Tenant(wb)
        process_VRF(wb)
        process_Bridge_Domains(wb)
        # process_DHCP_Relay(wb)
        process_EPGs(wb)
        # process_VMM(wb)

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
