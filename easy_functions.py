#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from collections import OrderedDict
from git import cmd, Repo
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from ordered_set import OrderedSet
from textwrap import fill
import jinja2
import json
import openpyxl
import os
import pkg_resources
import platform
import re
import requests
import shutil
import subprocess
import sys
import stdiomask
import time
import validating

# Global options for debugging
print_payload = False
print_response_always = False
print_response_on_fail = True

#======================================================
# Log Level - 0 = None, 1 = Class only, 2 = Line
#======================================================
log_level = 2

#======================================================
# Exception Classes
#======================================================
class InsufficientArgs(Exception):
    pass

#======================================================
# Function to GET to the APIC Config API
#======================================================
def apic_get(apic, cookies, uri, section=''):
    s = requests.Session()
    r = ''
    while r == '':
        try:
            r = s.get(
                'https://{}/{}.json'.format(apic, uri),
                cookies=cookies,
                verify=False
            )
            status = r.status_code
        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: {}"
                .format(e))
            time.sleep(5)
        except Exception as e:
            print("Method {} failed. Exception: {}".format(section[:-5], e))
            status = 666
            return(status)
    if print_response_always:
        print(r.text)
    if status != 200 and print_response_on_fail:
        print(r.text)
    return r

#======================================================
# Function to POST to the APIC Config API
#======================================================
def apic_post(apic, payload, cookies, uri, section=''):
    if print_payload:
        print(payload)
    s = requests.Session()
    r = ''
    while r == '':
        try:
            r = s.post('https://{}/{}.json'.format(apic, uri),
                    data=payload, cookies=cookies, verify=False)
            status = r.status_code
        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: {}"
                .format(e))
            time.sleep(5)
        except Exception as e:
            print("Method {} failed. Exception: {}".format(section[:-5], e))
            status = 666
            return(status)
    if print_response_always:
        print(r.text)
    if status != 200 and print_response_on_fail:
        print(r.text)
    return status

#======================================================
# Function to run 'terraform plan' and
# 'terraform apply' in the each folder of the
# Destination Directory.
#======================================================
def apply_terraform(args, folders, **easyDict):
    base_dir = args.dir
    jsonData = easyDict
    running_directory = os.getcwd()
    # tf_path = '.terraform/providers/registry.terraform.io/'
    opSystem = platform.system()
    if opSystem == 'Windows': path_sep = '\\'
    else: path_sep = '/'

    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'  Found the Followng Folders with uncommitted changes:\n')
    for folder in folders:
        print(f'  - {base_dir}{folder}')
    print(f'\n  Beginning Terraform Proceedures.')
    print(f'\n-----------------------------------------------------------------------------\n')

    tfe_dir = f'.terraform{path_sep}providers{path_sep}'
    tfe_cmd = subprocess.Popen(['terraform', 'init', '-upgrade'],
        cwd=running_directory,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT
    )
    output, err = tfe_cmd.communicate()
    tfe_cmd.wait()

    print(output.decode('utf-8'))
    response_p = ''
    response_a = ''
    for folder in folders:
        site_name = folder.split(path_sep)[0]
        site_match = 0
        for k, v in easyDict['sites']['site_settings'].items():
            if v[0]['site_name'] == site_name:
                run_loc = v[0]['run_location']
                site_match = 1
        if site_match == 0:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'  Could not determine the Run Location for folder {folder}:\n')
            print(f'  Defined Site Names not found in the Path.')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

        path = f'{base_dir}{folder}'
        if run_loc == 'local':
            if os.path.isfile(os.path.join(path, '.terraform.lock.hcl')):
                os.remove(os.path.join(path, '.terraform.lock.hcl'))
            if os.path.isdir(os.path.join(path, '.terraform')):
                shutil.rmtree(os.path.join(path, '.terraform'))
            lock_count = 0
            tfe_cmd = subprocess.Popen(
                ['terraform', 'init', f'-plugin-dir={running_directory}/{tfe_dir}'],
                cwd=path,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT
            )
            output, err = tfe_cmd.communicate()
            tfe_cmd.wait()
            print(output.decode('utf-8'))
            if 'does not match configured version' in output.decode('utf-8'):
                lock_count =+ 1

            if lock_count > 0:
                tfe_cmd = subprocess.Popen(
                    ['terraform', 'init', '-upgrade', f'-plugin-dir={running_directory}/{tfe_dir}']
                    , cwd=path
                )
                tfe_cmd.wait()
                print(output.decode('utf-8'))
            tfe_cmd = subprocess.Popen(['terraform', 'plan', '-out=main.plan'], cwd=path)
            output, err = tfe_cmd.communicate()
            tfe_cmd.wait()
            if not output == None:
                print(output.decode('utf-8'))
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
                tfe_cmd = subprocess.Popen(['terraform', 'apply', '-parallelism=1', 'main.plan'], cwd=path)
                tfe_cmd.wait()
                output, err = tfe_cmd.communicate()
                tfe_cmd.wait()
                if not output == None:
                    print(output.decode('utf-8'))
                print(f'\n--------------------------------------------------------------------------------\n')
                print(f'  Terraform Apply Complete.  Please Review for any errors and confirm next steps')
                print(f'\n--------------------------------------------------------------------------------\n')

        while True:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'  Folder: {path}.')
            print(f'  Please confirm if you want to commit the folder or just move forward.')
            print(f'  "C" to Commit the Folder and move forward.')
            print(f'  "M" to Move to the Next Folder.')
            print(f'  "Q" to Quit..')
            print(f'\n-----------------------------------------------------------------------------\n')
            response_a = input('  Please Enter ["C", "M" or "Q"]: ')
            if response_a == 'C':
                break
            elif response_a == 'M':
                break
            elif response_a == 'Q':
                exit()
            else:
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'  A Valid Response is either "C", "M" or "Q"...')
                print(f'\n-----------------------------------------------------------------------------\n')

        while True:
            if response_a == 'C':
                print(f'\n-----------------------------------------------------------------------------\n')
                commit_message = input(f'  Please Enter your Commit Message for the folder {folder}: ')
                baseRepo = Repo(args.dir)
                baseRepo.git.add(all=True)
                baseRepo.git.commit('-m', f'{commit_message}', '--', folder)
                baseRepo.git.push()
                break
            else:
                break

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
# Function to Create Interface Selectors
#======================================================
def create_selector(ws_sw, ws_sw_row_count, **templateVars):
    port_selector = ''
    for port in range(1, int(templateVars['port_count']) + 1):
        if port < 10:
            port_selector = 'Eth%s-0%s' % (templateVars['module'], port)
        elif port < 100:
            port_selector = 'Eth%s-%s' % (templateVars['module'], port)
        elif port > 99:
            port_selector = 'Eth%s_%s' % (templateVars['module'], port)
        modport = '%s/%s' % (templateVars['module'],port)
        # Copy the Port Selector to the Worksheet
        if templateVars['node_type'] == 'spine':
            data = [
                'intf_selector',
                templateVars['site_group'],
                templateVars['pod_id'],
                templateVars['node_id'],
                templateVars['switch_name'],
                port_selector,modport,
                'spine_pg','','','','',''
            ]
        else:
            data = [
                'intf_selector',
                templateVars['site_group'],
                templateVars['pod_id'],
                templateVars['node_id'],
                templateVars['switch_name'],
                port_selector,modport,
                '','','','','',''
            ]
        ws_sw.append(data)
        rc = f'{ws_sw_row_count}:{ws_sw_row_count}'
        for cell in ws_sw[rc]:
            if ws_sw_row_count % 2 == 0:
                cell.style = 'ws_odd'
            else:
                cell.style = 'ws_even'
        if templateVars['node_type'] == 'spine':
            templateVars['dv3'] = DataValidation(type="list", formula1='spine_pg', allow_blank=True)
        else:
            templateVars['dv3'] = DataValidation(type="list", formula1=f'INDIRECT(H{ws_sw_row_count})', allow_blank=True)
        ws_sw.add_data_validation(templateVars['dv3'])
        dv1_cell = f'A{ws_sw_row_count}'
        dv2_cell = f'H{ws_sw_row_count}'
        dv3_cell = f'I{ws_sw_row_count}'
        dv4_cell = f'K{ws_sw_row_count}'
        templateVars['dv1'].add(dv1_cell)
        templateVars['dv2'].add(dv2_cell)
        templateVars['dv3'].add(dv3_cell)
        templateVars['dv4'].add(dv4_cell)
        ws_sw_row_count += 1
    return ws_sw_row_count

#======================================================
# Function to Create Static Paths within EPGs
#======================================================
def create_static_paths(wb, wb_sw, row_num, wr_method, dest_dir, dest_file, template, **templateVars):
    wsheets = wb_sw.get_sheet_names()
    tf_file = ''
    for wsheet in wsheets:
        ws = wb_sw[wsheet]
        for row in ws.rows:
            if not (row[12].value == None or row[13].value == None):
                vlan_test = ''
                if re.search('^(individual|port-channel|vpc)$', row[7].value) and (re.search(r'\d+', str(row[12].value)) or re.search(r'\d+', str(row[13].value))):
                    if not row[12].value == None:
                        vlan = row[12].value
                        vlan_test = vlan_range(vlan, **templateVars)
                        if 'true' in vlan_test:
                            templateVars['mode'] = 'native'
                    if not 'true' in vlan_test:
                        templateVars['mode'] = 'regular'
                        if not row[13].value == None:
                            vlans = row[13].value
                            vlan_test = vlan_range(vlans, **templateVars)
                if vlan_test == 'true':
                    templateVars['Pod_ID'] = row[1].value
                    templateVars['Node_ID'] = row[2].value
                    templateVars['Interface_Profile'] = row[3].value
                    templateVars['Interface_Selector'] = row[4].value
                    templateVars['Port'] = row[5].value
                    templateVars['Policy_Group'] = row[6].value
                    templateVars['Port_Type'] = row[7].value
                    templateVars['Bundle_ID'] = row[9].value
                    Site_Group = templateVars['Site_Group']
                    pod = templateVars['Pod_ID']
                    node_id =  templateVars['Node_ID']
                    if templateVars['Port_Type'] == 'vpc':
                        ws_vpc = wb['Inventory']
                        for rx in ws_vpc.rows:
                            if rx[0].value == 'vpc_pair' and int(rx[1].value) == int(Site_Group) and str(rx[4].value) == str(node_id):
                                node1 = templateVars['Node_ID']
                                node2 = rx[5].value
                                templateVars['Policy_Group'] = '%s_vpc%s' % (row[3].value, templateVars['Bundle_ID'])
                                templateVars['tDn'] = 'topology/pod-%s/protpaths-%s-%s/pathep-[%s]' % (pod, node1, node2, templateVars['Policy_Group'])
                                templateVars['Static_Path'] = 'rspathAtt-[topology/pod-%s/protpaths-%s-%s/pathep-[%s]' % (pod, node1, node2, templateVars['Policy_Group'])
                                templateVars['GUI_Static'] = 'Pod-%s/Node-%s-%s/%s' % (pod, node1, node2, templateVars['Policy_Group'])
                                templateVars['Static_descr'] = 'Pod-%s_Nodes-%s-%s_%s' % (pod, node1, node2, templateVars['Policy_Group'])
                                tf_file = './ACI/%s/%s/%s' % (templateVars['Site_Name'], dest_dir, dest_file)
                                read_file = open(tf_file, 'r')
                                read_file.seek(0)
                                static_path_descr = 'resource "aci_epg_to_static_path" "%s_%s_%s"' % (templateVars['App_Profile'], templateVars['EPG'], templateVars['Static_descr'])
                                # if not static_path_descr in read_file.read():
                                #     create_tf_file(wr_method, dest_dir, dest_file, template, **templateVars)

                    elif templateVars['Port_Type'] == 'port-channel':
                        templateVars['Policy_Group'] = '%s_pc%s' % (row[3].value, templateVars['Bundle_ID'])
                        templateVars['tDn'] = 'topology/pod-%s/paths-%s/pathep-[%s]' % (pod, templateVars['Node_ID'], templateVars['Policy_Group'])
                        templateVars['Static_Path'] = 'rspathAtt-[topology/pod-%s/paths-%s/pathep-[%s]' % (pod, templateVars['Node_ID'], templateVars['Policy_Group'])
                        templateVars['GUI_Static'] = 'Pod-%s/Node-%s/%s' % (pod, templateVars['Node_ID'], templateVars['Policy_Group'])
                        templateVars['Static_descr'] = 'Pod-%s_Node-%s_%s' % (pod, templateVars['Node_ID'], templateVars['Policy_Group'])
                        read_file = open(tf_file, 'r')
                        read_file.seek(0)
                        static_path_descr = 'resource "aci_epg_to_static_path" "%s_%s_%s"' % (templateVars['App_Profile'], templateVars['EPG'], templateVars['Static_descr'])
                        # if not static_path_descr in read_file.read():
                        #     create_tf_file(wr_method, dest_dir, dest_file, template, **templateVars)

                    elif templateVars['Port_Type'] == 'individual':
                        port = 'eth%s' % (templateVars['Port'])
                        templateVars['tDn'] = 'topology/pod-%s/paths-%s/pathep-[%s]' % (pod, templateVars['Node_ID'], port)
                        templateVars['Static_Path'] = 'rspathAtt-[topology/pod-%s/paths-%s/pathep-[%s]' % (pod, templateVars['Node_ID'], port)
                        templateVars['GUI_Static'] = 'Pod-%s/Node-%s/%s' % (pod, templateVars['Node_ID'], port)
                        templateVars['Static_descr'] = 'Pod-%s_Node-%s_%s' % (pod, templateVars['Node_ID'], templateVars['Interface_Selector'])
                        read_file = open(tf_file, 'r')
                        read_file.seek(0)
                        static_path_descr = 'resource "aci_epg_to_static_path" "%s_%s_%s"' % (templateVars['App_Profile'], templateVars['EPG'], templateVars['Static_descr'])
                        # if not static_path_descr in read_file.read():
                        #     create_tf_file(wr_method, dest_dir, dest_file, template, **templateVars)
                        print('hello')

#======================================================
# Function to Append the easyDict Dictionary
#======================================================
def easyDict_append(templateVars, **kwargs):
    templateVars = OrderedDict(sorted(templateVars.items()))
    class_type = templateVars['class_type']
    data_type = templateVars['data_type']
    templateVars.pop('data_type')
    if not kwargs['easyDict'][class_type][data_type].get(kwargs['site_group']):
        kwargs['easyDict'][class_type][data_type].update({kwargs['site_group']:[]})
        
    kwargs['easyDict'][class_type][data_type][kwargs['site_group']].append(templateVars)
    return kwargs['easyDict']

#======================================================
# Function to Append the easyDict Dictionary
#======================================================
def easyDict_append_policy(templateVars, **kwargs):
    templateVars = OrderedDict(sorted(templateVars.items()))
    class_type = templateVars['class_type']
    data_type = templateVars['data_type']
    templateVars.pop('class_type')
    templateVars.pop('data_type')
    kwargs['easyDict'][class_type][data_type].update(templateVars)
    return kwargs['easyDict']

#======================================================
# Function to Append Subtype easyDict Dictionary
#======================================================
def easyDict_append_subtype(templateVars, **kwargs):
    templateVars = OrderedDict(sorted(templateVars.items()))
    class_type   = templateVars['class_type']
    data_type    = templateVars['data_type']
    data_subtype = templateVars['data_subtype']
    policy_name  = templateVars['policy_name']
    templateVars.pop('class_type')
    templateVars.pop('data_type')
    templateVars.pop('data_subtype')
    templateVars.pop('policy_name')
    templateVars.pop('site_group')
    if kwargs['easyDict'][class_type][data_type].get(kwargs['site_group']):
        for i in kwargs['easyDict'][class_type][data_type][kwargs['site_group']]:
            # print(json.dumps(i, indent=4))
            if class_type == 'tenants' and data_type == 'l3out_logical_node_profiles':
                if i['name'] == policy_name and i['tenant'] == templateVars['tenant'] and i['l3out'] == templateVars['l3out']:
                    templateVars.pop('l3out')
                    i[data_subtype].append(templateVars)
                    break
            elif class_type == 'tenants' and data_type == 'contracts':
                if i['name'] == policy_name and i['contract_type'] == templateVars['contract_type']:
                    templateVars.pop('contract_type')
                    templateVars.pop('tenant')
                    i[data_subtype].append(templateVars)
                    break
            elif class_type == 'tenants':
                if i['name'] == policy_name and i['tenant'] == templateVars['tenant']:
                    # templateVars.pop('tenant')
                    i[data_subtype].append(templateVars)
                    break
            else:
                if i['name'] == policy_name:
                    i[data_subtype].append(templateVars)
    elif 'Grp_' in kwargs['site_group']:
        site_group = kwargs['easyDict']['sites']['site_groups'][kwargs['site_group']][0]
        for site in site_group['sites']:
            for i in kwargs['easyDict'][class_type][data_type][str(site)]:
                if class_type == 'tenants':
                    if i['name'] == policy_name and i['tenant'] == templateVars['tenant']:
                        templateVars.pop['tenant']
                        i[data_subtype].append(templateVars)
                else:
                    if i['name'] == policy_name:
                        i[data_subtype].append(templateVars)

    # Return Dictionary
    return kwargs['easyDict']

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
# Function to Merge Easy ACI Repository to Dest Folder
#======================================================
def get_latest_versions(easyDict):
    # Get the Latest Release Tag for the provider-aci repository
    url = f'https://github.com/CiscoDevNet/terraform-provider-aci/tags/'
    r = requests.get(url, stream=True)
    repoVer = 'BLANK'
    stringMatch = False
    while stringMatch == False:
        for line in r.iter_lines():
            toString = line.decode("utf-8")
            if re.search(r'/releases/tag/v(\d+\.\d+\.\d+)\"', toString):
                repoVer = re.search('/releases/tag/v(\d+\.\d+\.\d+)', toString).group(1)
                break
        stringMatch = True
    
    # Set ACI Provider Version
    aci_provider_version = repoVer

    # Get the Latest Release Tag for the provider-mso repository
    url = f'https://github.com/CiscoDevNet/terraform-provider-mso/tags/'
    r = requests.get(url, stream=True)
    repoVer = 'BLANK'
    stringMatch = False
    while stringMatch == False:
        for line in r.iter_lines():
            toString = line.decode("utf-8")
            if re.search(r'/releases/tag/v(\d+\.\d+\.\d+)\"', toString):
                repoVer = re.search('/releases/tag/v(\d+\.\d+\.\d+)', toString).group(1)
                break
        stringMatch = True
    
    # Set NDO Provider Version
    ndo_provider_version = repoVer

    # Get the Latest Release Tag for Terraform
    url = f'https://github.com/hashicorp/terraform/tags'
    r = requests.get(url, stream=True)
    repoVer = 'BLANK'
    stringMatch = False
    while stringMatch == False:
        for line in r.iter_lines():
            toString = line.decode("utf-8")
            if re.search(r'/releases/tag/v(\d+\.\d+\.\d+)\"', toString):
                repoVer = re.search('/releases/tag/v(\d+\.\d+\.\d+)', toString).group(1)
                break
        stringMatch = True

    # Set Terraform Version
    terraform_version = repoVer

    easyDict['latest_versions']['aci_provider_version'] = aci_provider_version
    easyDict['latest_versions']['ndo_provider_version'] = ndo_provider_version
    easyDict['latest_versions']['terraform_version'] = terraform_version

    return easyDict

#======================================================
# Function to Check the Git Status of the Folders
#======================================================
def git_base_repo(args, wb):
    repoName = args.dir
    if not os.path.isdir(repoName):
        baseRepo = Repo.init(repoName, bare=True, mkdir=True)
        assert baseRepo.bare
    else:
        try: 
            baseRepo = Repo.init(repoName)
        except:
            baseRepo = Repo.init(repoName, bare=True, mkdir=True)
    base_dir = args.dir
    with baseRepo.config_reader() as git_config:
        try:
            git_config.get_value('user', 'email')
            git_config.get_value('user', 'name')
        except:
            valid = False
            while valid == False:
                templateVars = {}
                templateVars["Description"] = f'Git Email Configuration. i.e. user@example.com'
                templateVars["varInput"] = f'What is your Git email?'
                templateVars["minimum"] = 5
                templateVars["maximum"] = 128
                templateVars["pattern"] = '[\\S]+'
                templateVars["varName"] = 'Git Email'
                repoName = varStringLoop(**templateVars)
                valid = True

            valid = False
            while valid == False:
                templateVars = {}
                templateVars["Description"] = f'Git Username Configuration. i.e. user'
                templateVars["varInput"] = f'What is your Git Username?'
                templateVars["minimum"] = 5
                templateVars["maximum"] = 64
                templateVars["pattern"] = '[\\S]+'
                templateVars["varName"] = 'Git Email'
                repoName = varStringLoop(**templateVars)
                valid = True
    result = subprocess.Popen(['python3', '-m', 'git_status_checker', base_dir], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    while(True):
        # returns None while subprocess is running
        retcode = result.poll()
        line = result.stdout.readline()
        line = line.decode('utf-8')
        if 'this operation must be run in a work tree' in line:
            arg_folder = os.path.basename(os.path.normpath(args.dir))
            git_user = git_config.get_value('user', 'name')
            defaultUrl = f"github.com/{git_user}/{arg_folder}.git"
            valid = False
            while valid == False:
                templateVars = {}
                templateVars["varDefault"] = defaultUrl
                templateVars["Description"] = f'The Destination Directory is not currently a Git Repository.'
                templateVars["varInput"] = f'What is the Git URL (without https://) for "{args.dir}"? [{defaultUrl}]'
                templateVars["minimum"] = 5
                templateVars["maximum"] = 64
                templateVars["pattern"] = '[\\S]+'
                templateVars["varName"] = 'Git URL'
                gitUrl = varStringLoop(**templateVars)
                valid = True
            ws = wb['Sites']
            kwargs = {'ws':ws, 'row_num':0, 'url':gitUrl}
            validating.url('url', **kwargs)
            gitUrl = f'https://{gitUrl}'
            baseRepo.create_remote('origin', gitUrl)
            baseRepo.remotes.origin.push('master:master')
            break
        elif 'has outstanding commits' in line:
            break
        elif 'has outstanding pushes' in line:
            break
        
        print(baseRepo.is_dirty(untracked_files=True))
        try:
            baseRepo.remotes.origin.push('master:master')
        except Exception as e:
            print(f'Script Errored {e}')
            exit()
        break

#======================================================
# Function to Check the Git Status of the Folders
#======================================================
def git_check_status(args):
    baseRepo = Repo(args.dir)
    untrackedFiles = baseRepo.untracked_files
    random_folders = []
    modified = baseRepo.git.status()
    modifiedList = [y for y in (x.strip() for x in modified.splitlines()) if y]
    for line in modifiedList:
        if re.search(r'modified:   (.+\.auto\.tfvars)', line):
            file = re.search(r'modified:   (.+\.auto.tfvars)', line).group(1)
            dirname, filename = os.path.split(file)
            if not dirname in random_folders:
                random_folders.append(dirname)
            
    for file in untrackedFiles:
        dirname, filename = os.path.split(file)
        if not dirname in random_folders:
            random_folders.append(dirname)

    random_folders = list(set(random_folders))
    random_folders.sort()
    if not len(random_folders) > 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   There were no uncommitted changes in the environment.')
        print(f'   Proceedures Complete!!! Closing Environment and Exiting Script.')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

    strict_folders = []
    folder_order = ['access', 'common', 'mgmt']
    for folder in folder_order:
        for fx in random_folders:
            if folder in fx:
                strict_folders.append(fx)
    for folder in strict_folders:
        if folder in random_folders:
            random_folders.remove(folder)
    for folder in random_folders:
        strict_folders.append(folder)

    return strict_folders

#======================================================
# Function to Create Interface Selector Workbooks
#======================================================
def interface_selector_workbook(templateVars, **kwargs):
    # Set the Workbook var
    wb_sw = kwargs['wb_sw']

    # Use Switch_Type to Determine the Number of ports on the switch
    modules,port_count = switch_model_ports(kwargs['row_num'], templateVars['switch_model'])

    # Get the Interface Policy Groups from EasyDict
    if templateVars['node_type'] == 'spine':
        pg_list = ['spine_interface_policy_groups']
    else:
        pg_list = [
            'leaf_interfaces_policy_groups_access',
            'leaf_interfaces_policy_groups_breakout',
            'leaf_interfaces_policy_groups_bundle'
        ]
    switch_pgs = {}
    for pgroup in pg_list:
        switch_pgs[pgroup] = []
        for k, v in kwargs['easyDict']['access'][pgroup].items():
            if re.search('Grp_', k):
                site_group = kwargs['easyDict']['sites']['site_groups'][k][0]
                for site in site_group['sites']:
                    if int(templateVars['site_group']) == int(site):
                        for i in v:
                            switch_pgs[pgroup].append(i['name'])
            else:
                if int(k) == int(templateVars['site_group']):
                    for i in v:
                        switch_pgs[pgroup].append(i['name'])

    # Sort the Policy Group List and Convert to a string
    for pgroup in pg_list:
        switch_pgs[pgroup].sort()

    if not 'formulas' in wb_sw.sheetnames:
        ws_sw = wb_sw.create_sheet(title = 'formulas')
        ws_sw.column_dimensions['A'].width = 30
        ws_sw.column_dimensions['B'].width = 30
        ws_sw.column_dimensions['C'].width = 30
        ws_sw.column_dimensions['D'].width = 30
        data = ['access', 'breakout', 'bundle', 'spine_pg']
        ws_sw.append(data)
        for cell in ws_sw['1:1']:
            cell.style = 'Heading 1'
        wb_sw.save(kwargs['excel_workbook'])

    ws_sw = wb_sw['formulas']
    for pgroup in pg_list:
        if pgroup == 'leaf_interfaces_policy_groups_access': x = 1
        elif pgroup == 'leaf_interfaces_policy_groups_breakout': x = 2
        elif pgroup == 'leaf_interfaces_policy_groups_bundle': x = 3
        elif templateVars['node_type'] == 'spine': x = 4
        if len(switch_pgs[pgroup]) > 0:
            row_start = 2
            row_end = len(switch_pgs[pgroup]) + row_start - 1
            sw_pgs_count = 0
            for row_num in range(row_start, row_end + 1):
                rc = f'{row_num}:{row_num}'
                for cell in ws_sw[rc]:
                    if row_num % 2 == 0:
                        cell.style = 'ws_odd'
                    else:
                        cell.style = 'ws_even'
                ws_sw.cell(row=row_num, column=x, value=switch_pgs[pgroup][sw_pgs_count])
                sw_pgs_count += 1

        wb_sw.save(kwargs['excel_workbook'])
        last_row = len(switch_pgs[pgroup]) + 1
        defined_names = ['DSCP', 'leaf', 'spine', 'spine_modules', 'spine_type', 'switch_role', 'tag', 'Time_Zone']
        for dname in defined_names:
            if dname in wb_sw.defined_names:
                wb_sw.defined_names.delete(dname)
        if templateVars['node_type'] == 'spine':
            new_range = openpyxl.workbook.defined_name.DefinedName('spine_pg',attr_text=f"formulas!$D$2:$D{last_row}")
            if not 'spine_pg' in wb_sw.defined_names:
                # wb_sw.defined_names.delete('spine_pg')
                wb_sw.defined_names.append(new_range)
        elif pgroup == 'leaf_interfaces_policy_groups_access':
            new_range = openpyxl.workbook.defined_name.DefinedName('access',attr_text=f"formulas!$A$2:$A{last_row}")
            if not 'access' in wb_sw.defined_names:
                # wb_sw.defined_names.delete('access')
                wb_sw.defined_names.append(new_range)
        elif pgroup == 'leaf_interfaces_policy_groups_breakout':
            new_range = openpyxl.workbook.defined_name.DefinedName('breakout',attr_text=f"formulas!$B$2:$B{last_row}")
            if not 'breakout' in wb_sw.defined_names:
                # wb_sw.defined_names.delete('breakout')
                wb_sw.defined_names.append(new_range)
        elif pgroup == 'leaf_interfaces_policy_groups_bundle':
            new_range = openpyxl.workbook.defined_name.DefinedName('bundle',attr_text=f"formulas!$C$2:$C{last_row}")
            if not 'bundle' in wb_sw.defined_names:
                # wb_sw.defined_names.delete('bundle')
                wb_sw.defined_names.append(new_range)
        wb_sw.save(kwargs['excel_workbook'])

    # Check if there is a Worksheet for the Switch Already
    if not templateVars['switch_name'] in wb_sw.sheetnames:
        ws_sw = wb_sw.create_sheet(title = templateVars['switch_name'])
        ws_sw = wb_sw[templateVars['switch_name']]
        ws_sw.column_dimensions['A'].width = 15
        ws_sw.column_dimensions['B'].width = 15
        ws_sw.column_dimensions['C'].width = 10
        ws_sw.column_dimensions['D'].width = 10
        ws_sw.column_dimensions['E'].width = 20
        ws_sw.column_dimensions['F'].width = 20
        ws_sw.column_dimensions['G'].width = 20
        ws_sw.column_dimensions['H'].width = 20
        ws_sw.column_dimensions['I'].width = 20
        ws_sw.column_dimensions['J'].width = 20
        ws_sw.column_dimensions['K'].width = 20
        ws_sw.column_dimensions['L'].width = 25
        ws_sw.column_dimensions['M'].width = 30
        dv1 = DataValidation(type="list", formula1='"intf_selector"', allow_blank=True)
        if templateVars['node_type'] == 'spine':
            dv2 = DataValidation(type="list", formula1='"spine_pg"', allow_blank=True)
        else:
            dv2 = DataValidation(type="list", formula1='"access,breakout,bundle"', allow_blank=True)
        dv4 = DataValidation(type="list", formula1='"access,aaep_encap,trunk"', allow_blank=True)
        ws_sw.add_data_validation(dv1)
        ws_sw.add_data_validation(dv2)
        ws_sw.add_data_validation(dv4)
        ws_header = '%s Interface Selectors' % (kwargs['switch_name'])
        data = [ws_header]
        ws_sw.append(data)
        ws_sw.merge_cells('A1:M1')
        for cell in ws_sw['1:1']:
            cell.style = 'Heading 1'
        data = ['','Notes:']
        ws_sw.append(data)
        ws_sw.merge_cells('B2:M2')
        for cell in ws_sw['2:2']:
            cell.style = 'Heading 2'
        data = [
            'Type','site_group','pod_id','node_id','interface_profile','interface_selector','interface','policy_group_type',
            'policy_group','description','switchport_mode','access_or_native_vlan','trunk_port_allowed_vlans'
        ]
        ws_sw.append(data)
        for cell in ws_sw['3:3']:
            cell.style = 'Heading 3'

        ws_sw_row_count = 4
        templateVars['dv1'] = dv1
        templateVars['dv2'] = dv2
        templateVars['dv4'] = dv4
        templateVars['port_count'] = port_count
        sw_type = str(templateVars['switch_model'])
        if re.search('^(95[0-1][4-8])', sw_type):
            mod_keys = kwargs['easyDict']['access']['spine_modules'].keys()
            site_group = templateVars["site_group"]
            spine_modules = ''
            if site_group in mod_keys:
                spine_modules = kwargs['easyDict']['access']['spine_modules'][site_group][0]
            else:
                site_groups = kwargs['easyDict']['sites']['site_groups'].keys()
                x = []
                for s in site_groups:
                    if 'Grp_' in s:
                        x.append(s)
                site_groups = x
                for sgroup in site_groups:
                    if site_group in kwargs['easyDict']['sites']['site_groups'][sgroup][0]['sites']:
                        spine_modules = kwargs['easyDict']['access']['spine_modules'][sgroup][0]
            modDict = {}
            if not spine_modules == '':
                node_list = spine_modules['node_list']
                if str(templateVars['node_id']) in node_list:
                    modDict = spine_modules
            else:
                print(f"Error, Could not find the Module list for spine {templateVars['node_id']}")
                exit()
            
            for x in range(1, int(modules) + 1):
                module_type = modDict[f'module_{x}']
                if not module_type == None:
                    if re.search('^X97', module_type):
                        templateVars['module'] = x
                        templateVars['port_count'] = spine_module_port_count(module_type)
                        ws_sw_row_count = create_selector(ws_sw, ws_sw_row_count, **templateVars)
        else:
            templateVars['module'] = 1
            ws_sw_row_count = create_selector(ws_sw, ws_sw_row_count, **templateVars)

        # Save the Workbook
        wb_sw.save(kwargs['excel_workbook'])
        wb_sw.close()

#======================================================
# Function to Merge Easy ACI Repository to Dest Folder
#======================================================
def merge_easy_aci_repository(args, easy_jsonData, **easyDict):
    jsonData = easy_jsonData['components']['schemas']['easy_aci']['allOf'][1]['properties']
    baseRepo = args.dir
    
    # Setup Operating Environment
    opSystem = platform.system()
    tfe_dir = 'tfe_modules'
    if opSystem == 'Windows': path_sep = '\\'
    else: path_sep = '/'
    git_url = "https://github.com/terraform-cisco-modules/terraform-easy-aci"
    if not os.path.isdir(tfe_dir):
        os.mkdir(tfe_dir)
        Repo.clone_from(git_url, tfe_dir)
    else:
        g = cmd.Git(tfe_dir)
        g.pull()

    folders = []
    # Get All sub-folders from tfDir
    for k, v in easyDict['sites']['site_settings'].items():
        site_name = v[0]['site_name']
        site_dirs = next(os.walk(os.path.join(baseRepo, site_name)))[1]
        site_dirs.sort()
        for dir in site_dirs:
            folders.append(os.path.join(baseRepo, site_name, dir))
    
    # Now Loop over the folders and merge the module files
    module_folders = ['access', 'admin', 'fabric', 'switch', 'system_settings', 'tenant']
    for folder in folders:
        for mod in module_folders:
            if mod in folder:
                src_dir = os.path.join(tfe_dir, 'modules', mod)
                copy_files = os.listdir(src_dir)
                for fname in copy_files:
                    if not os.path.isdir(os.path.join(src_dir, fname)):
                        shutil.copy2(os.path.join(src_dir, fname), folder)

    # Loop over the folder list again and create blank auto.tfvars files for anything that doesn't already exist
    for folder in folders:
        if os.path.isdir(folder):
            for mod in module_folders:
                if mod in folder:
                    files = jsonData['files'][mod]
                    removeList = jsonData['remove_files']
                    for xRemove in removeList:
                        if xRemove in files:
                            files.remove(xRemove)
                    terraform_fmt(files, folder, path_sep)

#======================================================
# Function to GET to the NDO API
#======================================================
def ndo_get(ndo, cookies, uri, section=''):
    s = requests.Session()
    r = ''
    while r == '':
        try:
            r = s.get(
                'https://{}/{}'.format(ndo, uri),
                cookies=cookies,
                verify=False
            )
            status = r.status_code
        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: {}"
                .format(e))
            time.sleep(5)
        except Exception as e:
            print("Method {} failed. Exception: {}".format(section[:-5], e))
            status = 666
            return(status)
    if print_response_always:
        print(r.text)
    if status != 200 and print_response_on_fail:
        print(r.text)
    return r

#======================================================
# Function to validate input for each method
#======================================================
def process_kwargs(required_args, optional_args, **kwargs):
    # Validate all required kwargs passed
    # if all(item in kwargs for item in required_args.keys()) is not True:
    #    error_ = '\n***ERROR***\nREQUIRED Argument Not Found in Input:\n "%s"\nInsufficient required arguments.' % (item)
    #    raise InsufficientArgs(error_)
    row_num = kwargs["row_num"]
    ws = kwargs["ws"]
    error_count = 0
    error_list = []
    for item in required_args:
        if item not in kwargs.keys():
            error_count =+ 1
            error_list += [item]
    if error_count > 0:
        error_ = f'\n\n***Begin ERROR ***\n\nError on Worksheet {ws.title} row {row_num}\n - The Following REQUIRED Key(s) Were Not Found in kwargs: "{error_list}"\n\n****End ERROR****\n'
        raise InsufficientArgs(error_)

    error_count = 0
    error_list = []
    for item in optional_args:
        if item not in kwargs.keys():
            error_count =+ 1
            error_list += [item]
    if error_count > 0:
        error_ = f'\n\n***Begin ERROR***\n\nError on Worksheet {ws.title} row {row_num}\n - The Following Optional Key(s) Were Not Found in kwargs: "{error_list}"\n\n****End ERROR****\n'
        raise InsufficientArgs(error_)

    # Load all required args values from kwargs
    error_count = 0
    error_list = []
    for item in kwargs:
        if item in required_args.keys():
            required_args[item] = kwargs[item]
            if required_args[item] == None:
                error_count =+ 1
                error_list += [item]

    if error_count > 0:
        error_ = f'\n\n***Begin ERROR***\n\nError on Worksheet {ws.title} row {row_num}\n - The Following REQUIRED Key(s) Argument(s) are Blank:\nPlease Validate "{error_list}"\n\n****End ERROR****\n'
        raise InsufficientArgs(error_)

    for item in kwargs:
        if item in optional_args.keys():
            optional_args[item] = kwargs[item]
    # Combine option and required dicts for Jinja template render
    templateVars = {**required_args, **optional_args}
    return(templateVars)

#======================================================
# Function to Add Static Port Bindings to Bridge Domains Terraform Files
#======================================================
def process_workbook(templateVars, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    def process_site(siteDict, templateVars, **kwargs):
        # Create templateVars for Site_Name and APIC_URL
        templateVars['site_name'] =  siteDict['Site_Name']
        templateVars['site_group'] = siteDict['site_group']
        templateVars['controller'] =   siteDict['controller']

        # Pull in the Site Workbook
        excel_workbook = '%s_intf_selectors.xlsx' % (templateVars['site_name'])
        try:
            kwargs['wb_sw'] = load_workbook(excel_workbook)
        except Exception as e:
            print(f"Something went wrong while opening the workbook - {excel_workbook}... ABORT!")
            sys.exit(e)

        # Process the Interface Selectors for Static Port Paths
        create_static_paths(templateVars, **kwargs)

    if re.search('Grp_[A-F]', templateVars['site_group']):
        site_group = kwargs['easyDict']['sites']['site_groups'][kwargs['site_group']][0]
        for site in site_group['sites']:
            siteDict = kwargs['easyDict']['sites']['site_settings'][site][0]
            process_site(siteDict, templateVars, **kwargs)
    elif re.search(r'\d+', templateVars['Site_Group']):
        siteDict = kwargs['easyDict']['sites']['site_settings'][kwargs['site_group']][0]
        process_site(siteDict, templateVars, **kwargs)
    else:
        print(f"\n-----------------------------------------------------------------------------\n")
        print(f"   Error on Worksheet {ws.title}, Row {row_num} Site_Group, value {templateVars['Site_Group']}.")
        print(f"   Unable to Determine if this is a Single or Group of Site(s).  Exiting....")
        print(f"\n-----------------------------------------------------------------------------\n")
        exit()

#======================================================
# Function for Processing Loops to auto.tfvars files
#======================================================
def read_easy_jsonData(args, easy_jsonData, **easyDict):
    jsonData = easy_jsonData['components']['schemas']['easy_aci']['allOf'][1]['properties']
    classes = jsonData['classes']['enum']

    # Loop to write the Header and content to the files
    for class_type in classes:
        funcList = jsonData[f'class.{class_type}']['enum']
        for func in funcList:
            for k, v in easyDict[class_type][func].items():
                for i in v:
                    templateVars = i
                    kwargs = {
                        'args': args,
                        'easyDict': easyDict,
                        'row_num': f'{func}_section',
                        'site_group': k,
                        'ws': easyDict['wb']['System Settings']
                    }

                    # Add Variables for Template Functions
                    templateVars['template_type'] = func
                        
                    if re.search('^(apic_connectivity_preference|bgp_autonomous_system_number)$', func):
                        kwargs["template_file"] = 'template_open2.jinja2'
                    else:
                        kwargs["template_file"] = 'template_open.jinja2'

                    kwargs['tfvars_file'] = func
                    x = func.split('_')
                    policyType = ''
                    xcount = 0
                    for i in x:
                        if not i == 'and' and xcount == 0:
                            policyType = policyType + i.capitalize()
                        elif 'and' in i:
                            policyType = policyType + ' ' + i
                        else:
                            policyType = policyType + ' ' + i.capitalize()
                        xcount += 1
                    policyType = policyType.replace('Aaep', 'AAEP')
                    policyType = policyType.replace('Aes', 'AES')
                    policyType = policyType.replace('Apic', 'APIC')
                    policyType = policyType.replace('Cdp', 'CDP')
                    policyType = policyType.replace('Lldp', 'LLDP')
                    policyType = policyType.replace('Radius', 'RADIUS')
                    policyType = policyType.replace('Snmp', 'SNMP')
                    policyType = policyType.replace('Tacacs', 'TACACS+')
                    policyType = policyType.replace('Vpc', 'VPC')
                    templateVars['policy_type'] = policyType
                    
                    kwargs["initial_write"] = True
                    write_to_site(templateVars, **kwargs)

        for func in funcList:
            for k, v in easyDict[class_type][func].items():
                for i in v:
                    templateVars = i
                    kwargs = {
                        'args': args,
                        'easyDict': easyDict,
                        'row_num': f'{func}_section',
                        'site_group': k,
                        'ws': easyDict['wb']['System Settings']
                    }

                    # Write the template to the Template File
                    kwargs['tfvars_file'] = func
                    kwargs["initial_write"] = False
                    kwargs["template_file"] = f'{func}.jinja2'
                    write_to_site(templateVars, **kwargs)

    # Add Closing Bracket to auto.tfvars that are dictionaries    
    for k, v in easyDict['sites']['site_settings'].items():
        site_name = v[0]['site_name']
        siteDirs = next(os.walk(os.path.join(args.dir, site_name)))[1]
        siteDirs.sort()
        for folder in siteDirs:
            files = [f for f in os.listdir(os.path.join(args.dir, site_name, folder)) if 'auto.tfvars' in f]
            for file in files:
                if not re.search('(bgp_auto|connectivity|variables)', file):
                    file_name = open(os.path.join(args.dir, site_name, folder, file), 'r')
                    end_count = 0
                    for line in file_name:
                        if re.search(r'^}', line):
                            end_count += 1
                    file_name.close
                    if end_count == 0:
                            file_name = open(os.path.join(args.dir, site_name, folder, file), 'a+')
                            file_name.write('\n}\n')
                            file_name.close()

#======================================================
# Function to Read Excel Workbook Data
#======================================================
def read_in(excel_workbook):
    try:
        wb = load_workbook(excel_workbook)
    except Exception as e:
        print(f"Something went wrong while opening the workbook - {excel_workbook}... ABORT!")
        sys.exit(e)
    return wb

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
# Function to Add Required Arguments
#======================================================
def required_args_add(args_list, jsonData):
    for i in args_list:
        jsonData['required_args'].update({f'{i}': ''})
        jsonData['optional_args'].pop(i)
    return jsonData

#======================================================
# Function to Add Required Arguments
#======================================================
def required_args_remove(args_list, jsonData):
    for i in args_list:
        jsonData['optional_args'].update({f'{i}': ''})
        jsonData['required_args'].pop(i)
    return jsonData

#======================================================
# Function to loop through site_groups for sensitve vars
#======================================================
def sensitive_var_site_group(**kwargs):
    site_group = kwargs['site_group']
    sensitive_var = kwargs['Variable']

    # Add the Sensitive Variable to easyDict
    if 'tenants' in kwargs['class_type']:
        class_type = f"tenant_{kwargs['tenant']}"
    else:
        class_type = kwargs['class_type']
    if not 'tfcVariables' in class_type:
        if not kwargs['easyDict']['sensitive_vars'].get(site_group):
            kwargs['easyDict']['sensitive_vars'].update({site_group:{}})
        if not kwargs['easyDict']['sensitive_vars'][site_group].get(class_type):
            kwargs['easyDict']['sensitive_vars'][site_group].update({class_type:[]})
        kwargs['easyDict']['sensitive_vars'][site_group][class_type].append(sensitive_var)

    # Loop Through Site Groups to confirm Sensitive Variable in the Environment
    if re.search('Grp_[A-F]', site_group):
       siteGroup = kwargs['easyDict']['sites']['site_groups'][site_group][0]
       for site in siteGroup['sites']:
            siteDict = kwargs['easyDict']['sites']['site_settings'][site][0]
            if siteDict['run_location'] == 'local' or siteDict['configure_terraform_cloud'] == 'true':
                sensitive_var_value(**kwargs)
    else:
        siteDict = kwargs['easyDict']['sites']['site_settings'][site_group][0]
        if siteDict['run_location'] == 'local' or siteDict['configure_terraform_cloud'] == 'true':
            sensitive_var_value(**kwargs)
    return kwargs['easyDict']

#======================================================
# Function to add sensitive_var to Environment
#======================================================
def sensitive_var_value(**kwargs):
    sensitive_var = 'TF_VAR_%s' % (kwargs['Variable'])
    # -------------------------------------------------------------------------------------------------------------------------
    # Check to see if the Variable is already set in the Environment, and if not prompt the user for Input.
    #--------------------------------------------------------------------------------------------------------------------------
    if os.environ.get(sensitive_var) is None:
        print(f"\n----------------------------------------------------------------------------------\n")
        print(f"  The Script did not find {sensitive_var} as an 'environment' variable.")
        print(f"  To not be prompted for the value of {kwargs['Variable']} each time")
        print(f"  add the following to your local environemnt:\n")
        print(f"   - export {sensitive_var}='{kwargs['Variable']}_value'")
        print(f"\n----------------------------------------------------------------------------------\n")

    if os.environ.get(sensitive_var) is None:
        valid = False
        while valid == False:
            varValue = input('press enter to continue: ')
            if varValue == '':
                valid = True

        valid = False
        while valid == False:
            if kwargs.get('Multi_Line_Input'):
                print(f'Enter the value for {kwargs["Variable"]}:')
                lines = []
                while True:
                    # line = input('')
                    line = stdiomask.getpass(prompt='')
                    if line:
                        lines.append(line)
                    else:
                        break
                if not re.search('(certificate|private_key)', sensitive_var):
                    secure_value = '\\n'.join(lines)
                else:
                    secure_value = '\n'.join(lines)
            else:
                valid_pass = False
                while valid_pass == False:
                    password1 = stdiomask.getpass(prompt=f'Enter the value for {kwargs["Variable"]}: ')
                    password2 = stdiomask.getpass(prompt=f'Re-Enter the value for {kwargs["Variable"]}: ')
                    if password1 == password2:
                        secure_value = password1
                        valid_pass = True
                    else:
                        print('!!!Error!!! Sensitive Values did not match.  Please re-enter...')

            # Validate Sensitive Passwords
            cert_regex = re.compile(r'^\-{5}BEGIN (CERTIFICATE|PRIVATE KEY)\-{5}.*\-{5}END (CERTIFICATE|PRIVATE KEY)\-{5}$')
            if re.search('(certificate|certName|private_key|privateKey)', sensitive_var):
                if not re.search(cert_regex, secure_value):
                    valid = True
                else:
                    print(f'\n-------------------------------------------------------------------------------------------\n')
                    print(f'    Error!!! Invalid Value for the {sensitive_var}.  Please re-enter the {sensitive_var}.')
                    print(f'\n-------------------------------------------------------------------------------------------\n')
            elif re.search('(apikey|secretkey)', sensitive_var):
                if not sensitive_var == '':
                    valid = True
            else:
                if 'aes_passphrase' in sensitive_var:
                    sKey = 'aes_passphrase'
                    varTitle = 'Global AES Phassphrase'
                elif 'bgp_password' in sensitive_var:
                    sKey = 'password'
                    varTitle = 'BGP Password'
                elif 'apicPass' in sensitive_var:
                    sKey = 'password'
                    varTitle = 'APIC User Password.'
                elif 'eigrp_key' in sensitive_var:
                    sKey = 'eigrp_key'
                    varTitle = 'EIGRP Key.'
                elif 'ndoPass' in sensitive_var:
                    sKey = 'password'
                    varTitle = 'Nexus Dashboard Orchestrator User Password.'
                elif 'ntp_key' in sensitive_var:
                    sKey = 'key'
                    varTitle = 'NTP Key'
                elif 'ospf_key' in sensitive_var:
                    sKey = 'ospf_key'
                    varTitle = 'OSPF Authentication Password.'
                elif 'radius_key' in sensitive_var:
                    sKey = 'radius_key'
                    varTitle = 'The RADIUS shared secret cannot contain backslashes, space, or hashtag "#".'
                elif 'radius_monitoring_password' in sensitive_var:
                    sKey = 'radius_monitoring_password'
                    varTitle = 'RADIUS Monitoring Password.'
                elif re.search('snmp_(authorization|privacy)_key', sensitive_var):
                    sKey = 'snmp_key'
                    x = sensitive_var.split('_')
                    varType = '%s %s' % (x[0].capitalize(), x[1].capitalize())
                    varTitle = f'{varType}'
                elif 'snmp_community' in sensitive_var:
                    sKey = 'snmp_community'
                    varTitle = 'The Community may only contain letters, numbers and the special characters of \
                    underscore (_), hyphen (-), or period (.). The Community may not contain the @ symbol.'
                elif 'smtp_password' in sensitive_var:
                    sKey = 'smtp_password'
                    varTitle = 'Smart CallHome SMTP Server Password'
                elif 'tacacs_key' in sensitive_var:
                    sKey = 'tacacs_key'
                    varTitle = 'The TACACS+ shared secret cannot contain backslashes, space, or hashtag "#".'
                elif 'tacacs_monitoring_password' in sensitive_var:
                    sKey = 'tacacs_monitoring_password'
                    varTitle = 'TACACS+ Monitoring Password.'
                elif 'vmm_password' in sensitive_var:
                    sKey = 'vmm_password'
                    varTitle = 'Virtual Networking Controller Password.'
                else:
                    print(sensitive_var)
                    print('Could not Match Sensitive Value Type')
                    exit()
                minimum = kwargs['jsonData'][sKey]['minimum']
                maximum = kwargs['jsonData'][sKey]['maximum']
                pattern = kwargs['jsonData'][sKey]['pattern']
                valid = validating.length_and_regex_sensitive(pattern, varTitle, secure_value, minimum, maximum)

        # Add the Variable to the Environment
        os.environ[sensitive_var] = '%s' % (secure_value)
        var_value = secure_value

    else:
        # Add the Variable to the Environment
        if kwargs.get('Multi_Line_Input'):
            var_value = os.environ.get(sensitive_var)
            var_value = var_value.replace('\n', '\\n')
        else:
            var_value = os.environ.get(sensitive_var)

    return var_value

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

#======================================================
# Function to Determine Port count from Switch Model
#======================================================
def switch_model_ports(row_num, switch_type):
    modules = ''
    switch_type = str(switch_type)
    if re.search('^9396', switch_type):
        modules = '2'
        port_count = '48'
    elif re.search('^93', switch_type):
        modules = '1'

    if re.search('^9316', switch_type):
        port_count = '16'
    elif re.search('^(93120)', switch_type):
        port_count = '102'
    elif re.search('^(93108|93120|93216|93360)', switch_type):
        port_count = '108'
    elif re.search('^(93180|93240|9348|9396)', switch_type):
        port_count = '54'
    elif re.search('^(93240)', switch_type):
        port_count = '60'
    elif re.search('^9332', switch_type):
        port_count = '34'
    elif re.search('^(9336|93600)', switch_type):
        port_count = '36'
    elif re.search('^9364C-GX', switch_type):
        port_count = '64'
    elif re.search('^9364', switch_type):
        port_count = '66'
    elif re.search('^95', switch_type):
        port_count = '36'
        if switch_type == '9504':
            modules = '4'
        elif switch_type == '9508':
            modules = '8'
        elif switch_type == '9516':
            modules = '16'
        else:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Row {row_num}.  Unknown Switch Model {switch_type}')
            print(f'   Please verify Input Information.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()
    else:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Row {row_num}.  Unknown Switch Model {switch_type}')
        print(f'   Please verify Input Information.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    return modules,port_count

#======================================================
# Function to Determine Port Count on Modules
#======================================================
def spine_module_port_count(module_type):
    if re.search('X9716D-GX', module_type):
        port_count = '16'
    elif re.search('X9732C-EX', module_type):
        port_count = '32'
    elif re.search('X9736', module_type):
        port_count = '36'
    return port_count

#======================================================
# Function to GET from Terraform Cloud
#======================================================
def tfc_get(url, site_header, section=''):
    r = ''
    while r == '':
        try:
            r = requests.get(url, headers=site_header)
            status = r.status_code

            # Use this for Troubleshooting
            if print_response_always:
                print(status)
                print(r.text)

            if status == 200 or status == 404:
                json_data = r.json()
                return status,json_data
            else:
                validating.error_request(r.status_code, r.text)

        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: %s" % (e))
            time.sleep(5)
        except Exception as e:
            print("Method %s Failed. Exception: %s" % (section[:-5], e))
            exit()

#======================================================
# Function to PATCH to Terraform Cloud
#======================================================
def tfc_patch(url, payload, site_header, section=''):
    r = ''
    while r == '':
        try:
            r = requests.patch(url, data=payload, headers=site_header)

            # Use this for Troubleshooting
            if print_response_always:
                print(r.status_code)
                # print(r.text)

            if r.status_code != 200:
                validating.error_request(r.status_code, r.text)

            json_data = r.json()
            return json_data

        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: %s" % (e))
            time.sleep(5)
        except Exception as e:
            print("Method %s Failed. Exception: %s" % (section[:-5], e))
            exit()

#======================================================
# Function to POST to Terraform Cloud
#======================================================
def tfc_post(url, payload, site_header, section=''):
    r = ''
    while r == '':
        try:
            r = requests.post(url, data=payload, headers=site_header)

            # Use this for Troubleshooting
            if print_response_always:
                print(r.status_code)
                # print(r.text)

            if r.status_code != 201:
                validating.error_request(r.status_code, r.text)

            json_data = r.json()
            return json_data

        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: %s" % (e))
            time.sleep(5)
        except Exception as e:
            print("Method %s Failed. Exception: %s" % (section[:-5], e))
            exit()

#======================================================
# Function to Format Terraform Files
#======================================================
def terraform_fmt(files, folder, path_sep):
    # Create the Empty_variable_maps.auto.tfvars to house all the unused variables
    empty_auto_tfvars = f'{folder}{path_sep}Empty_variable_maps.auto.tfvars'
    wr_file = open(empty_auto_tfvars, 'w')
    wrString = f'#______________________________________________'\
              '\n#'\
              '\n# UNUSED Variables'\
              '\n#______________________________________________\n\n'
    wr_file.write(wrString)
    for file in files:
        varFiles = f"{file.split('.')[0]}.auto.tfvars"
        dest_file = f'{folder}{path_sep}{varFiles}'
        if not os.path.isfile(dest_file):
            x = file.split('.')
            wrString = f'{x[0]} = ''{}\n'
            wr_file.write(wrString)

    # Close the Unused Variables File
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
# Function to Validate Worksheet User Input
#======================================================
def validate_args(jsonData, **kwargs):
    globalData = kwargs['easy_jsonData']['components']['schemas']['globalData']['allOf'][1]['properties']
    global_args = [
        'admin_state',
        'application_epg',
        'application_profile',
        'annotation',
        'annotations',
        'audit_logs',
        'bridge_domain',
        'cdp_interface_policy',
        'description',
        'events',
        'faults',
        'global_alias',
        'l3out',
        'lldp_interface_policy',
        'login_domain',
        'management_epg',
        'management_epg_type',
        'monitoring_policy',
        'name',
        'name_alias',
        'node_id',
        'pod_id',
        'policies_tenant',
        'policy_name',
        'profile_name',
        'port_channel_policy',
        'qos_class',
        'schema',
        'session_logs',
        'sites',
        'target_dscp',
        'template',
        'tenant',
        'username',
        'vrf'
    ]
    for i in jsonData['required_args']:
        if i in global_args:
            if globalData[i]['type'] == 'integer':
                if kwargs[i] == None:
                    kwargs[i] = globalData[i]['default']
                else:
                    validating.number_check(i, globalData, **kwargs)
            elif globalData[i]['type'] == 'key_value':
                if not (kwargs[i] == None or kwargs[i] == ''):
                    validating.key_value(i, globalData, **kwargs)
            elif globalData[i]['type'] == 'list_of_string':
                if not (kwargs[i] == None or kwargs[i] == ''):
                    validating.string_list(i, globalData, **kwargs)
            elif globalData[i]['type'] == 'list_of_values':
                if kwargs[i] == None:
                    kwargs[i] = globalData[i]['default']
                else:
                    validating.list_values(i, globalData, **kwargs)
            elif globalData[i]['type'] == 'string':
                if not (kwargs[i] == None or kwargs[i] == ''):
                    validating.string_pattern(i, globalData, **kwargs)
            else:
                print(f'error validating.  Type not found {i}. 1')
                exit()
        elif i == 'site_group':
            validating.site_group('site_group', **kwargs)
        elif jsonData[i]['type'] == 'hostname':
            if not (kwargs[i] == None or kwargs[i] == ''):
                if ':' in kwargs[i]:
                    validating.ip_address(i, **kwargs)
                elif re.search('[a-z]', kwargs[i], re.IGNORECASE):
                    validating.dns_name(i, **kwargs)
                else:
                    validating.ip_address(i, **kwargs)
        elif jsonData[i]['type'] == 'email':
            if not (kwargs[i] == None or kwargs[i] == ''):
                validating.email(i, **kwargs)
        elif jsonData[i]['type'] == 'integer':
            if kwargs[i] == None:
                kwargs[i] = jsonData[i]['default']
            else:
                validating.number_check(i, jsonData, **kwargs)
        elif jsonData[i]['type'] == 'list_of_domains':
            if not (kwargs[i] == None or kwargs[i] == ''):
                count = 1
                for domain in kwargs[i]:
                    kwargs[f'domain_{count}'] = domain
                    validating.domain(f'domain_{count}', **kwargs)
                    kwargs.pop(f'domain_{count}')
                    count += 1
        elif jsonData[i]['type'] == 'list_of_hosts':
            if not (kwargs[i] == None or kwargs[i] == ''):
                count = 1
                for hostname in kwargs[i].split(','):
                    kwargs[f'{i}_{count}'] = hostname
                    if ':' in hostname:
                        validating.ip_address(f'{i}_{count}', **kwargs)
                    elif re.search('[a-z]', hostname, re.IGNORECASE):
                        validating.dns_name(f'{i}_{count}', **kwargs)
                    else:
                        validating.ip_address(f'{i}_{count}', **kwargs)
                    kwargs.pop(f'{i}_{count}')
                    count += 1
        elif jsonData[i]['type'] == 'list_of_integer':
            if kwargs[i] == None:
                kwargs[i] = jsonData[i]['default']
            else:
                validating.number_list(i, jsonData, **kwargs)
        elif jsonData[i]['type'] == 'list_of_string':
            if not (kwargs[i] == None or kwargs[i] == ''):
                validating.string_list(i, jsonData, **kwargs)
        elif jsonData[i]['type'] == 'list_of_values':
            if kwargs[i] == None:
                kwargs[i] = jsonData[i]['default']
            else:
                validating.list_values(i, jsonData, **kwargs)
        elif jsonData[i]['type'] == 'list_of_vlans':
            if not (kwargs[i] == None or kwargs[i] == ''):
                validating.vlans(i, **kwargs)
        elif jsonData[i]['type'] == 'string':
            if not (kwargs[i] == None or kwargs[i] == ''):
                validating.string_pattern(i, jsonData, **kwargs)
        else:
            print(f"error validating.  Type not found {jsonData[i]['type']}. 2.")
            exit()
    for i in jsonData['optional_args']:
        if not (kwargs[i] == None or kwargs[i] == ''):
            if i in global_args:
                if globalData[i]['type'] == 'integer':
                    validating.number_check(i, globalData, **kwargs)
                elif globalData[i]['type'] == 'key_value':
                    validating.key_value(i, globalData, **kwargs)
                elif globalData[i]['type'] == 'list_of_string':
                    validating.string_list(i, globalData, **kwargs)
                elif globalData[i]['type'] == 'list_of_values':
                    validating.list_values(i, globalData, **kwargs)
                elif globalData[i]['type'] == 'string':
                    validating.string_pattern(i, globalData, **kwargs)
                else:
                    validating.validator(i, **kwargs)
            elif re.search(r'^module_[\d]+$', i):
                validating.list_values_key('modules', i, jsonData, **kwargs)
            elif jsonData[i]['type'] == 'domain':
                validating.domain(i, **kwargs)
            elif jsonData[i]['type'] == 'email':
                validating.email(i, **kwargs)
            elif jsonData[i]['type'] == 'hostname':
                if ':' in kwargs[i]:
                    validating.ip_address(i, **kwargs)
                elif re.search('[a-z]', kwargs[i], re.IGNORECASE):
                    validating.dns_name(i, **kwargs)
                else:
                    validating.ip_address(i, **kwargs)
            elif jsonData[i]['type'] == 'integer':
                validating.number_check(i, jsonData, **kwargs)
            elif jsonData[i]['type'] == 'list_of_integer':
                validating.number_list(i, jsonData, **kwargs)
            elif jsonData[i]['type'] == 'list_of_hosts':
                count = 1
                for hostname in kwargs[i].split(','):
                    kwargs[f'{i}_{count}'] = hostname
                    if ':' in hostname:
                        validating.ip_address(f'{i}_{count}', **kwargs)
                    elif re.search('[a-z]', hostname, re.IGNORECASE):
                        validating.dns_name(f'{i}_{count}', **kwargs)
                    else:
                        validating.ip_address(f'{i}_{count}', **kwargs)
                    kwargs.pop(f'{i}_{count}')
                    count += 1
            elif jsonData[i]['type'] == 'list_of_macs':
                count = 1
                for mac in kwargs[i].split(','):
                    kwargs[f'{i}_{count}'] = mac
                    validating.mac_address(f'{i}_{count}', **kwargs)
                    kwargs.pop(f'{i}_{count}')
                    count += 1
            elif jsonData[i]['type'] == 'list_of_string':
                validating.string_list(i, jsonData, **kwargs)
            elif jsonData[i]['type'] == 'list_of_values':
                validating.list_values(i, jsonData, **kwargs)
            elif jsonData[i]['type'] == 'list_of_vlans':
                validating.vlans(i, **kwargs)
            elif jsonData[i]['type'] == 'mac_address':
                validating.mac_address(i, **kwargs)
            elif jsonData[i]['type'] == 'phone_number':
                validating.phone_number(i, **kwargs)
            elif jsonData[i]['type'] == 'string':
                validating.string_pattern(i, jsonData, **kwargs)
            else:
                print(f"error validating.  Type not found {jsonData[i]['type']}. 3.")
                exit()
    return kwargs

#======================================================
# Function to pull variables from easy_jsonData
#======================================================
def variablesFromAPI(**templateVars):
    valid = False
    while valid == False:
        json_vars = templateVars["jsonVars"]
        if 'popList' in templateVars:
            if len(templateVars["popList"]) > 0:
                for x in templateVars["popList"]:
                    varsCount = len(json_vars)
                    for r in range(0, varsCount):
                        if json_vars[r] == x:
                            json_vars.pop(r)
                            break
        print(f'\n-------------------------------------------------------------------------------------------\n')
        newDescr = templateVars["var_description"]
        if '\n' in newDescr:
            newDescr = newDescr.split('\n')
            for line in newDescr:
                if '*' in line:
                    print(fill(f'{line}',width=88, subsequent_indent='    '))
                else:
                    print(fill(f'{line}',88))
        else:
            print(fill(f'{templateVars["var_description"]}',88))
        print(f'\n    Select an Option Below:')
        for index, value in enumerate(json_vars):
            index += 1
            if value == templateVars["defaultVar"]:
                defaultIndex = index
            if index < 10:
                print(f'     {index}. {value}')
            else:
                print(f'    {index}. {value}')
        print(f'\n-------------------------------------------------------------------------------------------\n')
        if templateVars["multi_select"] == True:
            if not templateVars["defaultVar"] == '':
                var_selection = input(f'Please Enter the Option Number(s) to Select for {templateVars["varType"]}.  [{defaultIndex}]: ')
            else:
                var_selection = input(f'Please Enter the Option Number(s) to Select for {templateVars["varType"]}: ')
        else:
            if not templateVars["defaultVar"] == '':
                var_selection = input(f'Please Enter the Option Number to Select for {templateVars["varType"]}.  [{defaultIndex}]: ')
            else:
                var_selection = input(f'Please Enter the Option Number to Select for {templateVars["varType"]}: ')
        if not templateVars["defaultVar"] == '' and var_selection == '':
            var_selection = defaultIndex

        if templateVars["multi_select"] == False and re.search(r'^[0-9]+$', str(var_selection)):
            for index, value in enumerate(json_vars):
                index += 1
                if int(var_selection) == index:
                    selection = value
                    valid = True
        elif templateVars["multi_select"] == True and re.search(r'(^[0-9]+$|^[0-9\-,]+[0-9]$)', str(var_selection)):
            var_list = vlan_list_full(var_selection)
            var_length = int(len(var_list))
            var_count = 0
            selection = []
            for index, value in enumerate(json_vars):
                index += 1
                for vars in var_list:
                    if int(vars) == index:
                        var_count += 1
                        selection.append(value)
            if var_count == var_length:
                valid = True
            else:
                print(f'\n-------------------------------------------------------------------------------------------\n')
                print(f'  The list of Vars {var_list} did not match the available list.')
                print(f'\n-------------------------------------------------------------------------------------------\n')
        else:
            print(f'\n-------------------------------------------------------------------------------------------\n')
            print(f'  Error!! Invalid Selection.  Please Select a valid Option from the List.')
            print(f'\n-------------------------------------------------------------------------------------------\n')
    return selection

#======================================================
# Function to pull variables from easy_jsonData
#======================================================
def varBoolLoop(**templateVars):
    print(f'\n-------------------------------------------------------------------------------------------\n')
    newDescr = templateVars["Description"]
    if '\n' in newDescr:
        newDescr = newDescr.split('\n')
        for line in newDescr:
            if '*' in line:
                print(fill(f'{line}',width=88, subsequent_indent='    '))
            else:
                print(fill(f'{line}',88))
    else:
        print(fill(f'{templateVars["Description"]}',88))
    print(f'\n-------------------------------------------------------------------------------------------\n')
    valid = False
    while valid == False:
        varValue = input(f'{templateVars["varInput"]}  [{templateVars["varDefault"]}]: ')
        if varValue == '':
            if templateVars["varDefault"] == 'Y':
                varValue = True
            elif templateVars["varDefault"] == 'N':
                varValue = False
            valid = True
        elif varValue == 'N':
            varValue = False
            valid = True
        elif varValue == 'Y':
            varValue = True
            valid = True
        else:
            print(f'\n-------------------------------------------------------------------------------------------\n')
            print(f'   {templateVars["varName"]} value of "{varValue}" is Invalid!!! Please enter "Y" or "N".')
            print(f'\n-------------------------------------------------------------------------------------------\n')
    return varValue

#======================================================
# Function to pull variables from easy_jsonData
#======================================================
def varStringLoop(**templateVars):
    maximum = templateVars["maximum"]
    minimum = templateVars["minimum"]
    varName = templateVars["varName"]
    pattern = templateVars["pattern"]

    print(f'\n-------------------------------------------------------------------------------------------\n')
    newDescr = templateVars["Description"]
    if '\n' in newDescr:
        newDescr = newDescr.split('\n')
        for line in newDescr:
            if '*' in line:
                print(fill(f'{line}',width=88, subsequent_indent='    '))
            else:
                print(fill(f'{line}',88))
    else:
        print(fill(f'{templateVars["Description"]}',88))
    print(f'\n-------------------------------------------------------------------------------------------\n')
    valid = False
    while valid == False:
        varValue = input(f'{templateVars["varInput"]} ')
        if 'press enter to skip' in templateVars["varInput"] and varValue == '':
            valid = True
        elif not templateVars["varDefault"] == '' and varValue == '':
            varValue = templateVars["varDefault"]
            valid = True
        elif not varValue == '':
            valid = validating.length_and_regex(pattern, varName, varValue, minimum, maximum)
        else:
            print(f'\n-------------------------------------------------------------------------------------------\n')
            print(f'   {varName} value of "{varValue}" is Invalid!!! ')
            print(f'\n-------------------------------------------------------------------------------------------\n')
    return varValue

#======================================================
# Function to Expand the VLAN list
#======================================================
def vlan_list_full(vlan_list):
    full_vlan_list = []
    if re.search(r',', str(vlan_list)):
        vlist = vlan_list.split(',')
        for v in vlist:
            if re.fullmatch('^\\d{1,4}\\-\\d{1,4}$', v):
                a,b = v.split('-')
                a = int(a)
                b = int(b)
                vrange = range(a,b+1)
                for vl in vrange:
                    full_vlan_list.append(int(vl))
            elif re.fullmatch('^\\d{1,4}$', v):
                full_vlan_list.append(int(v))
    elif re.search('\\-', str(vlan_list)):
        a,b = vlan_list.split('-')
        a = int(a)
        b = int(b)
        vrange = range(a,b+1)
        for v in vrange:
            full_vlan_list.append(int(v))
    else:
        full_vlan_list.append(vlan_list)
    return full_vlan_list

#======================================================
# Function to Expand a VLAN Range to a VLAN List
#======================================================
def vlan_range(vlan_list, **templateVars):
    results = 'unknown'
    while results == 'unknown':
        if re.search(',', str(vlan_list)):
            vx = vlan_list.split(',')
            for vrange in vx:
                if re.search('-', vrange):
                    vl = vrange.split('-')
                    min_ = int(vl[0])
                    max_ = int(vl[1])
                    if (int(templateVars['VLAN']) >= min_ and int(templateVars['VLAN']) <= max_):
                        results = 'true'
                        return results
                else:
                    if templateVars['VLAN'] == vrange:
                        results = 'true'
                        return results
            results = 'false'
            return results
        elif re.search('-', str(vlan_list)):
            vl = vlan_list.split('-')
            min_ = int(vl[0])
            max_ = int(vl[1])
            if (int(templateVars['VLAN']) >= min_ and int(templateVars['VLAN']) <= max_):
                results = 'true'
                return results
        else:
            if int(templateVars['VLAN']) == int(vlan_list):
                results = 'true'
                return results
        results = 'false'
        return results

#======================================================
# Function to Determine which sites to write files to.
#======================================================
def write_to_site(templateVars, **kwargs):
    class_type = templateVars['class_type']
    aci_template_path = pkg_resources.resource_filename(f'classes', 'templates/')

    templateLoader = jinja2.FileSystemLoader(
        searchpath=(aci_template_path + '%s/') % (class_type))
    templateEnv = jinja2.Environment(loader=templateLoader)
    ws = kwargs["ws"]
    row_num = kwargs["row_num"]
    site_group = str(kwargs['site_group'])
    
    # Define the Template Source
    kwargs["template"] = templateEnv.get_template(kwargs["template_file"])

    # Process the template
    if 'tenants' in class_type:
        kwargs["dest_dir"] = 'tenant_%s' % (templateVars['tenant'])
    elif 'switches' in class_type:
        if templateVars['template_type'] == 'switch_profiles':
            if not templateVars['vpc_name'] == None:
                kwargs["dest_dir"] = 'switch_%s' % (templateVars['vpc_name'])
            else:
                kwargs["dest_dir"] = 'switch_%s' % (templateVars['switch_name'])
        elif templateVars['template_type'] == 'vpc_domains':
            kwargs["dest_dir"] = 'switch_%s' % (templateVars['name'])
    elif 'sites' in class_type:
        kwargs["dest_dir"] = kwargs["dest_dir"]
    else:
        kwargs["dest_dir"] = '%s' % (class_type)
    kwargs["dest_file"] = '%s.auto.tfvars' % (kwargs["tfvars_file"])
    if kwargs["initial_write"] == True:
        kwargs["write_method"] = 'w'
    else:
        kwargs["write_method"] = 'a'

    def process_siteDetails(site_dict, templateVars, **kwargs):
        # Create kwargs for site_name controller and controller_type
        kwargs['controller'] = site_dict.get('controller')
        kwargs['controller_type'] = site_dict.get('controller_type')
        templateVars['controller_type'] = site_dict.get('controller_type')
        kwargs['site_name'] = site_dict.get('site_name')
        kwargs['version'] = site_dict.get('version')
        if templateVars['template_type'] == 'firmware':
            templateVars['version'] = kwargs['version']
        if kwargs['controller_type'] == 'ndo' and templateVars['template_type'] == 'tenants':
            if templateVars['users'] == None:
                validating.error_tenant_users(**kwargs)
            else:
                for user in templateVars['users']:
                    regexp = '^[a-zA-Z0-9\_\-]+$'
                    validating.length_and_regex(regexp, 'users', user, 1, 64)
        # Create Terraform file from Template
        write_to_template(templateVars, **kwargs)

    if re.search('Grp_[A-F]', site_group):
        siteGroup = kwargs['easyDict']['sites']['site_groups'][kwargs['site_group']][0]
        for site in siteGroup['sites']:
            siteDict = kwargs['easyDict']['sites']['site_settings'][site][0]
            # Add Site Detials to kwargs and write to template
            process_siteDetails(siteDict, templateVars, **kwargs)

    elif re.search(r'\d+', site_group):
        siteDict = kwargs['easyDict']['sites']['site_settings'][kwargs['site_group']][0]

        # Add Site Detials to kwargs and write to template
        process_siteDetails(siteDict, templateVars, **kwargs)

    else:
        print(f"\n-----------------------------------------------------------------------------\n")
        print(f"   Error on Worksheet {ws.title}, Row {row_num} Site_Group, value {kwargs['site_group']}.")
        print(f"   Unable to Determine if this is a Single or Group of Site(s).  Exiting....")
        print(f"\n-----------------------------------------------------------------------------\n")
        exit()

#======================================================
# Function to write files from Templates
#======================================================
def write_to_template(templateVars, **kwargs):
    # Set Function Variables
    args = kwargs['args']
    baseRepo = args.dir
    dest_dir  = kwargs["dest_dir"]
    dest_file = kwargs["dest_file"]
    site_name = kwargs["site_name"]
    template  = kwargs["template"]
    wr_method = kwargs["write_method"]

    # Make sure the Destination Path and Folder Exist
    if not os.path.isdir(os.path.join(baseRepo, site_name, dest_dir)):
        opSystem = platform.system()
        if opSystem == 'Windows': path_sep = '\\'
        else: path_sep = '/'
        dest_path = f'{os.path.join(baseRepo, site_name)}{path_sep}{dest_dir}'
        os.makedirs(dest_path)
    dest_dir = os.path.join(baseRepo, site_name, dest_dir)
    if not os.path.exists(os.path.join(dest_dir, dest_file)):
        create_file = f'type nul >> {os.path.join(dest_dir, dest_file)}'
        os.system(create_file)
    tf_file = os.path.join(dest_dir, dest_file)
    wr_file = open(tf_file, wr_method)

    # Render Payload and Write to File
    templateVars = json.loads(json.dumps(templateVars))
    if templateVars['class_type'] == 'system_settings':
        payload = template.render(templateVars)
    else:
        templateVars = {'keys':templateVars}
        payload = template.render(templateVars)
    wr_file.write(payload)
    wr_file.close()
