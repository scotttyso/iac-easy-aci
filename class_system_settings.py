#!/usr/bin/env python3

import jinja2
import os
import re
import pkg_resources
import platform
import validating
from class_terraform import terraform_cloud
from easy_functions import process_kwargs
from easy_functions import write_to_site
from easy_functions import write_to_template
from openpyxl import load_workbook

aci_template_path = pkg_resources.resource_filename('class_system_settings', 'templates/')

# Exception Classes
class InsufficientArgs(Exception):
    pass

class ErrException(Exception):
    pass

class InvalidArg(Exception):
    pass

class LoginFailed(Exception):
    pass

class system_settings(object):
    def __init__(self, type):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(aci_template_path + '%s/') % (type))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)
        self.type = type

    #==============================================
    # Function - APIC Connectivity Preference
    #==============================================
    def apic_preference(self, **kwargs):
        # Dictionaries for required and optional args
        required_args = {'row_num': '',
                         'site_group': '',
                         'wb': '',
                         'ws': '',
                         'apic_connectivity_preference': ''}
        optional_args = {}

        policy_type = 'APIC Connectivity Preference'

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        templateVars["header"] = '%s Variables' % (policy_type)
        templateVars["initial_write"] = True
        templateVars["policy_type"] = policy_type
        templateVars["template_file"] = 'apic_connectivity_preference.jinja2'
        templateVars["template_type"] = 'apic_connectivity_preference'

        try:
            # Validate Required Arguments
            validating.site_group(kwargs["wb"], kwargs["ws"], 'Site_Group', kwargs['site_group'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), kwargs["wb"], kwargs["row_num"])
            raise ErrException(Error_Return)

        # Write to the Template file
        write_to_site(self, **templateVars)

# Class must be instantiated with Variables
class Site_Policies(object):
    def __init__(self, ws):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(aci_template_path))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def site_id(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'row_num': '',
                         'site_group': '',
                         'wb': '',
                         'ws': '',
                         'site_id': '',
                         'site_name': '',
                         'controller': '',
                         'version': '',
                         'auth_type': '',
                         'terraform_version': '',
                         'provider_version': '',
                         'run_location': '',
                         'state_location': ''}
        optional_args = {'tfc_organization': '',
                         'workspace_prefix': '',
                         'vcs_repo': '',
                         'agent_pool': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Variables
            validating.name_complexity(row_num, ws, 'site_name', templateVars['site_name'])
            controller = 'https://%s' % (templateVars['controller'])
            validating.url(row_num, ws, 'controller', controller)
            validating.values(row_num, ws, 'version', templateVars['version'], ['3.X', '4.X', '5.X'])
            validating.values(row_num, ws, 'auth_type', templateVars['auth_type'], ['ssh-key', 'user_pass'])
            validating.values(row_num, ws, 'state_location', templateVars['state_location'], ['local', 'tfc'])
            validating.values(row_num, ws, 'run_location', templateVars['run_location'], ['local', 'tfc'])
            validating.not_empty(row_num, ws, 'provider_version', templateVars['provider_version'])
            validating.not_empty(row_num, ws, 'terraform_version', templateVars['terraform_version'])
            if templateVars['State_Location'] == 'tfc':
                validating.not_empty(row_num, ws, 'tfc_organization', templateVars['tfc_organization'])
                validating.not_empty(row_num, ws, 'vcs_repo', templateVars['vcs_repo'])
                validating.not_empty(row_num, ws, 'agent_pool', templateVars['agent_pool'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), kwargs["wb"], kwargs["row_num"])
            raise ErrException(Error_Return)

        # Save the Site Information into Environment Variables
        site_id = 'site_id_%s' % (templateVars['site_id'])
        os.environ[site_id] = '%s' % (templateVars)

        folder_list = ['access', 'admin', 'fabric', 'system_settings']
        file_list = ['.gitignore_.gitignore', 'provider.jinja2_provider.tf', 'variables.jinja2_variables.tf']

        # Write the .gitignore to the Appropriate Directories
        if templateVars['controller_type'] == 'apic':
            for folder in folder_list:
                for file in file_list:
                    x = file.split('_')
                    template_file = x[0]
                    templateVars["dest_dir"] = folder
                    templateVars["dest_file"] = x[1]
                    templateVars["template"] = self.templateEnv.get_template(template_file)
                    templateVars["write_method"] = 'w'
                    write_to_template(**templateVars)

            # If the state_location is tfc configure workspaces in the cloud
            if templateVars['state_location'] == 'tfc' or templateVars['run_location'] == 'tfc':
                # Initialize the Class
                class_init = '%s()' % ('lib_terraform.Terraform_Cloud')

                # Get terraform_cloud_token
                templateVars["terraform_cloud_token"] = terraform_cloud().terraform_token()

                # Get workspace_ids
                workspace_dict = {}
                for folder in folder_list:
                    workspace_dict = tf_workspace(class_init, folder, workspace_dict, **kwargs)

            # If the run_location is Terraform_Cloud Configure Variables in the Cloud
            if templateVars['run_location'] == 'tfc':
                if templateVars['auth_type'] == 'user_pass':
                    var_list = ['apicUrl', 'aciUser', 'aciPass']
                else:
                    var_list = ['apicUrl', 'certName', 'privateKey']

                # Get var_ids
                tf_var_dict = {}
                for folder in folder_list:
                    folder_id = 'site_id_%s_%s' % (templateVars['site_id'], folder)
                    kwargs['workspace_id'] = workspace_dict[folder_id]
                    kwargs['description'] = ''
                    for var in var_list:
                        tf_var_dict = tf_variables(class_init, folder, var, tf_var_dict, **kwargs)

        site_wb = '%s_intf_selectors.xlsx' % (templateVars['site_name'])
        if not os.path.isfile(site_wb):
            wb.save(filename=site_wb)
            wb_wr = load_workbook(site_wb)
            ws_wr = wb_wr.get_sheet_names()
            for sheetName in ws_wr:
                if sheetName not in ['Sites']:
                    sheetToDelete = wb_wr.get_sheet_by_name(sheetName)
                    wb_wr.remove_sheet(sheetToDelete)
            wb_wr.save(filename=site_wb)

    # Method must be called with the following kwargs.
    # Group: Required.  A Group Name to represent a list of Site_ID's
    # site_1: Required.  The site_id for the First Site
    # site_2: Required.  The site_id for the Second Site
    # site_[3-15]: Optional.  The site_id for the 3rd thru the 15th Site(s)
    def group_id(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'row_num': '',
                         'site_group': '',
                         'wb': '',
                         'ws': '',
                         'group': '',
                         'site_1': ''}
        optional_args = {'site_2': '',
                         'site_3': '',
                         'site_4': '',
                         'site_5': '',
                         'site_6': '',
                         'site_7': '',
                         'site_8': '',
                         'site_9': '',
                         'site_10': '',
                         'site_11': '',
                         'site_12': '',
                         'site_13': '',
                         'site_14': '',
                         'site_15': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)
        row_num = templateVars["row_num"]
        wb = templateVars["wb"]
        ws = templateVars["ws"]

        for x in range(1, 16):
            site = 'site_%s' % (x)
            if not templateVars[site] == None:
                validating.site_group(wb, ws, site, templateVars[site])

        grp_count = 0
        for x in list(map(chr, range(ord('A'), ord('F')+1))):
            grp = 'Grp_%s' % (x)
            if templateVars['Group'] == grp:
                grp_count += 1
        if grp_count == 0:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} group, group_name "{kwargs["group"]}" is invalid.')
            print(f'   A valid Group Name is Grp_A thru Grp_F.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

        # Save the Site Information into Environment Variables
        group_id = '%s' % (templateVars['group'])
        os.environ[group_id] = '%s' % (templateVars)
