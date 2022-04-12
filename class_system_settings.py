#!/usr/bin/env python3

import ast
import jinja2
import os
import re
import pkg_resources
import platform
import validating
from class_terraform import terraform_cloud
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group, sensitive_var_value
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
        required_args = {
            'row_num': '',
            'site_group': '',
            'wb': '',
            'ws': '',
            'apic_connectivity_preference': ''
        }
        optional_args = {}

        policy_type = 'APIC Connectivity Preference'

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        templateVars["header"] = '%s - Variables' % (policy_type)
        templateVars["initial_write"] = True
        templateVars["policy_type"] = policy_type
        templateVars["template_file"] = 'apic_connectivity_preference.jinja2'
        templateVars["template_type"] = 'apic_connectivity_preference'
        templateVars["tfvars_file"] = 'apic_connectivity_preference'

        try:
            # Validate Required Arguments
            validating.site_group(kwargs["wb"], kwargs["ws"], 'Site_Group', kwargs['site_group'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), kwargs["wb"], kwargs["row_num"])
            raise ErrException(Error_Return)

        # Write to the Template file
        write_to_site(self, **templateVars)

    def bgp_asn(self, **kwargs):
        # Dictionaries for required and optional args
        required_args = {
            'row_num': '',
            'site_group': '',
            'wb': '',
            'ws': '',
            'autonomous_system_number': ''
        }
        optional_args = {}

        policy_type = 'BGP - ASN'

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        templateVars["header"] = '%s - Variables' % (policy_type)
        templateVars["initial_write"] = True
        templateVars["policy_type"] = policy_type
        templateVars["template_file"] = 'bgp_autonomous_system_number.jinja2'
        templateVars["template_type"] = 'autonomous_system_number'
        templateVars["tfvars_file"] = 'bgp'

        try:
            # Validate Required Arguments
            validating.site_group(kwargs["wb"], kwargs["ws"], 'Site_Group', kwargs['site_group'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), kwargs["wb"], kwargs["row_num"])
            raise ErrException(Error_Return)

        # Write to the Template file
        write_to_site(self, **templateVars)

    def bgp_rr(self, **kwargs):
        # Dictionaries for required and optional args
        required_args = {
            'wb': '',
            'ws': '',
            'row_num': '',
            'site_group': '',
            'pod_id': '',
            'node_list': ''
        }
        optional_args = {}

        policy_type = 'BGP - Route Reflectors'

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        templateVars["header"] = '%s - Variables' % (policy_type)
        templateVars["initial_write"] = False
        templateVars["policy_type"] = policy_type
        templateVars["template_file"] = 'bgp_route_reflectors.jinja2'
        templateVars["template_type"] = 'bgp_route_reflectors'
        templateVars["tfvars_file"] = 'bgp'

        try:
            # Validate Required Arguments
            validating.site_group(kwargs["wb"], kwargs["ws"], 'Site_Group', kwargs['site_group'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), kwargs["wb"], kwargs["row_num"])
            raise ErrException(Error_Return)

        # Write to the Template file
        write_to_site(self, **templateVars)

    def global_aes(self, **kwargs):
        # Dictionaries for required and optional args
        required_args = {
            'easy_jsonData': '',
            'wb': '',
            'ws': '',
            'row_num': '',
            'site_group': '',
            'clear_passphrase': '',
            'enable_encryption': '',
            'passphrase_key_derivation_version': ''
        }
        optional_args = {}

        policy_type = 'Global AES Passphrase'

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        templateVars["header"] = '%s - Variables' % (policy_type)
        templateVars["initial_write"] = True
        templateVars["policy_type"] = policy_type
        templateVars["template_file"] = 'global_aes_encryption_settings.jinja2'
        templateVars["template_type"] = 'global_aes_encryption_settings'
        templateVars["tfvars_file"] = 'global_aes_encryption_settings'

        try:
            # Validate Required Arguments
            validating.site_group(kwargs["wb"], kwargs["ws"], 'Site_Group', kwargs['site_group'])
            validating.bool(kwargs["wb"], kwargs["ws"], 'enable_encryption', templateVars['enable_encryption'])

        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), kwargs["wb"], kwargs["row_num"])
            raise ErrException(Error_Return)

        if templateVars['enable_encryption'] == 'True':
            templateVars["Variable"] = 'aes_passphrase'
            sensitive_var_site_group(**templateVars)
        
        # Write to the Template file
        write_to_site(self, **templateVars)

# Class must be instantiated with Variables
class site_policies(object):
    def __init__(self, class_folder):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(aci_template_path + '%s/') % (class_folder))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def site_id(self, **kwargs):
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']
        # Dicts for required and optional args
        required_args = {
            'site_id': '',
            'site_name': '',
            'controller': '',
            'controller_type': '',
            'version': '',
            'auth_type': '',
            'terraform_version': '',
            'provider_version': '',
            'run_location': '',
            'configure_terraform_cloud': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Variables
            validating.name_complexity(row_num, ws, 'site_name', templateVars['site_name'])
            controller = 'https://%s' % (templateVars['controller'])
            validating.url(row_num, ws, 'controller', controller)
            validating.values(row_num, ws, 'version', templateVars['version'], ['5.2', '5.1', '5.0','4.2', '3.X'])
            validating.values(row_num, ws, 'auth_type', templateVars['auth_type'], ['ssh-key', 'username'])
            validating.values(row_num, ws, 'run_location', templateVars['run_location'], ['local', 'tfc'])
            validating.not_empty(row_num, ws, 'provider_version', templateVars['provider_version'])
            validating.not_empty(row_num, ws, 'terraform_version', templateVars['terraform_version'])
            validating.bool(row_num, ws, 'configure_terraform_cloud', templateVars['configure_terraform_cloud'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), kwargs["wb"], kwargs["row_num"])
            raise ErrException(Error_Return)

        # Save the Site Information into Environment Variables
        site_id = 'site_id_%s' % (templateVars['site_id'])
        os.environ[site_id] = '%s' % (templateVars)

        folder_list = ['access', 'admin', 'fabric', 'system_settings']
        # file_list = ['provider.jinja2_provider.tf', 'variables.jinja2_variables.tf']
        #    
        # # Write the Files to the Appropriate Directories
        # if templateVars['controller_type'] == 'apic':
        #     for folder in folder_list:
        #         for file in file_list:
        #             x = file.split('_')
        #             template_file = x[0]
        #             templateVars["dest_dir"] = folder
        #             templateVars["dest_file"] = x[1]
        #             templateVars["template"] = self.templateEnv.get_template(template_file)
        #             templateVars["write_method"] = 'w'
        #             write_to_template(**templateVars)

            # If the state_location is tfc configure workspaces in the cloud
        if templateVars['run_location'] == 'tfc' and templateVars['configure_terraform_cloud'] == True:
            # Initialize the Class
            class_init = '%s()' % ('lib_terraform.Terraform_Cloud')

            # Get terraform_cloud_token
            terraform_cloud().terraform_token()

            # Get workspace_ids
            easy_jsonData = kwargs['easy_jsonData']
            terraform_cloud().create_terraform_workspaces(easy_jsonData, folder_list, templateVars["site_name"])

            if templateVars['auth_type'] == 'user_pass' and templateVars["controller_type"] == 'apic':
                var_list = ['apicUrl', 'aciUser', 'aciPass']
            elif templateVars["controller_type"] == 'apic':
                var_list = ['apicUrl', 'certName', 'privateKey']
            else:
                var_list = ['ndoUrl', 'ndoDomain', 'ndoUser', 'ndoPass']

            # Get var_ids
            tf_var_dict = {}
            for folder in folder_list:
                folder_id = 'site_id_%s_%s' % (templateVars['site_id'], folder)
                # kwargs['workspace_id'] = workspace_dict[folder_id]
                kwargs['description'] = ''
                # for var in var_list:
                #     tf_var_dict = tf_variables(class_init, folder, var, tf_var_dict, **kwargs)

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
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']
        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'site_1': '',
            'site_2': ''
        }
        optional_args = {
            'site_2': '',
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
            'site_15': ''
        }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        for x in range(1, 16):
            site = 'site_%s' % (x)
            if not templateVars[site] == None:
                validating.site_group(wb, ws, site, templateVars[site])

        grp_count = 0
        for x in list(map(chr, range(ord('A'), ord('F')+1))):
            grp = 'Grp_%s' % (x)
            if templateVars['site_group'] == grp:
                grp_count += 1
        if grp_count == 0:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} group, group_name "{kwargs["group"]}" is invalid.')
            print(f'   A valid Group Name is Grp_A thru Grp_F.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

        # Save the Site Information into Environment Variables
        group_id = '%s' % (templateVars['site_group'])
        os.environ[group_id] = '%s' % (templateVars)
