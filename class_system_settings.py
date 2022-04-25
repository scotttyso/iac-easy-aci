#!/usr/bin/env python3

import jinja2
import os
import pkg_resources
import re
import validating
from class_terraform import terraform_cloud
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group
from easy_functions import update_easyDict
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
        # Set Locally Used Variables
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dictionaries for required and optional args
        required_args = {
            'site_group': '',
            'apic_connectivity_preference': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        # Get Variable Values
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.apicPreference']['allOf'][1]['properties']

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', kwargs['site_group'])
            validating.values(row_num, ws, 'apic_connectivity_preference', kwargs['apic_connectivity_preference'], 
                jsonData['apic_connectivity_preference']['enum'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'apic_connectivity_preference'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    def bgp_asn(self, **kwargs):
        # Set Locally Used Variables
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dictionaries for required and optional args
        required_args = {
            'site_group': '',
            'autonomous_system_number': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, kwargs["ws"], 'Site_Group', kwargs['site_group'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'bgp_asn'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    def bgp_rr(self, **kwargs):
        # Set Locally Used Variables
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dictionaries for required and optional args
        required_args = {
            'site_group': '',
            'pod_id': '',
            'node_list': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', kwargs['site_group'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'bgp_rr'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    def global_aes(self, **kwargs):
        # Set Locally Used Variables
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dictionaries for required and optional args
        required_args = {
            'site_group': '',
            'clear_passphrase': '',
            'enable_encryption': '',
            'passphrase_key_derivation_version': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', kwargs['site_group'])
            validating.values(row_num, ws, 'enable_encryption', kwargs['enable_encryption'], ['true', 'false'])

        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' \
                % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        if kwargs['enable_encryption'] == 'true':
            kwargs["Variable"] = 'aes_passphrase'
            sensitive_var_site_group(**kwargs)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'global_aes_encryption_settings'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

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
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {
            'auth_type': '',
            'configure_terraform_cloud': '',
            'controller': '',
            'controller_type': '',
            'provider_version': '',
            'run_location': '',
            'site_id': '',
            'site_name': '',
            'terraform_version': '',
            'version': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        # Get Variable Values
        jsonData = kwargs['easy_jsonData']['components']['schemas']['site.Variables']['allOf'][1]['properties']

        try:
            # Validate Variables
            validating.name_complexity(row_num, ws, 'site_name', kwargs['site_name'])
            controller = 'https://%s' % (kwargs['controller'])
            validating.url(row_num, ws, 'controller', controller)
            validating.values(row_num, ws, 'auth_type', kwargs['auth_type'],
                jsonData['auth_type']['enum'])
            validating.values(row_num, ws, 'configure_terraform_cloud', kwargs['configure_terraform_cloud'],
                jsonData['configure_terraform_cloud']['enum'])
            validating.values(row_num, ws, 'controller_type', kwargs['controller_type'],
                jsonData['controller_type']['enum'])
            validating.values(row_num, ws, 'run_location', kwargs['run_location'],
                jsonData['run_location']['enum'])
            if kwargs['controller_type'] == 'apic':
                validating.values(row_num, ws, 'provider_version', kwargs['provider_version'],
                    jsonData['provider_version_apic']['enum'])
            else:
                validating.values(row_num, ws, 'provider_version', kwargs['provider_version'],
                    jsonData['provider_version_ndo']['enum'])
            validating.values(row_num, ws, 'terraform_version', kwargs['terraform_version'],
                jsonData['terraform_version']['enum'])
            if kwargs['controller_type'] == 'apic':
                validating.values(row_num, ws, 'version', kwargs['version'],
                    jsonData['version_apic']['enum'])
            else:
                validating.values(row_num, ws, 'version', kwargs['version'],
                    jsonData['version_ndo']['enum'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Save the Site Information into Environment Variables
        site_id = 'site_id_%s' % (kwargs['site_id'])
        os.environ[site_id] = '%s' % (templateVars)

        folder_list = ['access', 'admin', 'fabric', 'system_settings']
        # file_list = ['provider.jinja2_provider.tf', 'variables.jinja2_variables.tf']
        #    
        # # Write the Files to the Appropriate Directories
        # if kwargs['controller_type'] == 'apic':
        #     for folder in folder_list:
        #         for file in file_list:
        #             x = file.split('_')
        #             template_file = x[0]
        #             kwargs["dest_dir"] = folder
        #             kwargs["dest_file"] = x[1]
        #             kwargs["template"] = self.templateEnv.get_template(template_file)
        #             kwargs["write_method"] = 'w'
        #             write_to_template(**kwargs)

            # If the state_location is tfc configure workspaces in the cloud
        if kwargs['run_location'] == 'tfc' and kwargs['configure_terraform_cloud'] == 'true':
            # Initialize the Class
            class_init = '%s()' % ('lib_terraform.Terraform_Cloud')

            # Get terraform_cloud_token
            terraform_cloud().terraform_token()

            # Get workspace_ids
            easy_jsonData = kwargs['easy_jsonData']
            terraform_cloud().create_terraform_workspaces(easy_jsonData, folder_list, kwargs["site_name"])

            if kwargs['auth_type'] == 'user_pass' and kwargs["controller_type"] == 'apic':
                var_list = ['apicUrl', 'aciUser', 'aciPass']
            elif kwargs["controller_type"] == 'apic':
                var_list = ['apicUrl', 'certName', 'privateKey']
            else:
                var_list = ['ndoUrl', 'ndoDomain', 'ndoUser', 'ndoPass']

            # Get var_ids
            tf_var_dict = {}
            for folder in folder_list:
                folder_id = 'site_id_%s_%s' % (kwargs['site_id'], folder)
                # kwargs['workspace_id'] = workspace_dict[folder_id]
                kwargs['description'] = ''
                # for var in var_list:
                #     tf_var_dict = tf_variables(class_init, folder, var, tf_var_dict, **kwargs)

        site_wb = '%s_intf_selectors.xlsx' % (kwargs['site_name'])
        if not os.path.isfile(site_wb):
            wb.save(filename=site_wb)
            wb_wr = load_workbook(site_wb)
            ws_wr = wb_wr.get_sheet_names()
            for sheetName in ws_wr:
                if sheetName not in ['Sites']:
                    sheetToDelete = wb_wr.get_sheet_by_name(sheetName)
                    wb_wr.remove_sheet(sheetToDelete)
            wb_wr.save(filename=site_wb)

        # Return Dictionary
        return kwargs['easyDict']

    # Method must be called with the following kwargs.
    # Group: Required.  A Group Name to represent a list of Site_ID's
    # site_1: Required.  The site_id for the First Site
    # site_2: Required.  The site_id for the Second Site
    # site_[3-15]: Optional.  The site_id for the 3rd thru the 15th Site(s)
    def group_id(self, **kwargs):
        # Set Locally Used Variables
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
            if not kwargs[site] == None:
                validating.site_group(row_num, ws, site, kwargs[site])

        grp_count = 0
        for x in list(map(chr, range(ord('A'), ord('F')+1))):
            grp = 'Grp_%s' % (x)
            if kwargs['site_group'] == grp:
                grp_count += 1
        if grp_count == 0:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} group, group_name "{kwargs["group"]}" is invalid.')
            print(f'   A valid Group Name is Grp_A thru Grp_F.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

        # Save the Site Information into Environment Variables
        group_id = '%s' % (kwargs['site_group'])
        os.environ[group_id] = '%s' % (templateVars)

        # Return Dictionary
        return kwargs['easyDict']
