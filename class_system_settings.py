#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from class_terraform import terraform_cloud
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group
from easy_functions import update_easyDict
from openpyxl import load_workbook
import os
import validating

#======================================================
# Exception Classes
#======================================================
class InsufficientArgs(Exception):
    pass

class ErrException(Exception):
    pass

class InvalidArg(Exception):
    pass

class LoginFailed(Exception):
    pass

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Function.
#=====================================================================================
class system_settings(object):
    def __init__(self, type):
        self.type = type

    #======================================================
    # Function - APIC Connectivity Preference
    #======================================================
    def apic_preference(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.apicConnectivityPreference']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('apic_connectivity_preference', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'apic_connectivity_preference'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - BGP Autonomous System Number
    #======================================================
    def bgp_asn(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.bgpASN']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'bgp_asn'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - BGP Route Reflectors
    #======================================================
    def bgp_rr(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.bgpRouteReflector']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'bgp_rr'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Global AES Passphrase Encryption Settings
    #======================================================
    def global_aes(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.globalAesEncryptionSettings']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('enable_encryption', jsonData, **kwargs)

        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        if kwargs['enable_encryption'] == 'true':
            kwargs["Variable"] = 'aes_passphrase'
            kwargs['jsonData'] = jsonData
            sensitive_var_site_group(**kwargs)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'global_aes_encryption_settings'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Function.
#=====================================================================================
class site_policies(object):
    def __init__(self, type):
        self.type = type

    #======================================================
    # Function - Site Settings
    #======================================================
    def site_id(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['site.Identifiers']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Variables
            validating.validator('site_name', **kwargs)
            validating.url('controller', **kwargs)
            validating.values('auth_type', jsonData, **kwargs)
            validating.values('configure_terraform_cloud', jsonData, **kwargs)
            validating.values('controller_type', jsonData, **kwargs)
            validating.values('run_location', jsonData, **kwargs)
            if kwargs['controller_type'] == 'apic':
                validating.values('provider_version', jsonData, **kwargs)
            else:
                validating.values('provider_version', jsonData, **kwargs)
            validating.values('terraform_version', jsonData, **kwargs)
            if kwargs['controller_type'] == 'apic':
                validating.values('version', jsonData, **kwargs)
            else:
                validating.values('version', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify input information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
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
            kwargs['wb'].save(filename=site_wb)
            wb_wr = load_workbook(site_wb)
            ws_wr = wb_wr.get_sheet_names()
            for sheetName in ws_wr:
                if sheetName not in ['Sites']:
                    sheetToDelete = wb_wr.get_sheet_by_name(sheetName)
                    wb_wr.remove_sheet(sheetToDelete)
            wb_wr.save(filename=site_wb)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - Site Groups
    #======================================================
    def group_id(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['site.Groups']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        for x in range(1, 16):
            site = 'site_%s' % (x)
            if not kwargs[site] == None:
                validating.site_group('site_group', **kwargs)

        grp_count = 0
        for x in list(map(chr, range(ord('A'), ord('F')+1))):
            grp = 'Grp_%s' % (x)
            if kwargs['site_group'] == grp:
                grp_count += 1
        if grp_count == 0:
            ws = kwargs['ws']
            row_num = kwargs['row_num']
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
