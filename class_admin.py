#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from easy_functions import countKeys, findVars
from easy_functions import easyDict_append, easyDict_update
from easy_functions import process_kwargs
from easy_functions import required_args_add, required_args_remove
from easy_functions import sensitive_var_site_group
from easy_functions import validate_args
import pkg_resources
import re
import validating

aci_template_path = pkg_resources.resource_filename('class_admin', 'templates/')

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
class admin(object):
    def __init__(self, type):
        self.type = type

    #======================================================
    # Function - Authentication
    #======================================================
    def auth(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.Authentication']['allOf'][1]['properties']
        
        # Check for Variable values that could change required arguments
        if 'console_realm' in kwargs:
            if not kwargs['console_realm'] == 'local':
                jsonData = required_args_add(['console_login_domain'], jsonData)
        else:
            kwargs['console_realm'] == 'local'
        if 'default_realm' in kwargs:
            if not kwargs['default_realm'] == 'local':
                jsonData = required_args_add(['default_login_domain'], jsonData)
        else:
            kwargs['default_realm'] == 'local'
        
        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Reset jsonData
        if not kwargs['console_realm'] == 'local':
            jsonData = required_args_remove(['console_login_domain'], jsonData)
        if not kwargs['default_realm'] == 'local':
            jsonData = required_args_remove(['default_login_domain'], jsonData)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'authentication'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Configuration Backup - Export Policies
    #======================================================
    def export_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.exportPolicy']['allOf'][1]['properties']

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        Additions = {
            'configuration_export': [],
            'window_description': kwargs['description']
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - RADIUS Authentication
    #======================================================
    def radius(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'RADIUS_Server': '',
                         'Port': '',
                         'RADIUS_Secret': '',
                         'Authz_Proto': '',
                         'Timeout': '',
                         'Retry_Interval': '',
                         'Mgmt_EPG': '',
                         'Login_Domain': '',
                         'Domain_Order': ''}
        optional_args = {'Description': '',
                         'Domain_Descr': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.ip_address('RADIUS_Server', templateVars['RADIUS_Server'])
            validating.validator('login_domain', **kwargs)
            validating.number_check('Domain_Order', templateVars['Domain_Order'], 0, 17)
            validating.number_check('Port', templateVars['Port'], 1, 65535)
            validating.number_check('Retry_Interval', templateVars['Retry_Interval'], 1, 5)
            validating.sensitive_var('RADIUS_Secret', templateVars['RADIUS_Secret'])
            validating.timeout('Timeout', templateVars['Timeout'])
            validating.values('Authz_Proto', templateVars['Authz_Proto'], ['chap', 'mschap', 'pap'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg('Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['Description'] == None:
                validating.description('Description', templateVars['Description'])
            if not templateVars['Domain_Descr'] == None:
                validating.description('Domain_Descr', templateVars['Domain_Descr'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        if re.search(r'\.', templateVars['RADIUS_Server']):
            templateVars['RADIUS_Server_'] = templateVars['RADIUS_Server'].replace('.', '-')
        else:
            templateVars['RADIUS_Server_'] = templateVars['RADIUS_Server'].replace(':', '-')

        if not templateVars['RADIUS_Secret'] == None:
            x = templateVars['RADIUS_Secret'].split('r')
            key_number = x[1]
            templateVars['sensitive_var'] = 'RADIUS_Secret%s' % (key_number)

    #======================================================
    # Function - Configuration Backup  - Remote Host
    #======================================================
    def remote_host(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.remoteHost']['allOf'][1]['properties']

        # Check for Variable values that could change required arguments
        if 'authentication_type' in kwargs:
            if kwargs['authentication_type'] == 'usePassword':
                jsonData = required_args_add(['username'], jsonData)
        else:
            kwargs['authentication_type'] == 'usePassword'
            jsonData = required_args_add(['username'], jsonData)

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars['jsonData'] = jsonData
        if templateVars['authentication_type'] == 'usePassword':
            # Check if the Password is in the Environment.  If not Add it.
            templateVars["Variable"] = f'remote_password_{kwargs["password"]}'
            sensitive_var_site_group(**templateVars)
        else:
            # Check if the SSH Key/Passphrase is in the Environment.  If not Add it.
            templateVars["Variable"] = 'ssh_key_contents'
            sensitive_var_site_group(**templateVars)
            templateVars["Variable"] = 'ssh_key_passphrase'
            sensitive_var_site_group(**templateVars)

        # Reset jsonData
        if kwargs['authentication_type'] == 'usePassword':
            jsonData = required_args_remove(['username'], jsonData)
        
        # Add Dictionary to Policy
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        templateVars['data_subtype'] = 'configuration_export'
        templateVars['policy_name'] = kwargs['scheduler_name']
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Global Security Settings
    #======================================================
    def security(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.globalSecurity']['allOf'][1]['properties']

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'global_security'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - TACACS+ Authentication
    #======================================================
    def tacacs(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'TACACS_Server': '',
                         'Port': '',
                         'TACACS_Secret': '',
                         'Auth_Proto': '',
                         'Timeout': '',
                         'Retry_Interval': '',
                         'Mgmt_EPG': '',
                         'Login_Domain': '',
                         'Domain_Order': '',
                         'Acct_DestGrp_Name': ''}
        optional_args = {'Description': '',
                         'Domain_Descr': '',
                         'Login_Domain_Descr': ''}

        # Temporarily Move the Provider Description
        tacacs_descr = kwargs['Description']

        ws_admin = kwargs['wb']['Admin']
        rows = ws_admin.max_row
        row_bundle = ''
        func = 'login_domain'
        count = countKeys(ws_admin, func)
        var_dict = findVars(ws_admin, func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('Name') == kwargs.get('Policy_Group'):
                row_bundle = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}
                break

        kwargs['Login_Domain_Descr'] = kwargs['Description']
        kwargs['Description'] = tacacs_descr

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.ip_address('TACACS_Server', templateVars['TACACS_Server'])
            validating.validator('login_domain', **kwargs)
            validating.number_check('Domain_Order', templateVars['Domain_Order'], 0, 17)
            validating.number_check('Port', templateVars['Port'], 1, 65535)
            validating.number_check('Retry_Interval', templateVars['Retry_Interval'], 1, 5)
            validating.sensitive_var('TACACS_Secret', templateVars['TACACS_Secret'])
            validating.timeout('Timeout', templateVars['Timeout'])
            validating.values('Auth_Proto', templateVars['Auth_Proto'], ['chap', 'mschap', 'pap'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg('Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['Description'] == None:
                validating.description('Description', templateVars['Description'])
            if not templateVars['Domain_Descr'] == None:
                validating.description('Domain_Descr', templateVars['Domain_Descr'])
            if not templateVars['Login_Domain_Descr'] == None:
                validating.description('Login_Domain_Descr', templateVars['Login_Domain_Descr'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)
