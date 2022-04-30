#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from easy_functions import countKeys, findVars
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group
from easy_functions import update_easyDict
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
    # Function - Configuration Backup - Export Policies
    #======================================================
    def export_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.exportPolicy']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.number_check('max_snapshot_count', jsonData, **kwargs)
            validating.number_check('scheduled_hour', jsonData, **kwargs)
            validating.number_check('scheduled_minute', jsonData, **kwargs)
            validating.site_group('site_group', **kwargs)
            validating.validator('name', **kwargs)
            validating.values('format', jsonData, **kwargs)
            validating.values('scheduled_days', jsonData, **kwargs)
            validating.values('snapshot', jsonData, **kwargs)
            validating.values('start_now', jsonData, **kwargs)
            if not templateVars['description'] == None:
                validating.validator('description', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        Additions = {
            'window_description': kwargs['description']
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
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
    # Function - Authentication Realms
    #======================================================
    def realm(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'Auth_Realm': '',
                         'Domain_Type': ''}
        optional_args = {'Login_Domain': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.login_type('Auth_Realm', templateVars['Auth_Realm'], 'Domain_Type', templateVars['Domain_Type'])
            if not templateVars['Domain_Type'] == 'local':
                validating.validator('login_domain', **kwargs)
            validating.values('Auth_Realm', templateVars['Auth_Realm'], ['console', 'default'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        if templateVars['Auth_Realm'] == 'console':
            templateVars['child_class'] = 'aaaConsoleAuth'
        elif templateVars['Auth_Realm'] == 'default':
            templateVars['child_class'] = 'aaaDefaultAuth'

    #======================================================
    # Function - Configuration Backup  - Remote Host
    #======================================================
    def remote_host(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.remoteHost']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.number_check('remote_port', jsonData, **kwargs)
            validating.site_group('site_group', **kwargs)
            validating.validator('management_epg', **kwargs)
            validating.values('authentication_type', jsonData, **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            validating.values('protocol', jsonData, **kwargs)
            if templateVars['authentication_type'] == 'usePassword':
                validating.validator('username', **kwargs)
            if not templateVars['description'] == None:
                validating.validator('description', **kwargs)
            count = 1
            for hostname in kwargs['remote_hosts'].split(','):
                kwargs[f'remote_host_{count}'] = hostname
                if ':' in hostname:
                    validating.ip_address(f'remote_host_{count}', **kwargs)
                elif re.search('[a-z]', hostname, re.IGNORECASE):
                    validating.dns_name(f'remote_host_{count}', **kwargs)
                else:
                    validating.ip_address(f'remote_host_{count}', **kwargs)
                kwargs.pop(f'remote_host_{count}')
                count += 1
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

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

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Security Settings
    #======================================================
    def security(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.globalSecurity']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'Passwd_Strength': '',
                         'Enforce_Intv': '',
                         'Expiration_Warn': '',
                         'Passwd_Intv': '',
                         'Number_Allowed': '',
                         'Passwd_Store': '',
                         'Lockout': '',
                         'Failed_Attempts': '',
                         'Time_Period': '',
                         'Dur_Lockout': '',
                         'Token_Timeout': '',
                         'Maximum_Valid': '',
                         'Web_Timeout': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Arguments
            validating.number_check('lockout_duration', jsonData, **kwargs)
            validating.number_check('max_failed_attempts', jsonData, **kwargs)
            validating.number_check('max_failed_attempts_window', jsonData, **kwargs)
            validating.number_check('maximum_validity_period', jsonData, **kwargs)
            validating.number_check('password_change_interval', jsonData, **kwargs)
            validating.number_check('password_changes_within_interval', jsonData, **kwargs)
            validating.number_check('password_expiration_warn_time', jsonData, **kwargs)
            validating.number_check('user_passwords_to_store_count', jsonData, **kwargs)
            validating.number_check('web_session_idle_timeout', jsonData, **kwargs)
            validating.number_check('web_token_timeout', jsonData, **kwargs)
            validating.site_group('site_group', **kwargs)
            validating.values('enable_lockout', jsonData, **kwargs)
            validating.values('password_change_interval_enforce', jsonData, **kwargs)
            validating.values('password_strength_check', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'global_security'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
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
