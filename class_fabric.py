#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from easy_functions import easyDict_append, easyDict_update
from easy_functions import process_kwargs
from easy_functions import required_args_add, required_args_remove
from easy_functions import sensitive_var_site_group
from easy_functions import validate_args
import json
import re
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
class fabric(object):
    def __init__(self, type):
        self.type = type

    #======================================================
    # Function - Date and Time Policy
    #======================================================
    def date_time(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.DateandTime']['allOf'][1]['properties']

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # If NTP server state is disabled disable master mode as well.
        if templateVars['server_state'] == 'disabled':
            templateVars['master_mode'] = 'disabled'
        
        Additions = {
            'authentication_keys': [],
            'name':'default',
            'ntp_servers': [],
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'date_and_time'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - DNS Profiles
    #======================================================
    def dns_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.dnsProfiles']['allOf'][1]['properties']

        if not kwargs['domain_list'] == None:
            if ',' in kwargs['domain_list']:
                kwargs['domain_list'] = kwargs['domain_list'].split(',')
            else:
                kwargs['domain_list'] = [kwargs['domain_list']]
            if not kwargs['default_domain'] == None:
                if not kwargs['default_domain'] in kwargs['domain_list']:
                    kwargs['domain_list'].append(kwargs['default_domain'])

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
            'name':'default',
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'dns_profiles'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Date and Time Policy - NTP Servers
    #======================================================
    def ntp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Ntp']['allOf'][1]['properties']

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
        for items in kwargs['easyDict']['fabric']['date_and_time']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['ntp_servers'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - Date and Time Policy - NTP Keys
    #======================================================
    def ntp_key(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.NtpKeys']['allOf'][1]['properties']

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Check if the NTP Key is in the Environment.  If not Add it.
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'ntp_key_{kwargs["key_id"]}'
        sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['date_and_time']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['authentication_keys'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - Smart CallHome Policy
    #======================================================
    def smart_callhome(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartCallHome']['allOf'][1]['properties']

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
            'smtp_server': [],
            'name':'default',
            'smart_destinations': [],
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'smartcallhome'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Smart CallHome Policy - Smart Destinations
    #======================================================
    def smart_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartDestinations']['allOf'][1]['properties']

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
        for items in kwargs['easyDict']['fabric']['smartcallhome']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['smart_destinations'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - Smart CallHome Policy - SMTP Server
    #======================================================
    def smart_smtp_server(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartSmtpServer']['allOf'][1]['properties']

        # Check for Variable values that could change required arguments
        if 'secure_smtp' in kwargs:
            if 'true' in kwargs['secure_smtp']:
                jsonData = required_args_add(['username'], jsonData)
        else:
            kwargs['secure_smtp'] == 'false'

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Check if the Smart CallHome SMTP Password is in the Environment and if not add it.
        if 'true' in kwargs['secure_smtp']:
            templateVars['jsonData'] = jsonData
            templateVars["Variable"] = f'smtp_password'
            sensitive_var_site_group(**templateVars)
            templateVars.pop('jsonData')
            templateVars.pop('Variable')

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['smartcallhome']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['smtp_server'].append(templateVars)
        
        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - Client Groups
    #======================================================
    def snmp_clgrp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpClientGroups']['allOf'][1]['properties']

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'snmp_client_groups'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - Communities
    #======================================================
    def snmp_community(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpCommunities']['allOf'][1]['properties']

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Check if the SNMP Community is in the Environment.  If not Add it.
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'snmp_community_{kwargs["community_variable"]}'
        sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'snmp_communities'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - SNMP Trap Destinations
    #======================================================
    def snmp_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpDestinations']['allOf'][1]['properties']

        # Check for Variable values that could change required arguments
        if 'version' in kwargs:
            if re.fullmatch('(v1|v2c)', kwargs['version']):
                jsonData = required_args_add(['community_variable'], jsonData)
            elif 'v3' in kwargs['version']:
                jsonData = required_args_add(['username', 'v3_security_level'], jsonData)
        else:
            kwargs['version'] = 'v2c'

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Reset Arguments
        if re.fullmatch('(v1|v2c)', kwargs['version']):
            jsonData = required_args_remove(['community_variable'], jsonData)
        elif 'v3' in kwargs['version']:
            jsonData = required_args_remove(['username', 'v3_security_level'], jsonData)
        
        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'snmp_destinations'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy
    #======================================================
    def snmp_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpPolicy']['allOf'][1]['properties']

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
            'name':'default',
            'snmp_client_groups': [],
            'snmp_communities': [],
            'snmp_destinations': [],
            'users': [],
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - SNMP Users
    #======================================================
    def snmp_user(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpUsers']['allOf'][1]['properties']

        # Check for Variable values that could change required arguments
        if 'privacy_key' in kwargs:
            if not kwargs['privacy_key'] == 'none':
                jsonData = required_args_add(['privacy_key'], jsonData)
        else:
            kwargs['privacy_key'] = 'none'

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Check if the Authorization and Privacy Keys are in the environment and if not add them.
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'snmp_authorization_key_{kwargs["authorization_key"]}'
        sensitive_var_site_group(**templateVars)
        if not kwargs['privacy_type'] == 'none':
            templateVars["Variable"] = f'snmp_privacy_key_{kwargs["privacy_key"]}'
            sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Reset jsonData
        if not kwargs['privacy_key'] == 'none':
            jsonData = required_args_remove(['privacy_key'], jsonData)
        
        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'users'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Syslog Policy
    #======================================================
    def syslog(self, **kwargs):
       # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Syslog']['allOf'][1]['properties']

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
            'name':'default',
            'remote_destinations': []
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'syslog'
        kwargs['easyDict'] = easyDict_update(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Syslog Policy - Syslog Destinations
    #======================================================
    def syslog_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.syslogRemoteDestinations']['allOf'][1]['properties']

        try:
            # Validate User Input
            validate_args(jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'syslog'
        templateVars['data_subtype'] = 'remote_destinations'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']
