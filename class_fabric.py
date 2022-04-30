#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group
from easy_functions import update_easyDict
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

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('administrative_state', jsonData, **kwargs)
            validating.values('display_format', jsonData, **kwargs)
            validating.values('master_mode', jsonData, **kwargs)
            validating.values('offset_state', jsonData, **kwargs)
            validating.values('server_state', jsonData, **kwargs)
            validating.values('stratum_value', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

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
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - DNS Profiles
    #======================================================
    def dns_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.dnsProfiles']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.number_check('preferred', jsonData, **kwargs)
            validating.site_group('site_group', **kwargs)
            validating.validator('management_epg', **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            count = 1
            for hostname in kwargs['dns_providers'].split(','):
                kwargs[f'dns_provider_{count}'] = hostname
                if ':' in hostname:
                    validating.ip_address(f'dns_provider_{count}', **kwargs)
                elif re.search('[a-z]', hostname, re.IGNORECASE):
                    validating.dns_name(f'dns_provider_{count}', **kwargs)
                else:
                    validating.ip_address(f'dns_provider_{count}', **kwargs)
                kwargs.pop(f'dns_provider_{count}')
                count += 1
            if not kwargs['default_domain'] in kwargs['domain_list']:
                kwargs['domain_list'].append(kwargs['default_domain'])
            count = 1
            for domain in kwargs['domain_list'].split(','):
                kwargs[f'domain_{count}'] = domain
                validating.domain(f'domain_{count}', **kwargs)
                count += 1
            if not templateVars['description'] == None:
                validating.validator('description', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        Additions = {
            'name':'default',
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'dns_profile'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Date and Time Policy - NTP Servers
    #======================================================
    def ntp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Ntp']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.number_check('maximum_polling_interval', jsonData, **kwargs)
            validating.number_check('minimum_polling_interval', jsonData, **kwargs)
            validating.validator('management_epg', **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            validating.values('preferred', jsonData, **kwargs)
            if ':' in templateVars['hostname']:
                validating.ip_address('hostname', **kwargs)
            elif re.search('[a-z]', templateVars['hostname'], re.IGNORECASE):
                validating.dns_name('hostname', **kwargs)
            else:
                validating.ip_address('hostname', **kwargs)
            if not templateVars['description'] == None:
                validating.validator('description', **kwargs)
            if not templateVars['key_id'] == None:
                validating.number_check('key_id', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

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

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.number_check('key_id', jsonData, **kwargs)
            validating.values('authentication_type', jsonData, **kwargs)
            validating.values('trusted', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Check if the NTP Key is in the Environment.  If not Add it.
        templateVars["Variable"] = f'ntp_key_{kwargs["key_id"]}'
        sensitive_var_site_group(**templateVars)

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

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('admin_state', jsonData, **kwargs)
            validating.values('audit_logs', jsonData, **kwargs)
            validating.values('events', jsonData, **kwargs)
            validating.values('faults', jsonData, **kwargs)
            validating.values('session_logs', jsonData, **kwargs)
            if not kwargs['customer_contact_email'] == None:
                validating.email('customer_contact_email', **kwargs)
            if not kwargs['from_email'] == None:
                validating.email('from_email', **kwargs)
            if not kwargs['reply_to_email'] == None:
                validating.email('reply_to_email', **kwargs)
            if not kwargs['contact_information'] == None:
                validating.validator('contact_information', **kwargs)
            if not kwargs['phone_contact'] == None:
                validating.phone_number('phone_contact', **kwargs)
            if not kwargs['street_address'] == None:
                validating.validator('street_address', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        Additions = {
            'smtp_server': [],
            'name':'default',
            'smart_destinations': [],
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'smartcallhome'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Smart CallHome Policy - Smart Destinations
    #======================================================
    def smart_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartCallHome']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.email('email', **kwargs)
            validating.values('admin_state', jsonData, **kwargs)
            validating.values('rfc_compliant', jsonData, **kwargs)
            validating.values('format', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

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

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('management_epg', **kwargs)
            validating.number_check('port_number', jsonData, **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            validating.values('secure_smtp', jsonData, **kwargs)
            if ':' in kwargs['smtp_server']:
                validating.ip_address('smtp_server', **kwargs)
            elif re.search('[a-z]', kwargs['smtp_server'], re.IGNORECASE):
                validating.dns_name('smtp_server', **kwargs)
            else:
                validating.ip_address('smtp_server', **kwargs)
            if 'true' in kwargs['secure_smtp']:
                validating.not_empty('username', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Check if the Smart CallHome SMTP Password is in the Environment and if not add it.
        if 'true' in kwargs['secure_smtp']:
            templateVars["Variable"] = f'smtp_password'
            sensitive_var_site_group(**templateVars)

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

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('management_epg',  **kwargs)
            validating.validator('name',  **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
            count = 1
            for hostname in kwargs['clients'].split(','):
                kwargs[f'client_{count}'] = hostname
                if ':' in hostname:
                    validating.ip_address(f'client_{count}', **kwargs)
                elif re.search('[a-z]', hostname, re.IGNORECASE):
                    validating.dns_name(f'client_{count}', **kwargs)
                else:
                    validating.ip_address(f'client_{count}', **kwargs)
                kwargs.pop(f'client_{count}')
                count += 1
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['snmp_policies']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['snmp_client_groups'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - Communities
    #======================================================
    def snmp_community(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpCommunities']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            if not templateVars['description'] == None:
                validating.validator('description', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Check if the SNMP Community is in the Environment.  If not Add it.
        templateVars["Variable"] = f'snmp_community_{kwargs["community_variable"]}'
        sensitive_var_site_group(**templateVars)

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['snmp_policies']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['snmp_communities'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - SNMP Trap Destinations
    #======================================================
    def snmp_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpDestinations']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.number_check('port', jsonData, **kwargs)
            validating.site_group('site_group', **kwargs)
            validating.validator('management_epg',  **kwargs)
            validating.validator('name',  **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            validating.values('version', jsonData, **kwargs)
            if ':' in kwargs['host']:
                validating.ip_address('host', **kwargs)
            elif re.search('[a-z]', kwargs['host'], re.IGNORECASE):
                validating.dns_name('host', **kwargs)
            else:
                validating.ip_address('host', **kwargs)
            if re.fullmatch('(v1|v2c)', kwargs['version']):
                validating.number_check('community_variable', **kwargs)
                count = 0
                for i in kwargs['easyDict']['fabric']['snmp_policies']['snmp_communities']:
                    for k, v in i.items():
                        if int(v['community_variable']) == int(kwargs['community_variable']):
                            count += 1
                if not count == 1:
                    validating.error_snmp_community(kwargs['row_num'], kwargs['community_variable'])
            elif 'v3' in kwargs['version']:
                validating.values('v3_security_level', **kwargs)
                count = 0
                for i in kwargs['easyDict']['fabric']['snmp_policies']['snmp_users']:
                    for k, v in i.items():
                        if int(v['username']) == int(kwargs['username']):
                            count += 1
                if not count == 1:
                    validating.error_snmp_user(kwargs['row_num'], kwargs['username'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['snmp_policies']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['snmp_destinations'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy
    #======================================================
    def snmp_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpPolicy']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('admin_state', jsonData, **kwargs)
            validating.values('audit_logs', jsonData, **kwargs)
            validating.values('events', jsonData, **kwargs)
            validating.values('faults', jsonData, **kwargs)
            validating.values('session_logs', jsonData, **kwargs)
            if not kwargs['contact'] == None:
                validating.validator('contact', **kwargs)
            if not kwargs['location'] == None:
                validating.validator('location', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

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
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - SNMP Users
    #======================================================
    def snmp_user(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpUsers']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('username',  **kwargs)
            validating.values('authorization_type', jsonData, **kwargs)
            if not kwargs['authorization_key'] == None:
                validating.number_check('authorization_key', jsonData, **kwargs)
            if not kwargs['privacy_key'] == None:
                validating.number_check('privacy_key', jsonData, **kwargs)
            if kwargs['privacy_type'] == None:
                kwargs['privacy_type'] == 'none'
            else:
                validating.values('privacy_type', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Check if the Authorization and Privacy Keys are in the environment and if not add them.
        templateVars["Variable"] = f'snmp_authorization_key_{kwargs["authorization_key"]}'
        sensitive_var_site_group(**templateVars)
        if not kwargs['privacy_type'] == 'none':
            templateVars["Variable"] = f'snmp_privacy_key_{kwargs["privacy_key"]}'
            sensitive_var_site_group(**templateVars)

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['snmp_policies']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['users'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - Syslog Policy
    #======================================================
    def syslog(self, **kwargs):
       # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Syslog']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('admin_state', jsonData, **kwargs)
            validating.values('audit_logs', jsonData, **kwargs)
            validating.values('console_admin_state', jsonData, **kwargs)
            validating.values('console_severity', jsonData, **kwargs)
            validating.values('events', jsonData, **kwargs)
            validating.values('faults', jsonData, **kwargs)
            validating.values('format', jsonData, **kwargs)
            validating.values('local_admin_state', jsonData, **kwargs)
            validating.values('local_severity', jsonData, **kwargs)
            validating.values('session_logs', jsonData, **kwargs)
            validating.values('show_milliseconds_in_timestamp', jsonData, **kwargs)
            validating.values('show_time_zone_in_timestamp', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        Additions = {
            'name':'default',
            'remote_destinations': []
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'syslog'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Syslog Policy - Syslog Destinations
    #======================================================
    def syslog_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.syslogRemoteDestinations']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.number_check('port', jsonData, **kwargs)
            validating.site_group('site_group', **kwargs)
            validating.validator('management_epg',  **kwargs)
            validating.values('admin_state', jsonData, **kwargs)
            validating.values('forwarding_facility', jsonData, **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            validating.values('severity', jsonData, **kwargs)
            validating.values('transport', jsonData, **kwargs)
            if ':' in kwargs['host']:
                validating.ip_address('host', **kwargs)
            elif re.search('[a-z]', kwargs['host'], re.IGNORECASE):
                validating.dns_name('host', **kwargs)
            else:
                validating.ip_address('host', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['syslog']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['remote_destinations'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']
