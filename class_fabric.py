#!/usr/bin/env python3

from class_terraform import terraform_cloud
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group
from easy_functions import write_to_site
from easy_functions import update_easyDict
import re
import validating

# Exception Classes
class InsufficientArgs(Exception):
    pass

class ErrException(Exception):
    pass

class InvalidArg(Exception):
    pass

class LoginFailed(Exception):
    pass

class fabric(object):
    def __init__(self, type):
        self.type = type

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def date_time(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.DateandTime']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
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

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def dns_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.dnsProfiles']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.name_rule('management_epg', **kwargs)
            validating.number_check('preferred', jsonData, **kwargs)
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
                count += 1
            if not kwargs['default_domain'] in kwargs['domain_list']:
                kwargs['domain_list'].append(kwargs['default_domain'])
            count = 1
            for domain in kwargs['domain_list'].split(','):
                kwargs[f'domain_{count}'] = domain
                validating.domain(f'domain_{count}', **kwargs)
                count += 1
            if not templateVars['description'] == None:
                validating.description('description', **kwargs)
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


    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def ntp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Ntp']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Arguments
            validating.site_group('site_group', **kwargs)
            validating.name_rule('management_epg', **kwargs)
            validating.number_check('maximum_polling_interval', jsonData, **kwargs)
            validating.number_check('minimum_polling_interval', jsonData, **kwargs)
            validating.values('management_epg_type', jsonData, **kwargs)
            validating.values('preferred', jsonData, **kwargs)
            if ':' in templateVars['hostname']:
                validating.ip_address('hostname', **kwargs)
            elif re.search('[a-z]', templateVars['hostname'], re.IGNORECASE):
                validating.dns_name('hostname', **kwargs)
            else:
                validating.ip_address('hostname', **kwargs)
            if not templateVars['description'] == None:
                validating.description('description', **kwargs)
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

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def ntp_key(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.NtpKeys']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.number_check('key_id', jsonData, **kwargs)
            validating.values('authentication_type', jsonData, **kwargs)
            validating.values('trusted', jsonData, **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

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

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def sch_smtp_server(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.schSmtpServer']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.name_rule('management_epg', **kwargs)
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

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def smart_callhome(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartCallHome']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
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
                validating.description('contact_information', **kwargs)
            if not kwargs['phone_contact'] == None:
                validating.phone_number('phone_contact', **kwargs)
            if not kwargs['street_address'] == None:
                validating.description('street_address', **kwargs)
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

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def smart_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartCallHome']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
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

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_community(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Policy': '',
                         'SNMP_Community': ''}
        optional_args = {'description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.snmp_string('SNMP_Community', templateVars['SNMP_Community'])
            if not templateVars['description'] == None:
                validating.description('description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_clgrp(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Policy': '',
                         'Client_Group': '',
                         'Mgmt_EPG': ''}
        optional_args = {'description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.name_rule('Client_Group', templateVars['Client_Group'])
            validating.name_rule('SNMP_Policy', templateVars['SNMP_Policy'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg('Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['description'] == None:
                validating.description('description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartCallHome']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('admin_state', jsonData, **kwargs)
            validating.values('audit_logs', jsonData, **kwargs)
            validating.values('events', jsonData, **kwargs)
            validating.values('faults', jsonData, **kwargs)
            validating.values('session_logs', jsonData, **kwargs)
            if not kwargs['contact'] == None:
                validating.description('contact', **kwargs)
            if not kwargs['from_email'] == None:
                validating.email('from_email', **kwargs)
            if not kwargs['reply_to_email'] == None:
                validating.email('reply_to_email', **kwargs)
            if not kwargs['contact_information'] == None:
                validating.description('contact_information', **kwargs)
            if not kwargs['phone_contact'] == None:
                validating.phone_number('phone_contact', **kwargs)
            if not kwargs['street_address'] == None:
                validating.description('street_address', **kwargs)
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

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Policy': '',
                         'Admin_State': ''}
        optional_args = {'description': '',
                         'SNMP_Contact': '',
                         'SNMP_Location': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.name_rule('SNMP_Policy', templateVars['SNMP_Policy'])
            if not templateVars['description'] == None:
                validating.description('description', templateVars['description'])
            if not templateVars['SNMP_Contact'] == None:
                validating.description('SNMP_Contact', templateVars['SNMP_Contact'])
            if not templateVars['SNMP_Location'] == None:
                validating.description('SNMP_Location', templateVars['SNMP_Location'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_destinations(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Policy': '',
                         'SNMP_Trap_DG': '',
                         'Trap_Server': '',
                         'Destination_Port': '',
                         'Version': '',
                         'Community_or_Username': '',
                         'Security_Level': '',
                         'Mgmt_EPG': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        # Set noauth if v1 or v2c
        if re.search('(v1|v2c)', templateVars['Version']):
            templateVars['Security_Level'] = 'noauth'

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.ip_address('Trap_Server', templateVars['Trap_Server'])
            validating.number_check('Destination_Port', templateVars['Destination_Port'], 1, 65535)
            validating.values('Version', templateVars['Version'], ['v1', 'v2c', 'v3'])
            validating.values('Security_Level', templateVars['Security_Level'], ['auth', 'noauth', 'priv'])
            validating.snmp_string('Community_or_Username', templateVars['Community_or_Username'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg('Mgmt_EPG', templateVars['Mgmt_EPG'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_user(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Policy': '',
                         'SNMP_User': '',
                         'Authorization_Type': '',
                         'Authorization_Key': ''}
        optional_args = {'Privacy_Type': '',
                         'Privacy_Key': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            auth_type = templateVars['Authorization_Type']
            auth_key = templateVars['Authorization_Key']
            validating.snmp_auth(templateVars['Privacy_Type'], templateVars['Privacy_Key'], auth_type, auth_key)
            validating.snmp_string('SNMP_User', templateVars['SNMP_User'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def syslog(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'Dest_Grp_Name': '',
                         'Minimum_Level': '',
                         'Log_Format': '',
                         'Console': '',
                         'Console_Level': '',
                         'Local': '',
                         'Local_Level': '',
                         'Include_msec': '',
                         'Include_timezone': '',
                         'Audit': '',
                         'Events': '',
                         'Faults': '',
                         'Session': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.log_level('Minimum_Level', templateVars['Minimum_Level'])
            validating.log_level('Local_Level', templateVars['Local_Level'])
            validating.log_level('Console_Level', templateVars['Console_Level'])
            validating.name_rule('Dest_Grp_Name', templateVars['Dest_Grp_Name'])
            validating.values('Console', templateVars['Console'], ['disabled', 'enabled'])
            validating.values('Local', templateVars['Local'], ['disabled', 'enabled'])
            validating.values('Include_msec', templateVars['Include_msec'], ['no', 'yes'])
            validating.values('Include_timezone', templateVars['Include_timezone'], ['no', 'yes'])
            validating.values('Audit', templateVars['Audit'], ['no', 'yes'])
            validating.values('Events', templateVars['Events'], ['no', 'yes'])
            validating.values('Faults', templateVars['Faults'], ['no', 'yes'])
            validating.values('Session', templateVars['Session'], ['no', 'yes'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def syslog_destinations(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'Dest_Grp_Name': '',
                         'Syslog_Server': '',
                         'Port': '',
                         'Mgmt_EPG': '',
                         'Severity': '',
                         'Facility': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.ip_address('Syslog_Server', templateVars['Syslog_Server'])
            validating.log_level('Severity', templateVars['Severity'])
            validating.number_check('Port', templateVars['Port'], 1, 65535)
            validating.syslog_fac('Facility', templateVars['Facility'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg('Mgmt_EPG', templateVars['Mgmt_EPG'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)
