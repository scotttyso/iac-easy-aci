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
        # Set Locally Used Variables
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'administrative_state': '',
            'display_format': '',
            'master_mode': '',
            'offset_state': '',
            'server_state': '',
            'stratum_value': '',
            'time_zone': '',
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        # Get Variable Values
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.dateandTime']['allOf'][1]['properties']

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.values(row_num, ws, 'administrative_state', templateVars['administrative_state'], 
                jsonData['administrative_state']['enum'])
            validating.values(row_num, ws, 'display_format', templateVars['display_format'],
                jsonData['display_format']['enum'])
            validating.values(row_num, ws, 'master_mode', templateVars['master_mode'],
                jsonData['master_mode']['enum'])
            validating.values(row_num, ws, 'offset_state', templateVars['offset_state'],
                jsonData['offset_state']['enum'])
            validating.values(row_num, ws, 'server_state', templateVars['server_state'],
                jsonData['server_state']['enum'])
            validating.values(row_num, ws, 'stratum_value', templateVars['stratum_value'],
                jsonData['stratum_value']['enum'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        if templateVars['server_state'] == 'disabled':
            templateVars['master_mode'] = 'disabled'
        
        upDates = {
            'authentication_keys': [],
            'name':'default',
            'ntp_servers': [],
        }
        templateVars.update(upDates)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'date_and_time'
        kwargs['easyDict'] = update_easyDict(templateVars, **kwargs)
        return kwargs['easyDict']

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def dns(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'dns_profile': '',
            'DNS_Server': '',
            'Preferred': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.ip_address(row_num, ws, 'DNS_Server', templateVars['DNS_Server'])
            validating.values(row_num, ws, 'Preferred', templateVars['Preferred'], ['no', 'yes'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        if re.search(r'\.', templateVars['DNS_Server']):
            templateVars['DNS_Server_'] = templateVars['DNS_Server'].replace('.', '-')
        else:
            templateVars['DNS_Server_'] = templateVars['DNS_Server'].replace(':', '-')

        # Define the Template Source
        template_file = "dns.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'dns_profile_%s.tf' % (templateVars['dns_profile'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def dns_profile(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'name': '',
            'Mgmt_EPG': ''
        }
        optional_args = {
            'description': ''
        }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.name_rule(row_num, ws, 'Name', templateVars['Name'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['description'] == None:
                validating.description(row_num, ws, 'description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "dns_profile.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'dns_profile_%s.tf' % (templateVars['Name'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def domain(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'dns_profile': '',
            'default_domain': '',
            'domain': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.domain(row_num, ws, 'Domain', templateVars['Domain'])
            validating.values(row_num, ws, 'Default_Domain', templateVars['Default_Domain'], ['no', 'yes'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        templateVars['Domain_'] = templateVars['Domain'].replace('.', '-')

        # Define the Template Source
        template_file = "domain.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'dns_profile_%s.tf' % (templateVars['dns_profile'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def ntp(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'hostname': '',
            'management_epg': '',
            'management_epg_type': '',
            'maximum_polling_interval': '',
            'minimum_polling_interval': '',
            'preferred': '',
        }
        optional_args = {
            'description': '',
            'key_id': '',
        }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.name_rule(row_num, ws, 'management_epg', templateVars['management_epg'])
            validating.number_check(row_num, ws, 'maximum_polling_interval', templateVars['maximum_polling_interval'], 4, 16)
            validating.number_check(row_num, ws, 'minimum_polling_interval', templateVars['minimum_polling_interval'], 4, 16)
            validating.values(row_num, ws, 'management_epg_type', templateVars['management_epg_type'], ['inb', 'oob'])
            validating.values(row_num, ws, 'preferred', templateVars['preferred'], ['true', 'false'])
            if ':' in templateVars['hostname']:
                validating.ip_address(row_num, ws, 'hostname', templateVars['hostname'])
            elif re.search('[a-z]', templateVars['hostname'], re.IGNORECASE):
                validating.dns_name(row_num, ws, 'hostname', templateVars['hostname'])
            else:
                validating.ip_address(row_num, ws, 'hostname', templateVars['hostname'])
            if not templateVars['description'] == None:
                validating.description(row_num, ws, 'description', templateVars['description'])
            if not templateVars['key_id'] == None:
                validating.number_check(row_num, ws, 'key_id', templateVars['key_id'], 1, 65535)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        ntpArgs = {
            'description':kwargs['description'],
            'hostname':kwargs['hostname'],
            'key_id':kwargs['key_id'],
            'management_epg':kwargs['management_epg'],
            'management_epg_type':kwargs['management_epg_type'],
            'maximum_polling_interval':kwargs['maximum_polling_interval'],
            'minimum_polling_interval':kwargs['minimum_polling_interval'],
            'preferred':kwargs['preferred']
        }
        
        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['fabric']['date_and_time']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['ntp_servers'].append(ntpArgs)

        # Return Dictionary
        return kwargs['easyDict']

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def ntp_key(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'authentication_type': '',
            'key_id': '',
            'trusted': ''
        }
        optional_args = {}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.number_check(row_num, ws, 'key_id', templateVars['key_id'], 1, 65535)
            validating.values(row_num, ws, 'authentication_type', templateVars['authentication_type'], ['md5', 'sha1'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        templateVars["Variable"] = f'ntp_key_{kwargs["key_id"]}'
        sensitive_var_site_group(**templateVars)

        keyArgs = {
            'authentication_type':kwargs['authentication_type'],
            'key_id':kwargs['key_id'],
            'trusted':kwargs['trusted']
        }
        
        # Add Dictionary to easyDict
        count = 0
        for i in kwargs['easyDict']['fabric']['date_and_time']:
            for k, v in i.items():
                if k == kwargs['site_group']:
                    kwargs['easyDict']['fabric']['date_and_time'][count][kwargs['site_group']]['authentication_keys'].append(keyArgs)
            count =+ 1

        # Return Dictionary
        return kwargs['easyDict']

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def smart_callhome(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'DestGrp_Name': '',
                         'Admin_State': '',
                         'SMTP_Port': '',
                         'SMTP_Relay': '',
                         'Mgmt_EPG': '',
                         'From_Email': '',
                         'Reply_Email': '',
                         'To_Email': '',
                         'Contract_Id': '',
                         'Customer_Id': '',
                         'Site_Id': ''}
        optional_args = {'description': '',
                         'Phone_Number': '',
                         'Contact_Info': '',
                         'Street_Address': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            if re.match(r'\:', templateVars['SMTP_Relay']):
                validating.ip_address(row_num, ws, 'SMTP_Relay', templateVars['SMTP_Relay'])
            elif re.match(r'[a-zA-Z]', templateVars['SMTP_Relay']):
                validating.dns_name(row_num, ws, 'SMTP_Relay', templateVars['SMTP_Relay'])
            else:
                validating.ip_address(row_num, ws, 'SMTP_Relay', templateVars['SMTP_Relay'])
            validating.email(row_num, ws, 'From_Email', templateVars['From_Email'])
            validating.email(row_num, ws, 'Reply_Email', templateVars['Reply_Email'])
            validating.email(row_num, ws, 'To_Email', templateVars['To_Email'])
            validating.name_rule(row_num, ws, 'DestGrp_Name', templateVars['DestGrp_Name'])
            validating.values(row_num, ws, 'Admin_State', templateVars['Admin_State'], ['disabled', 'enabled'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['Contact_Info'] == None:
                validating.description(row_num, ws, 'Contact_Info', templateVars['Contact_Info'])
            if not templateVars['description'] == None:
                validating.description(row_num, ws, 'description', templateVars['description'])
            if not templateVars['Phone_Number'] == None:
                validating.phone(row_num, ws, 'Phone_Number', templateVars['Phone_Number'])
            if not templateVars['Street_Address'] == None:
                validating.description(row_num, ws, 'Street_Address', templateVars['Street_Address'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "smartcallhome_dg.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Smart_Callhome_%s.tf' % (templateVars['DestGrp_Name'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def sch_receiver(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'DestGrp_Name': '',
                         'Receiver_Name': '',
                         'Admin_State': '',
                         'Email': '',
                         'Format': '',
                         'RFC_Compliant': '',
                         'Audit': '',
                         'Events': '',
                         'Faults': '',
                         'Session': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.email(row_num, ws, 'Email', templateVars['Email'])
            validating.name_rule(row_num, ws, 'DestGrp_Name', templateVars['DestGrp_Name'])
            validating.name_rule(row_num, ws, 'Receiver_Name', templateVars['Receiver_Name'])
            validating.values(row_num, ws, 'Admin_State', templateVars['Admin_State'], ['disabled', 'enabled'])
            validating.values(row_num, ws, 'RFC_Compliant', templateVars['RFC_Compliant'], ['no', 'yes'])
            validating.values(row_num, ws, 'Audit', templateVars['Audit'], ['no', 'yes'])
            validating.values(row_num, ws, 'Events', templateVars['Events'], ['no', 'yes'])
            validating.values(row_num, ws, 'Faults', templateVars['Faults'], ['no', 'yes'])
            validating.values(row_num, ws, 'Session', templateVars['Session'], ['no', 'yes'])
            validating.values(row_num, ws, 'Format', templateVars['Format'], ['aml', 'short-txt', 'xml'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        incl_list = ''
        if not templateVars['Audit'] == 'no':
            incl_list = 'audit'
        if not templateVars['Events'] == 'no':
            if incl_list == '':
                incl_list = 'events'
            else:
                incl_list = incl_list + ',events'
        if not templateVars['Faults'] == 'no':
            if incl_list == '':
                incl_list = 'faults'
            else:
                incl_list = incl_list + ',faults'
        if not templateVars['Session'] == 'no':
            if incl_list == '':
                incl_list = 'session'
            else:
                incl_list = incl_list + ',session'
        templateVars['Included_Types'] = incl_list

        # Define the Template Source
        template_file = "smartcallhome_receiver.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Smart_Callhome_%s.tf' % (templateVars['DestGrp_Name'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "smartcallhome_source.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Smart_Callhome_%s.tf' % (templateVars['DestGrp_Name'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_client(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Policy': '',
                         'Client_Group': '',
                         'SNMP_Client': '',
                         'SNMP_Client_Name': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.ip_address(row_num, ws, 'SNMP_Client', templateVars['SNMP_Client'])
            validating.name_rule(row_num, ws, 'SNMP_Client_Name', templateVars['SNMP_Client_Name'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        if re.search(r'\.', templateVars['SNMP_Client']):
            templateVars['SNMP_Client_'] = templateVars['SNMP_Client'].replace('.', '-')
        else:
            templateVars['SNMP_Client_'] = templateVars['SNMP_Client'].replace(':', '-')

        # Define the Template Source
        template_file = "snmp_client.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Policy_%s.tf' % (templateVars['SNMP_Policy'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_community(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Policy': '',
                         'SNMP_Community': ''}
        optional_args = {'description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.snmp_string(row_num, ws, 'SNMP_Community', templateVars['SNMP_Community'])
            if not templateVars['description'] == None:
                validating.description(row_num, ws, 'description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        if not templateVars['SNMP_Community'] == None:
            x = templateVars['SNMP_Community'].split('r')
            key_number = x[1]
            templateVars['sensitive_var'] = 'SNMP_Community%s' % (key_number)

        # Define the Template Source
        template_file = "snmp_comm.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Policy_%s.tf' % (templateVars['SNMP_Policy'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

        # site_dict = ast.literal_eval(os.environ[Site_ID])

        # Define the Template Source
        template_file = "variables.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'variable_%s.tf' % (templateVars['sensitive_var'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_clgrp(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

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
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.name_rule(row_num, ws, 'Client_Group', templateVars['Client_Group'])
            validating.name_rule(row_num, ws, 'SNMP_Policy', templateVars['SNMP_Policy'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['description'] == None:
                validating.description(row_num, ws, 'description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "snmp_client_group.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Policy_%s.tf' % (templateVars['SNMP_Policy'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_policy(self, **kwargs):
        # Set Locally Used Variables
        wb = kwargs['wb']
        ws = kwargs['ws']
        row_num = kwargs['row_num']

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
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.name_rule(row_num, ws, 'SNMP_Policy', templateVars['SNMP_Policy'])
            if not templateVars['description'] == None:
                validating.description(row_num, ws, 'description', templateVars['description'])
            if not templateVars['SNMP_Contact'] == None:
                validating.description(row_num, ws, 'SNMP_Contact', templateVars['SNMP_Contact'])
            if not templateVars['SNMP_Location'] == None:
                validating.description(row_num, ws, 'SNMP_Location', templateVars['SNMP_Location'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "snmp_policy.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Policy_%s.tf' % (templateVars['SNMP_Policy'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_trap(self, wb, ws, row_num, **kwargs):
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
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.ip_address(row_num, ws, 'Trap_Server', templateVars['Trap_Server'])
            validating.number_check(row_num, ws, 'Destination_Port', templateVars['Destination_Port'], 1, 65535)
            validating.values(row_num, ws, 'Version', templateVars['Version'], ['v1', 'v2c', 'v3'])
            validating.values(row_num, ws, 'Security_Level', templateVars['Security_Level'], ['auth', 'noauth', 'priv'])
            validating.snmp_string(row_num, ws, 'Community_or_Username', templateVars['Community_or_Username'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        if re.search(r'\.', templateVars['Trap_Server']):
            templateVars['Trap_Server_'] = templateVars['Trap_Server'].replace('.', '-')
        else:
            templateVars['Trap_Server_'] = templateVars['Trap_Server'].replace(':', '-')

        if not templateVars['Community_or_Username'] == None:
            x = templateVars['Community_or_Username'].split('r')
            key_number = x[1]
            templateVars['sensitive_var'] = 'Community_or_Username%s' % (key_number)

        # Define the Template Source
        template_file = "snmp_trap.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Policy_%s.tf' % (templateVars['SNMP_Policy'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "snmp_trap_destgrp_reciever.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Trap_DestGrp_%s.tf' % (templateVars['SNMP_Trap_DG'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

        # site_dict = ast.literal_eval(os.environ[Site_ID])

        # Define the Template Source
        template_file = "variables.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'variable_%s.tf' % (templateVars['sensitive_var'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def snmp_user(self, wb, ws, row_num, **kwargs):
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
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            auth_type = templateVars['Authorization_Type']
            auth_key = templateVars['Authorization_Key']
            validating.snmp_auth(row_num, ws, templateVars['Privacy_Type'], templateVars['Privacy_Key'], auth_type, auth_key)
            validating.snmp_string(row_num, ws, 'SNMP_User', templateVars['SNMP_User'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        # Modify User Input of templateVars['Privacy_Type'] or templateVars['Authorization_Type'] to send to APIC
        if templateVars['Privacy_Type'] == 'none':
            templateVars['Privacy_Type'] = None
        if templateVars['Authorization_Type'] == 'md5':
            templateVars['Authorization_Type'] = None
        if templateVars['Authorization_Type'] == 'sha1':
            templateVars['Authorization_Type'] = 'hmac-sha1-96'

        if not templateVars['Privacy_Key'] == None:
            x = templateVars['Privacy_Key'].split('r')
            key_number = x[1]
            templateVars['sensitive_var1'] = 'Privacy_Key%s' % (key_number)

        if not templateVars['Authorization_Key'] == None:
            x = templateVars['Authorization_Key'].split('r')
            key_number = x[1]
            templateVars['sensitive_var2'] = 'Authorization_Key%s' % (key_number)

        # Define the Template Source
        template_file = "snmp_user.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Policy_%s.tf' % (templateVars['SNMP_Policy'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

        # site_dict = ast.literal_eval(os.environ[Site_ID])

        # Define the Template Source
        template_file = "variables.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Create Variables File for the Privacy Key & Authorization Key
        if not templateVars['Privacy_Key'] == None:
            dest_file = 'variable_%s.tf' % (templateVars['sensitive_var1'])
            dest_dir = 'Fabric'
            templateVars['sensitive_var'] = templateVars['sensitive_var1']
            write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

            sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

        if not templateVars['Authorization_Key'] == None:
            dest_file = 'variable_%s.tf' % (templateVars['sensitive_var2'])
            dest_dir = 'Fabric'
            templateVars['sensitive_var'] = templateVars['sensitive_var2']
            write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

            sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def syslog_dg(self, wb, ws, row_num, **kwargs):
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
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.log_level(row_num, ws, 'Minimum_Level', templateVars['Minimum_Level'])
            validating.log_level(row_num, ws, 'Local_Level', templateVars['Local_Level'])
            validating.log_level(row_num, ws, 'Console_Level', templateVars['Console_Level'])
            validating.name_rule(row_num, ws, 'Dest_Grp_Name', templateVars['Dest_Grp_Name'])
            validating.values(row_num, ws, 'Console', templateVars['Console'], ['disabled', 'enabled'])
            validating.values(row_num, ws, 'Local', templateVars['Local'], ['disabled', 'enabled'])
            validating.values(row_num, ws, 'Include_msec', templateVars['Include_msec'], ['no', 'yes'])
            validating.values(row_num, ws, 'Include_timezone', templateVars['Include_timezone'], ['no', 'yes'])
            validating.values(row_num, ws, 'Audit', templateVars['Audit'], ['no', 'yes'])
            validating.values(row_num, ws, 'Events', templateVars['Events'], ['no', 'yes'])
            validating.values(row_num, ws, 'Faults', templateVars['Faults'], ['no', 'yes'])
            validating.values(row_num, ws, 'Session', templateVars['Session'], ['no', 'yes'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)
        incl_list = ''
        if not templateVars['Audit'] == 'no':
            incl_list = 'audit'
        if not templateVars['Events'] == 'no':
            if incl_list == '':
                incl_list = 'events'
            else:
                incl_list = incl_list + ',events'
        if not templateVars['Faults'] == 'no':
            if incl_list == '':
                incl_list = 'faults'
            else:
                incl_list = incl_list + ',faults'
        if not templateVars['Session'] == 'no':
            if incl_list == '':
                incl_list = 'session'
            else:
                incl_list = incl_list + ',session'
        templateVars['Included_Types'] = incl_list

        # Define the Template Source
        template_file = "syslog_dg.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Syslog_DestGrp_%s.tf' % (templateVars['Dest_Grp_Name'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def syslog_rmt(self, wb, ws, row_num, **kwargs):
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
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.ip_address(row_num, ws, 'Syslog_Server', templateVars['Syslog_Server'])
            validating.log_level(row_num, ws, 'Severity', templateVars['Severity'])
            validating.number_check(row_num, ws, 'Port', templateVars['Port'], 1, 65535)
            validating.syslog_fac(row_num, ws, 'Facility', templateVars['Facility'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        if re.search(r'\.', templateVars['Syslog_Server']):
            templateVars['Syslog_Server_'] = templateVars['Syslog_Server'].replace('.', '-')
        else:
            templateVars['Syslog_Server_'] = templateVars['Syslog_Server'].replace(':', '-')

        # Define the Template Source
        template_file = "syslog_rmt.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Syslog_DestGrp_%s.tf' % (templateVars['Dest_Grp_Name'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def trap_groups(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'SNMP_Trap_DG': '',
                         'SNMP_Source': '',
                         'Audit': '',
                         'Events': '',
                         'Faults': '',
                         'Session': ''}
        optional_args = {'description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'site_group', templateVars['site_group'])
            validating.name_rule(row_num, ws, 'SNMP_Trap_DG', templateVars['SNMP_Trap_DG'])
            validating.name_rule(row_num, ws, 'SNMP_Source', templateVars['SNMP_Source'])
            validating.values(row_num, ws, 'Audit', templateVars['Audit'], ['no', 'yes'])
            validating.values(row_num, ws, 'Events', templateVars['Events'], ['no', 'yes'])
            validating.values(row_num, ws, 'Faults', templateVars['Faults'], ['no', 'yes'])
            validating.values(row_num, ws, 'Session', templateVars['Session'], ['no', 'yes'])
            if not templateVars['description'] == None:
                validating.description(row_num, ws, 'description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(errorReturn)

        incl_list = ''
        if not templateVars['Audit'] == 'no':
            incl_list = 'audit'
        if not templateVars['Events'] == 'no':
            if incl_list == '':
                incl_list = 'events'
            else:
                incl_list = incl_list + ',events'
        if not templateVars['Faults'] == 'no':
            if incl_list == '':
                incl_list = 'faults'
            else:
                incl_list = incl_list + ',faults'
        if not templateVars['Session'] == 'no':
            if incl_list == '':
                incl_list = 'session'
            else:
                incl_list = incl_list + ',session'
        templateVars['Included_Types'] = incl_list

        # Define the Template Source
        template_file = "snmp_trap_destgrp.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Trap_DestGrp_%s.tf' % (templateVars['SNMP_Trap_DG'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "snmp_trap_source.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'SNMP_Trap_DestGrp_%s.tf' % (templateVars['SNMP_Trap_DG'])
        dest_dir = 'Fabric'
        write_to_site(wb, ws, row_num, 'a+', dest_dir, dest_file, template, **templateVars)
