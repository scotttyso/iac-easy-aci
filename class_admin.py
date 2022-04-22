#!/usr/bin/env python3

import jinja2
import os
import pkg_resources
import re
import validating
from class_terraform import terraform_cloud
from easy_functions import countKeys, findVars
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group
from easy_functions import write_to_site
from easy_functions import write_to_template
from openpyxl import load_workbook

aci_template_path = pkg_resources.resource_filename('class_admin', 'templates/')

# Exception Classes
class InsufficientArgs(Exception):
    pass

class ErrException(Exception):
    pass

class InvalidArg(Exception):
    pass

class LoginFailed(Exception):
    pass

# Terraform ACI Provider - Admin Policies
# Class must be instantiated with Variables
class admin(object):
    def __init__(self, type):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(aci_template_path + '%s/') % (type))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)
        self.type = type

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def export_policy(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'Scheduler_Name': '',
                         'Days': '',
                         'Backup_Hour': '',
                         'Backup_Minute': '',
                         'Concurrent_Capacity': '',
                         'Export_Name': '',
                         'Format': '',
                         'Start_Now': '',
                         'Snapshot': '',
                         'Remote_Host': '',
                         'Encryption_Key': ''}
        optional_args = {'Scheduler_Descr': '',
                         'Export_Descr': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.days(row_num, ws, 'Days', templateVars['Days'])
            validating.number_check(row_num, ws, 'Backup_Hour', templateVars['Backup_Hour'], 0, 23)
            validating.number_check(row_num, ws, 'Backup_Minute', templateVars['Backup_Minute'], 0, 59)
            validating.values(row_num, ws, 'Concurrent_Capacity', templateVars['Concurrent_Capacity'], ['unlimited'])
            validating.values(row_num, ws, 'Format', templateVars['Format'], ['json', 'xml'])
            validating.values(row_num, ws, 'Start_Now', templateVars['Start_Now'], ['triggered', 'untriggered'])
            validating.values(row_num, ws, 'Snapshot', templateVars['Snapshot'], ['no', 'yes'])
            validating.sensitive_var(row_num, ws, 'Encryption_Key', templateVars['Encryption_Key'])
            if not templateVars['Scheduler_Descr'] == None:
                validating.description(row_num, ws, 'Scheduler_Descr', templateVars['Scheduler_Descr'])
            if not templateVars['Export_Descr'] == None:
                validating.description(row_num, ws, 'Export_Descr', templateVars['Export_Descr'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        if re.search(r'\.', templateVars['Remote_Host']):
            templateVars['Remote_Host_'] = templateVars['Remote_Host'].replace('.', '-')
        else:
            templateVars['Remote_Host_'] = templateVars['Remote_Host'].replace(':', '-')

        if not templateVars['Encryption_Key'] == None:
            x = templateVars['Encryption_Key'].split('r')
            key_number = x[1]
            templateVars['sensitive_var'] = 'Encryption_Key%s' % (key_number)

        # Define the Template Source
        template_file = "global_key.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Global_Key.tf'
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "variables.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Create Variables File for the Sensitive Variables
        dest_file = 'variable_%s.tf' % (templateVars['sensitive_var'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "export_policy.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Configuration_Export_Policy_%s.tf' % (templateVars['Scheduler_Name'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def login_domain(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'Login_Domain': '',
                         'Realm_Type': ''}
        optional_args = {'Description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.name_complexity(row_num, ws, 'Login_Domain', templateVars['Login_Domain'])
            validating.values(row_num, ws, 'Realm_Type', templateVars['Realm_Type'], ['RADIUS', 'TACACS'])
            if not templateVars['Description'] == None:
                validating.description(row_num, ws, 'Description', templateVars['Description'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        # Define the Template Source
        template_file = "Login_Domain_%s.jinja2" % (templateVars['Realm_Type'])
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Login_Domain_%s_%s.tf' % (templateVars['Realm_Type'], templateVars['Login_Domain'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def maint_group(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'MG_Name': '',
                         'Admin_State': '',
                         'Admin_Notify': '',
                         'Graceful': '',
                         'Ignore_Compatability': '',
                         'Run_Mode': '',
                         'SW_Version': '',
                         'Ver_Check_Override': '',
                         'FW_Type': '',
                         'MG_Type': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.name_rule(row_num, ws, 'MG_Name', templateVars['MG_Name'])
            validating.sw_version(row_num, ws, 'SW_Version', templateVars['SW_Version'])
            validating.values(row_num, ws, 'Admin_State', templateVars['Admin_State'], ['triggered', 'untriggered'])
            validating.values(row_num, ws, 'Admin_Notify', templateVars['Admin_Notify'], ['notifyAlwaysBetweenSets', 'notifyNever', 'notifyOnlyOnFailures'])
            validating.values(row_num, ws, 'Graceful', templateVars['Graceful'], ['no', 'yes'])
            validating.values(row_num, ws, 'Ignore_Compatability',templateVars['Ignore_Compatability'], ['no', 'yes'])
            validating.values(row_num, ws, 'Run_Mode', templateVars['Run_Mode'], ['pauseAlwaysBetweenSets', 'pauseNever', 'pauseOnlyOnFailures'])
            validating.values(row_num, ws, 'Ver_Check_Override', templateVars['Ver_Check_Override'], ['trigger', 'trigger-immediate', 'triggered', 'untriggered'])
            validating.values(row_num, ws, 'MG_Type', templateVars['MG_Type'], ['ALL', 'range'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        # Define the Template Source
        template_file = "maintenance_group.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Maintenance_Group_%s.tf' % (templateVars['MG_Name'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def radius(self, wb, ws, row_num, **kwargs):
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
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.ip_address(row_num, ws, 'RADIUS_Server', templateVars['RADIUS_Server'])
            validating.name_complexity(row_num, ws, 'Login_Domain', templateVars['Login_Domain'])
            validating.number_check(row_num, ws, 'Domain_Order', templateVars['Domain_Order'], 0, 17)
            validating.number_check(row_num, ws, 'Port', templateVars['Port'], 1, 65535)
            validating.number_check(row_num, ws, 'Retry_Interval', templateVars['Retry_Interval'], 1, 5)
            validating.sensitive_var(row_num, ws, 'RADIUS_Secret', templateVars['RADIUS_Secret'])
            validating.timeout(row_num, ws, 'Timeout', templateVars['Timeout'])
            validating.values(row_num, ws, 'Authz_Proto', templateVars['Authz_Proto'], ['chap', 'mschap', 'pap'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['Description'] == None:
                validating.description(row_num, ws, 'Description', templateVars['Description'])
            if not templateVars['Domain_Descr'] == None:
                validating.description(row_num, ws, 'Domain_Descr', templateVars['Domain_Descr'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        if re.search(r'\.', templateVars['RADIUS_Server']):
            templateVars['RADIUS_Server_'] = templateVars['RADIUS_Server'].replace('.', '-')
        else:
            templateVars['RADIUS_Server_'] = templateVars['RADIUS_Server'].replace(':', '-')

        if not templateVars['RADIUS_Secret'] == None:
            x = templateVars['RADIUS_Secret'].split('r')
            key_number = x[1]
            templateVars['sensitive_var'] = 'RADIUS_Secret%s' % (key_number)

        # Define the Template Source
        template_file = "radius.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'RADIUS_Provider_%s.tf' % (templateVars['RADIUS_Server_'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "variables.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Create Variables File for the Sensitive Variables
        dest_file = 'variable_%s.tf' % (templateVars['sensitive_var'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)
        sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def realm(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'Auth_Realm': '',
                         'Domain_Type': ''}
        optional_args = {'Login_Domain': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.login_type(row_num, ws, 'Auth_Realm', templateVars['Auth_Realm'], 'Domain_Type', templateVars['Domain_Type'])
            if not templateVars['Domain_Type'] == 'local':
                validating.name_complexity(row_num, ws, 'Login_Domain', templateVars['Login_Domain'])
            validating.values(row_num, ws, 'Auth_Realm', templateVars['Auth_Realm'], ['console', 'default'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        if templateVars['Auth_Realm'] == 'console':
            templateVars['child_class'] = 'aaaConsoleAuth'
        elif templateVars['Auth_Realm'] == 'default':
            templateVars['child_class'] = 'aaaDefaultAuth'

        # Define the Template Source
        template_file = "realm.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        if templateVars['Auth_Realm'] == 'console':
            dest_file = 'REALM_Console.tf'
        else:
            dest_file = 'REALM_default.tf'
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def remote_host(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'Remote_Host': '',
                         'Mgmt_EPG': '',
                         'Protocol': '',
                         'Remote_Path': '',
                         'Port': '',
                         'Auth_Type': '',
                         'Pwd_or_SSHPhrase': ''}
        optional_args = {'Description': '',
                         'Backup_User': '',
                         'SSH_Key': '',
                         'Description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])

            if re.match(r'\:', templateVars['Remote_Host']):
                validating.ip_address(row_num, ws, 'Remote_Host', templateVars['Remote_Host'])
            elif re.match('[a-z]', templateVars['Remote_Host'], re.IGNORECASE):
                validating.dns_name(row_num, ws, 'Remote_Host', templateVars['Remote_Host'])
            else:
                validating.ip_address(row_num, ws, 'Remote_Host', templateVars['Remote_Host'])

            validating.sensitive_var(row_num, ws, 'Pwd_or_SSHPhrase', templateVars['Pwd_or_SSHPhrase'])
            if templateVars['Auth_Type'] == 'password':
                validating.sensitive_var(row_num, ws, 'Backup_User', templateVars['Backup_User'])
            else:
                validating.sensitive_var(row_num, ws, 'SSH_Key', templateVars['SSH_Key'])

            if not templateVars['Description'] == None:
                validating.description(row_num, ws, 'Description', templateVars['Description'])

            validating.number_check(row_num, ws, 'Port', templateVars['Port'], 1, 65535)
            validating.values(row_num, ws, 'Auth_Type', templateVars['Auth_Type'], ['password', 'ssh-key'])
            validating.values(row_num, ws, 'Protocol', templateVars['Protocol'], ['ftp', 'scp', 'sftp'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        if re.search(':', templateVars['Remote_Host']):
            templateVars['Remote_Host_'] = templateVars['Remote_Host'].replace(':', '-')
        else:
            templateVars['Remote_Host_'] = templateVars['Remote_Host'].replace('.', '-')
        if templateVars['Auth_Type'] == 'password':
            templateVars['Auth_Type'] = 'usePassword'
        elif templateVars['Auth_Type'] == 'ssh-key':
            templateVars['Auth_Type'] = 'useSshKeyContents'

        if not templateVars['Pwd_or_SSHPhrase'] == None:
            x = templateVars['Pwd_or_SSHPhrase'].split('r')
            key_number = x[1]
            templateVars['sensitive_var1'] = 'Pwd_or_SSHPhrase%s' % (key_number)

        if templateVars['Auth_Type'] == 'usePassword':
            if not templateVars['Backup_User'] == None:
                x = templateVars['Backup_User'].split('r')
                key_number = x[1]
                templateVars['sensitive_var2'] = 'Backup_User%s' % (key_number)
        else:
            if not templateVars['SSH_Key'] == None:
                x = templateVars['SSH_Key'].split('r')
                key_number = x[1]
                templateVars['sensitive_var2'] = 'SSH_Key%s' % (key_number)

        # Define the Template Source
        template_file = "remote_host.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Remote_Location_%s.tf' % (templateVars['Remote_Host_'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "variables.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Create Variables File for the Sensitive Variables
        templateVars['sensitive_var'] = templateVars['sensitive_var1']
        dest_file = 'variable_%s.tf' % (templateVars['sensitive_var1'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

        templateVars['sensitive_var'] = templateVars['sensitive_var2']
        dest_file = 'variable_%s.tf' % (templateVars['sensitive_var2'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def security(self, wb, ws, row_num, **kwargs):
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
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.number_check(row_num, ws, 'Expiration_Warn', templateVars['Expiration_Warn'], 0, 30)
            validating.number_check(row_num, ws, 'Passwd_Intv', templateVars['Passwd_Intv'], 0, 745)
            validating.number_check(row_num, ws, 'Number_Allowed', templateVars['Number_Allowed'], 0, 10)
            validating.number_check(row_num, ws, 'Passwd_Store', templateVars['Passwd_Store'], 0, 15)
            validating.number_check(row_num, ws, 'Failed_Attempts', templateVars['Failed_Attempts'], 1, 15)
            validating.number_check(row_num, ws, 'Time_Period', templateVars['Time_Period'], 1, 720)
            validating.number_check(row_num, ws, 'Dur_Lockout', templateVars['Dur_Lockout'], 1, 1440)
            validating.number_check(row_num, ws, 'Token_Timeout', templateVars['Token_Timeout'], 300, 9600)
            validating.number_check(row_num, ws, 'Maximum_Valid', templateVars['Maximum_Valid'], 0, 24)
            validating.number_check(row_num, ws, 'Web_Timeout', templateVars['Web_Timeout'], 60, 65525)
            validating.values(row_num, ws, 'Enforce_Intv', templateVars['Enforce_Intv'], ['disable', 'enable'])
            validating.values(row_num, ws, 'Lockout', templateVars['Lockout'], ['disable', 'enable'])
            validating.values(row_num, ws, 'Passwd_Strength', templateVars['Passwd_Strength'], ['no', 'yes'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        # Define the Template Source
        template_file = "security.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'Global_Security.tf'
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def tacacs(self, wb, ws, row_num, **kwargs):
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

        ws_admin = wb['Admin']
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
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.ip_address(row_num, ws, 'TACACS_Server', templateVars['TACACS_Server'])
            validating.name_complexity(row_num, ws, 'Login_Domain', templateVars['Login_Domain'])
            validating.number_check(row_num, ws, 'Domain_Order', templateVars['Domain_Order'], 0, 17)
            validating.number_check(row_num, ws, 'Port', templateVars['Port'], 1, 65535)
            validating.number_check(row_num, ws, 'Retry_Interval', templateVars['Retry_Interval'], 1, 5)
            validating.sensitive_var(row_num, ws, 'TACACS_Secret', templateVars['TACACS_Secret'])
            validating.timeout(row_num, ws, 'Timeout', templateVars['Timeout'])
            validating.values(row_num, ws, 'Auth_Proto', templateVars['Auth_Proto'], ['chap', 'mschap', 'pap'])
            templateVars['Mgmt_EPG'] = validating.mgmt_epg(row_num, ws, 'Mgmt_EPG', templateVars['Mgmt_EPG'])
            if not templateVars['Description'] == None:
                validating.description(row_num, ws, 'Description', templateVars['Description'])
            if not templateVars['Domain_Descr'] == None:
                validating.description(row_num, ws, 'Domain_Descr', templateVars['Domain_Descr'])
            if not templateVars['Login_Domain_Descr'] == None:
                validating.description(row_num, ws, 'Login_Domain_Descr', templateVars['Login_Domain_Descr'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        if re.search(r'\.', templateVars['TACACS_Server']):
            templateVars['TACACS_Server_'] = templateVars['TACACS_Server'].replace('.', '-')
        else:
            templateVars['TACACS_Server_'] = templateVars['TACACS_Server'].replace(':', '-')

        if not templateVars['TACACS_Secret'] == None:
            x = templateVars['TACACS_Secret'].split('r')
            key_number = x[1]
            templateVars['sensitive_var'] = 'TACACS_Secret%s' % (key_number)

        # Define the Template Source
        template_file = "tacacs.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'TACACS_Provider_%s.tf' % (templateVars['TACACS_Server_'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

        # Define the Template Source
        template_file = "variables.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Create Variables File for the Sensitive Variables
        dest_file = 'variable_%s.tf' % (templateVars['sensitive_var'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)
        sensitive_var_site_group(wb, ws, row_num, dest_dir, dest_file, template, **templateVars)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def tacacs_acct(self, wb, ws, row_num, **kwargs):
        # Dicts for required and optional args
        required_args = {'Site_Group': '',
                         'Acct_DestGrp_Name': '',
                         'Acct_SrcGrp_Name': ''}
        optional_args = {'Description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group(row_num, ws, 'Site_Group', templateVars['Site_Group'])
            validating.name_rule(row_num, ws, 'Acct_DestGrp_Name', templateVars['Acct_DestGrp_Name'])
            validating.name_rule(row_num, ws, 'Acct_SrcGrp_Name', templateVars['Acct_SrcGrp_Name'])
            if not templateVars['Description'] == None:
                validating.description(row_num, ws, 'Description', templateVars['Description'])
        except Exception as err:
            Error_Return = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws, row_num)
            raise ErrException(Error_Return)

        # Define the Template Source
        template_file = "tacacs_accounting.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'TACACS_Accounting_DestGrp_%s.tf' % (templateVars['Acct_DestGrp_Name'])
        dest_dir = 'Admin'
        write_to_site(wb, ws, row_num, 'w', dest_dir, dest_file, template, **templateVars)

