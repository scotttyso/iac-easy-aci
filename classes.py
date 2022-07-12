#!/usr/bin/env python3

#=============================================================================
# Source Modules
#=============================================================================
from collections import OrderedDict
from easy_functions import apic_get, apic_post, countKeys, findKeys, findVars, sensitive_var_value
from easy_functions import easyDict_append, easyDict_append_policy, easyDict_append_subtype
from easy_functions import interface_selector_workbook, process_kwargs
from easy_functions import required_args_add, required_args_remove
from easy_functions import sensitive_var_site_group, stdout_log, tfc_get, tfc_patch, tfc_post
from easy_functions import validate_args, varBoolLoop, variablesFromAPI, varStringLoop
from easy_functions import vlan_list_full, write_to_site
from openpyxl import load_workbook
from requests.api import delete
import jinja2
import json
import os
import pkg_resources
import platform
import re
import requests
import stdiomask
import sys
import validating
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Global options for debugging
print_payload = False
print_response_always = False
print_response_on_fail = True

# Log levels 0 = None, 1 = Class only, 2 = Line
log_level = 2

# Global path to main Template directory
template_path = pkg_resources.resource_filename('classes', 'templates/')

class LoginFailed(Exception):
    pass

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class access(object):
    def __init__(self, type):
        self.type = type

    #=============================================================================
    # Function - Domain - Layer 3
    #=============================================================================
    def domains_l3(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.domains.Layer3']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'domains_layer3'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Domains - Physical
    #=============================================================================
    def domains_phys(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.domains.Physical']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'domains_physical'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Global Policies - AAEP Profiles
    #=============================================================================
    def global_aaep(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.global.attachableAccessEntityProfile']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        domain_list = ['physical_domains', 'l3_domains', 'vmm_domains']
        for i in domain_list:
            if not templateVars[f'{i}'] == None:
                templateVars[f'{i}'] = templateVars[f'{i}'].split(',')
                    

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'global_attachable_access_entity_profiles'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Policy Groups - Interface Policies
    # Shared Policies with Access and Bundle Poicies Groups
    #=============================================================================
    def interface_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.interfacePolicies']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('interface_policy')
        policy_dict = {kwargs['interface_policy']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'access'
        policy_dict['data_type'] = 'interface_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Leaf Policy Group
    #=============================================================================
    def leaf_pg(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.switches.leafPolicyGroup']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'switches_leaf_policy_groups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Policy Group - Access
    #=============================================================================
    def pg_access(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.leafAccessPort']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        if not templateVars['netflow_monitor_policies'] == None:
            if ',' in templateVars['netflow_monitor_policies']:
                templateVars['netflow_monitor_policies'] = templateVars['netflow_monitor_policies'].split(',')

        # Attach the Interface Policy Additional Attributes
        if kwargs['easyDict']['access']['interface_policies'].get(templateVars['interface_policy']):
            templateVars.update(kwargs['easyDict']['access']['interface_policies'][templateVars['interface_policy']])
        else:
            validating.error_policy_not_found('interface_policy', **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'leaf_interfaces_policy_groups_access'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Policy Group - Breakout
    #=============================================================================
    def pg_breakout(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.leafBreakOut']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'leaf_interfaces_policy_groups_breakout'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Policy Group - VPC/Port-Channel
    #=============================================================================
    def pg_bundle(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.leafBundle']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        if not templateVars['netflow_monitor_policies'] == None:
            if ',' in templateVars['netflow_monitor_policies']:
                templateVars['netflow_monitor_policies'] = templateVars['netflow_monitor_policies'].split(',')

        # Attach the Interface Policy Additional Attributes
        if kwargs['easyDict']['access']['interface_policies'].get(templateVars['interface_policy']):
            templateVars.update(kwargs['easyDict']['access']['interface_policies'][templateVars['interface_policy']])
        else:
            validating.error_policy_not_found('interface_policy', **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'leaf_interfaces_policy_groups_bundle'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Policy Group - Spine
    #=============================================================================
    def pg_spine(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.spineAccessPort']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'spine_interface_policy_groups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - CDP
    #=============================================================================
    def pol_cdp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.cdpInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_cdp_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - Fibre Channel
    #=============================================================================
    def pol_fc(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.fibreChannelInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_fibre_channel_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - L2 Interfaces
    #=============================================================================
    def pol_l2(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.L2Interface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_l2_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - Link Level (Speed)
    #=============================================================================
    def pol_link_level(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.linkLevel']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_link_level'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - LLDP
    #=============================================================================
    def pol_lldp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.lldpInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_lldp_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - Mis-Cabling Protocol
    #=============================================================================
    def pol_mcp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.mcpInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_mcp_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - Port Channel
    #=============================================================================
    def pol_port_ch(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.PortChannel']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_port_channel'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - Port Security
    #=============================================================================
    def pol_port_sec(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.portSecurity']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_port_security'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Interface Policies - Spanning Tree
    #=============================================================================
    def pol_stp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.spanningTreeInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'policies_spanning_tree_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - VLAN Pools
    #=============================================================================
    def pools_vlan(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.pools.Vlan']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'pools_vlan'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Spine Policy Group
    #=============================================================================
    def spine_pg(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.switches.spinePolicyGroup']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'switches_spine_policy_groups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Virtual Networking - Controllers
    #=============================================================================
    def vmm_controllers(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.vmm.Controllers']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'virtual_networking'
        templateVars['data_subtype'] = 'controllers'
        templateVars['policy_name'] = kwargs['domain_name']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Virtual Networking - Credentials
    #=============================================================================
    def vmm_creds(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.vmm.Credentials']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'virtual_networking'
        templateVars['data_subtype'] = 'credentials'

        # Check Environment for VMM Credentials Password
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'vmm_password_{kwargs["password"]}'
        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['policy_name'] = kwargs['domain_name']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Virtual Networking - Domains
    #=============================================================================
    def vmm_domain(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.vmm.Domains']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Convert to Lists
        if not templateVars["uplink_names"] == None:
            if ',' in templateVars["uplink_names"]:
                templateVars["uplink_names"] = templateVars["uplink_names"].split(',')
        else:
            templateVars["uplink_names"] = []

        upDating = {
            'controllers':[],
            'credentials':[],
            'enhanced_lag_policy':[],
            'domain':[templateVars],
            'name':templateVars['name'],
            'vswitch_policy':[]
        }
        templateVars = upDating
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'virtual_networking'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Virtual Networking - Controllers
    #=============================================================================
    def vmm_elagp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.vmm.enhancedLag']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'virtual_networking'
        templateVars['data_subtype'] = 'enhanced_lag_policy'
        templateVars['policy_name'] = kwargs['domain_name']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Virtual Networking - Controllers
    #=============================================================================
    def vmm_vswitch(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.vmm.vswitchPolicy']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'virtual_networking'
        templateVars['data_subtype'] = 'vswitch_policy'
        templateVars['policy_name'] = kwargs['domain_name']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class admin(object):
    def __init__(self, type):
        self.type = type

    #=============================================================================
    # Function - Authentication
    #=============================================================================
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
        
        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Configuration Backup - Export Policies
    #=============================================================================
    def export_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.exportPolicy']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars['schedule'] = {
            'days':templateVars['days'],
            'hour':templateVars['hour'],
            'minute':templateVars['minute']
        }
        templateVars.update({'configuration_export': []})
        remove_list = ['days', 'hour', 'minute']
        for i in remove_list:
            templateVars.pop(i)
    
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Global Security Settings
    #=============================================================================
    def mg_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.firmware.Policy']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['maintenance_groups'] = []

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'firmware'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Configuration Backup  - Remote Host
    #=============================================================================
    def maint_group(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.firmware.MaintenanceGroups']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Convert to Lists
        if ',' in templateVars["node_list"]:
            templateVars["node_list"] = templateVars["node_list"].split(',')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'firmware'
        templateVars['data_subtype'] = 'maintenance_groups'
        templateVars['policy_name'] = kwargs['maintenance_group_policy']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - RADIUS Authentication
    #=============================================================================
    def radius(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.Radius']['allOf'][1]['properties']
        
        # Check for Variable values that could change required arguments
        if 'server_monitoring' in kwargs:
            if kwargs['server_monitoring'] == 'enabled':
                jsonData = required_args_add(['monitoring_password', 'username'], jsonData)
        else:
            kwargs['server_monitoring'] == 'disabled'
        
        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'radius'

        # Check if the secert is in the Environment.  If not Add it.
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'radius_key_{kwargs["key"]}'
        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        if templateVars['server_monitoring'] == 'enabled':
            # Check if the Password is in the Environment.  If not Add it.
            templateVars["Variable"] = f'radius_monitoring_password_{kwargs["monitoring_password"]}'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Reset jsonData
        if not kwargs['server_monitoring'] == 'disabled':
            jsonData = required_args_remove(['monitoring_password', 'username'], jsonData)
        
        # Add Dictionary to easyDict
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Configuration Backup  - Remote Host
    #=============================================================================
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

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        templateVars['data_subtype'] = 'configuration_export'
        
        # Check Environment for Sensitive Variables
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        if templateVars['authentication_type'] == 'usePassword':
            # Check if the Password is in the Environment.  If not Add it.
            templateVars["Variable"] = f'remote_password_{kwargs["password"]}'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        else:
            # Check if the SSH Key/Passphrase is in the Environment.  If not Add it.
            templateVars["Variable"] = 'ssh_key_contents'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
            templateVars["Variable"] = 'ssh_key_passphrase'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')


        # Reset jsonData
        if kwargs['authentication_type'] == 'usePassword':
            jsonData = required_args_remove(['username'], jsonData)
        
        # Convert to Lists
        if ',' in templateVars["remote_hosts"]:
            templateVars["remote_hosts"] = templateVars["remote_hosts"].split(',')

        # Add Dictionary to Policy
        templateVars['policy_name'] = kwargs['scheduler_name']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Global Security Settings
    #=============================================================================
    def security(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.globalSecurity']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'security'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - TACACS+ Authentication
    #=============================================================================
    def tacacs(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.Tacacs']['allOf'][1]['properties']
        
        # Check for Variable values that could change required arguments
        if 'server_monitoring' in kwargs:
            if kwargs['server_monitoring'] == 'enabled':
                jsonData = required_args_add(['monitoring_password', 'username'], jsonData)
        else:
            kwargs['server_monitoring'] == 'disabled'
        
        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'tacacs'

        # Check if the secert is in the Environment.  If not Add it.
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'tacacs_key_{kwargs["key"]}'
        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        if templateVars['server_monitoring'] == 'enabled':
            # Check if the Password is in the Environment.  If not Add it.
            templateVars["Variable"] = f'tacacs_monitoring_password_{kwargs["monitoring_password"]}'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Reset jsonData
        if not kwargs['server_monitoring'] == 'disabled':
            jsonData = required_args_remove(['monitoring_password', 'username'], jsonData)
        
        # Add Dictionary to easyDict
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class apicLogin(object):
    def __init__(self, apic, user, pword):
        self.apic = apic
        self.user = user
        self.pword = pword

    def login(self):
        # Load login json payload
        payload = '''
        {{
            "aaaUser": {{
                "attributes": {{
                    "name": "{user}",
                    "pwd": "{pword}"
                }}
            }}
        }}
        '''.format(user=self.user, pword=self.pword)
        payload = json.loads(payload,
                             object_pairs_hook=OrderedDict)
        s = requests.Session()
        # Try the request, if exception, exit program w/ error
        try:
            # Verify is disabled as there are issues if it is enabled
            r = s.post('https://{}/api/aaaLogin.json'.format(self.apic),
                       data=json.dumps(payload), verify=False)
            # Capture HTTP status code from the request
            status = r.status_code
            # Capture the APIC cookie for all other future calls
            cookies = r.cookies
            # Log login status/time(?) somewhere
            if status == 400:
                print("Error 400 - Bad Request - ABORT!")
                print("Probably have a bad URL")
                sys.exit()
            if status == 401:
                print("Error 401 - Unauthorized - ABORT!")
                print("Probably have incorrect credentials")
                sys.exit()
            if status == 403:
                print("Error 403 - Forbidden - ABORT!")
                print("Server refuses to handle your request")
                sys.exit()
            if status == 404:
                print("Error 404 - Not Found - ABORT!")
                print("Seems like you're trying to POST to a page that doesn't"
                      " exist.")
                sys.exit()
        except Exception as e:
            print("Something went wrong logging into the APIC - ABORT!")
            # Log exit reason somewhere
            raise LoginFailed(e)
        self.cookies = cookies
        return cookies

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class fabric(object):
    def __init__(self, type):
        self.type = type

    #=============================================================================
    # Function - Date and Time Policy
    #=============================================================================
    def date_time(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.DateandTime']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - DNS Profiles
    #=============================================================================
    def dns_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.dnsProfiles']['allOf'][1]['properties']

        if not kwargs['dns_domains'] == None:
            if ',' in kwargs['dns_domains']:
                kwargs['dns_domains'] = kwargs['dns_domains'].split(',')
            else:
                kwargs['dns_domains'] = [kwargs['dns_domains']]
            if not kwargs['default_domain'] == None:
                if not kwargs['default_domain'] in kwargs['dns_domains']:
                    kwargs['dns_domains'].append(kwargs['default_domain'])

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['name'] = 'default'
        
        # Convert to Lists
        if ',' in templateVars["dns_providers"]:
            templateVars["dns_providers"] = templateVars["dns_providers"].split(',')

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'dns_profiles'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Date and Time Policy - NTP Servers
    #=============================================================================
    def ntp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Ntp']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'date_and_time'
        templateVars['data_subtype'] = 'ntp_servers'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Date and Time Policy - NTP Keys
    #=============================================================================
    def ntp_key(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.NtpKeys']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'date_and_time'
        templateVars['data_subtype'] = 'authentication_keys'

        # Check if the NTP Key is in the Environment.  If not Add it.
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'ntp_key_{kwargs["key_id"]}'
        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Smart CallHome Policy
    #=============================================================================
    def smart_callhome(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartCallHome']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['name'] = 'default'
        templateVars['smtp_server'] = []
        templateVars['smart_destinations'] = []

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'smart_callhome'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Smart CallHome Policy - Smart Destinations
    #=============================================================================
    def smart_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartDestinations']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'smart_callhome'
        templateVars['data_subtype'] = 'smart_destinations'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Smart CallHome Policy - SMTP Server
    #=============================================================================
    def smart_smtp_server(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartSmtpServer']['allOf'][1]['properties']

        # Check for Variable values that could change required arguments
        if 'secure_smtp' in kwargs:
            if 'true' in kwargs['secure_smtp']:
                jsonData = required_args_add(['username'], jsonData)
        else:
            kwargs['secure_smtp'] == 'false'

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'smart_callhome'
        templateVars['data_subtype'] = 'smtp_server'

        # Check if the Smart CallHome SMTP Password is in the Environment and if not add it.
        if 'true' in kwargs['secure_smtp']:
            templateVars['easyDict'] = kwargs['easyDict']
            templateVars['jsonData'] = jsonData
            templateVars["Variable"] = f'smtp_password'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
            templateVars.pop('easyDict')
            templateVars.pop('jsonData')
            templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - SNMP Policy - Client Groups
    #=============================================================================
    def snmp_clgrp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpClientGroups']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Convert to Lists
        if not templateVars["clients"] == None:
            clientDict = {}
            templateVars["clients"] = templateVars["clients"].split(',')
            for i in templateVars["clients"]:
                clientDict.update({i:{}})
            templateVars["clients"] = clientDict

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'snmp_client_groups'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - SNMP Policy - Communities
    #=============================================================================
    def snmp_community(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpCommunities']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'snmp_communities'

        # Check if the SNMP Community is in the Environment.  If not Add it.
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'snmp_community_{kwargs["community_variable"]}'
        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - SNMP Policy - SNMP Trap Destinations
    #=============================================================================
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

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - SNMP Policy
    #=============================================================================
    def snmp_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpPolicy']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['name'] = 'default'
        templateVars['snmp_client_groups'] = []
        templateVars['snmp_communities'] = []
        templateVars['snmp_destinations'] = []
        templateVars['users'] = []

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - SNMP Policy - SNMP Users
    #=============================================================================
    def snmp_user(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpUsers']['allOf'][1]['properties']

        # Check for Variable values that could change required arguments
        if 'privacy_key' in kwargs:
            if not kwargs['privacy_key'] == 'none':
                jsonData = required_args_add(['privacy_key'], jsonData)
        else:
            kwargs['privacy_key'] = 'none'

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'users'

        # Check if the Authorization and Privacy Keys are in the environment and if not add them.
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'snmp_authorization_key_{kwargs["authorization_key"]}'
        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        if not kwargs['privacy_type'] == 'none':
            templateVars["Variable"] = f'snmp_privacy_key_{kwargs["privacy_key"]}'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Reset jsonData
        if not kwargs['privacy_key'] == 'none':
            jsonData = required_args_remove(['privacy_key'], jsonData)
        
        # Add Dictionary to Policy
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Syslog Policy
    #=============================================================================
    def syslog(self, **kwargs):
       # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Syslog']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        Additions = {'name':'default', 'remote_destinations': []}
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'syslog'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Syslog Policy - Syslog Destinations
    #=============================================================================
    def syslog_destinations(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.syslogRemoteDestinations']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'syslog'
        templateVars['data_subtype'] = 'remote_destinations'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class switches(object):
    def __init__(self, type):
        self.type = type
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(template_path + 'switches/'))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)
    #=============================================================================
    # Function - Interface Selectors
    #=============================================================================
    def intf_selector(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.profiles.interfaceSelectors']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['interface_description'] = templateVars['description']
        templateVars['interface_description'] = templateVars['description']
        if len(templateVars['interface'].split(',')) > 2:
            templateVars['sub_port'] = 'true'
        else:
            templateVars['sub_port'] = 'false'
        pop_list = ['access_or_native_vlan', 'description', 'interface_profile',
            'interface_selector', 'node_id', 'pod_id',  'switchport_mode', 
            'trunk_port_allowed_vlans', 
        ]
        for i in pop_list:
            templateVars.pop(i)
        pgt = templateVars['policy_group_type']
        if pgt == 'spine_pg': templateVars.pop('policy_group_type')

        templateVars['class_type'] = 'switches'
        templateVars['data_type'] = 'switch_profiles'
        templateVars['data_subtype'] = 'interfaces'
        templateVars['policy_name'] = kwargs['interface_profile']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Port Conversion
    #=============================================================================
    def port_cnvt(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.switches.portConvert']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        def process_site(site_dict, tempalteVars, **kwargs):
            if site_dict['auth_type'] == 'username':
                if not site_dict['login_domain'] == None:
                    apic_user = f"apic#{site_dict['login_domain']}\\{site_dict['username']}"
                else:
                    apic_user = site_dict['username']
                
                # Add Dictionary to Policy
                templateVars['easyDict'] = kwargs['easyDict']
                templateVars['jsonData'] = jsonData
                templateVars["Variable"] = 'apicPass'
                kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
                templateVars.pop('easyDict')
                templateVars.pop('jsonData')
                templateVars.pop('Variable')
                apic_pass = os.environ.get('TF_VAR_apicPass')
                node_list = vlan_list_full(templateVars['node_list'])
                port_list = vlan_list_full(templateVars['port_list'])

                controller = site_dict['controller']
                fablogin = apicLogin(controller, apic_user, apic_pass)
                cookies = fablogin.login()

                for node in node_list:
                    templateVars['node_id'] = node
                    for port in port_list:
                        # Locate template for method
                        template_file = "check_ports.json"
                        template = self.templateEnv.get_template(template_file)
                        # Render template w/ values from dicts
                        payload = template.render(templateVars)
                        uri = 'ncapi/config'
                        # port_modes = get(controller, payload, cookies, uri, template_file)

                        # Locate template for method
                        templateVars['port'] = f"1/{port}"
                        template_file = "port_convert.json"
                        template = self.templateEnv.get_template(template_file)
                        # Render template w/ values from dicts
                        payload = template.render(templateVars)
                        uri = 'ncapi/config'
                        apic_post(controller, payload, cookies, uri, template_file)

        # Loop Through the Site Groups
        if re.search('Grp_', templateVars['site_group']):
            site_group = kwargs['easyDict']['sites']['site_groups'][kwargs['site_group']][0]
            for site in site_group['sites']:
                # Process the Site Port Conversions
                siteDict = kwargs['easyDict']['sites']['site_settings'][site][0]
                process_site(siteDict, templateVars, **kwargs)
        else:
            # Process the Site Port Conversions
            siteDict = kwargs['easyDict']['sites']['site_settings'][kwargs['site_group']][0]
            process_site(siteDict, templateVars, **kwargs)

        return kwargs['easyDict']

    #=============================================================================
    # Function - Switch Inventory
    #=============================================================================
    def switch(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.profiles.switchProfiles']['allOf'][1]['properties']

        if re.search('Grp_[A-F]', kwargs['site_group']):
            print(f"\n-----------------------------------------------------------------------------\n")
            print(f"   Error on Worksheet {kwargs['ws'].title}, Row {kwargs['row_num']} site_group, value {kwargs['site_group']}.")
            print(f"   A Leaf can only be assigned to one Site.  Exiting....")
            print(f"\n-----------------------------------------------------------------------------\n")
            exit()

        mgmt_list = ['inband', 'ooband']
        atype_list = ['ipv4', 'ipv6']
        for mgmt in mgmt_list:
            for atype in atype_list:
                if f'{mgmt}_{atype}' in kwargs:
                    if not kwargs[f'{mgmt}_{atype}'] == None:
                        jsonData = required_args_add([f'{mgmt}_{atype}', f'{mgmt}_{atype}_gateway'], jsonData)


        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        for mgmt in mgmt_list:
            for atype in atype_list:
                if f'{mgmt}_{atype}' in jsonData['required_args']:
                    jsonData = required_args_remove([f'{mgmt}_{atype}', f'{mgmt}_{atype}_gateway'], jsonData)

        # Modify the Format of the IP Addressing
        Additions = {
            'inband_addressing':{
                'ipv4_address':templateVars['inband_ipv4'],
                'ipv4_gateway':templateVars['inband_ipv4_gateway'],
                'ipv6_address':templateVars['inband_ipv6'],
                'ipv6_gateway':templateVars['inband_ipv6_gateway'],
                'management_epg':templateVars['inband_mgmt_epg'],
            },
            'interfaces':[],
            'name':templateVars['switch_name'],
            'ooband_addressing':{
                'ipv4_address':templateVars['ooband_ipv4'],
                'ipv4_gateway':templateVars['ooband_ipv4_gateway'],
                'ipv6_address':templateVars['ooband_ipv6'],
                'ipv6_gateway':templateVars['ooband_ipv6_gateway'],
                'management_epg':templateVars['ooband_mgmt_epg'],
            },
        }
        templateVars.update(Additions)
        ptypes = ['ipv4', 'ipv6']
        mtypes = ['inband', 'ooband']
        for mtype in mtypes:
            templateVars.pop(f'{mtype}_mgmt_epg')
            for ptype in ptypes:
                templateVars.pop(f'{mtype}_{ptype}')
                templateVars.pop(f'{mtype}_{ptype}_gateway')
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'switches'
        templateVars['data_type'] = 'switch_profiles'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        # return kwargs['easyDict']

        # Create or Modify the Interface Selector Workbook
        siteDict = kwargs['easyDict']['sites']['site_settings'][kwargs['site_group']][0]
        kwargs['excel_workbook'] = '%s_interface_selectors.xlsx' % (siteDict['site_name'])
        kwargs['wb_sw'] = load_workbook(kwargs['excel_workbook'])
        interface_selector_workbook(templateVars, **kwargs)

        # Remove Site Worksheet if it Exists
        ws_site = kwargs['wb_sw'].get_sheet_names()
        for sheetName in ws_site:
            if sheetName in ['Sites']:
                sheetToDelete = kwargs['wb_sw'].get_sheet_by_name(sheetName)
                kwargs['wb_sw'].remove_sheet(sheetToDelete)
                kwargs['wb_sw'].save(kwargs['excel_workbook'])

        # Set the wb and ws before it is over-written
        wb = kwargs['wb']
        ws = kwargs['ws']

        # Evaluate The Interface Selectors Worksheet in the Site Workbook
        wb = kwargs['wb_sw']
        class_init = 'switches'
        class_folder = 'switches'
        func_regex = '^intf_selector$'
        ws = wb[f"{templateVars['switch_name']}"]
        rows = ws.max_row
        func_list = findKeys(ws, func_regex)
        stdout_log(ws, None, 'begin')
        for func in func_list:
            count = countKeys(ws, func)
            var_dict = findVars(ws, func, rows, count)
            for pos in var_dict:
                row_num = var_dict[pos]['row']
                del var_dict[pos]['row']
                stdout_log(ws, row_num, 'begin')
                var_dict[pos].update(
                    {
                        'class_folder':class_folder,
                        'easyDict':kwargs['easyDict'],
                        'easy_jsonData':kwargs['easy_jsonData'],
                        'row_num':row_num,
                        'wb':wb,
                        'ws':ws
                    }
                )
                easyDict = eval(f"{class_init}(class_folder).{func}(**var_dict[pos])")

        # Set the wb and ws back
        kwargs['wb'] = wb
        kwargs['ws'] = ws
        kwargs['wb_sw'].close()

        if not templateVars['node_type'] == 'spine':
            if not templateVars['vpc_name'] == None:
                if len(kwargs['easyDict']['switches']['vpc_domains']) > 0:
                    vpc_count = 0
                    if templateVars['site_group'] in kwargs['easyDict']['switches']['vpc_domains'].keys():
                        for i in kwargs['easyDict']['switches']['vpc_domains'][templateVars['site_group']]:
                            if i['name'] == templateVars['vpc_name']:
                                i['switches'].append(templateVars['node_id'])
                                vpc_count =+ 1
                    if vpc_count == 0:
                        # Add Dictionary to easyDict
                        vpcArgs = {
                            'name':templateVars['vpc_name'],
                            'domain_id':templateVars['vpc_domain_id'],
                            'site_group':templateVars['site_group'],
                            'switches':[templateVars['node_id']],
                            'vpc_domain_policy':'default',
                        }
                        vpcArgs['class_type'] = 'switches'
                        vpcArgs['data_type'] = 'vpc_domains'
                        kwargs['easyDict'] = easyDict_append(vpcArgs, **kwargs)

                else:
                    # Add Dictionary to easyDict
                    vpcArgs = {
                        'name':templateVars['vpc_name'],
                        'domain_id':templateVars['vpc_domain_id'],
                        'site_group':[templateVars['site_group']],
                        'switches':[templateVars['node_id']],
                        'vpc_domain_policy':'default',
                    }
                    vpcArgs['class_type'] = 'switches'
                    vpcArgs['data_type'] = 'vpc_domains'
                    kwargs['easyDict'] = easyDict_append(vpcArgs, **kwargs)

        return easyDict

    #=============================================================================
    # Function - Switch Modules
    #=============================================================================
    def sw_modules(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.profiles.switchModules']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Split the Node List into Nodes
        node_list = []
        if ',' in templateVars['node_list']:
            templateVars['node_list'] = templateVars['node_list'].split(',')
        else:
            templateVars['node_list'] = [templateVars['node_list']]
 
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'spine_modules'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class site_policies(object):
    def __init__(self, type):
        self.type = type

    #=============================================================================
    # Function - Site Settings
    #=============================================================================
    def site_id(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['site.Identifiers']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        kwargs["multi_select"] = False
        jsonVars = kwargs['easy_jsonData']['components']['schemas']['easy_aci']['allOf'][1]['properties']
        # Prompt User for the Version of the Controller
        if templateVars['controller_type'] == 'apic':
            # Obtain the APIC version from the API
            templateVars['easyDict'] = kwargs['easyDict']
            templateVars['jsonData'] = jsonData
            templateVars["Variable"] = 'apicPass'
            apic_pass = sensitive_var_value(**templateVars)
            pop_list = ['easyDict', 'jsonData', 'Variable']
            for i in pop_list:
                templateVars.pop(i)
            if not kwargs['login_domain'] == None:
                apic_user = f"apic#{kwargs['login_domain']}\\{kwargs['username']}"
            else:
                apic_user = kwargs['username']
            fablogin = apicLogin(kwargs['controller'], apic_user, apic_pass)
            cookies = fablogin.login()

            # Locate template for method
            template_file = "aaaRefresh.json"
            uri = 'api/aaaRefresh'
            uriResponse = apic_get(kwargs['controller'], cookies, uri, template_file)
            verJson = uriResponse.json()
            templateVars['version'] = verJson['imdata'][0]['aaaLogin']['attributes']['version']
        else:
            # NDO Version
            kwargs["var_description"] = f'Select the Nexus Dashboard Orchestrator Version'\
                f' for the Site "{templateVars["site_name"]}".'
            kwargs["jsonVars"] = jsonVars['easyDict']['latest_versions']['ndo_versions']['enum']
            kwargs["defaultVar"] = jsonVars['easyDict']['latest_versions']['ndo_versions']['default']
            kwargs["varType"] = 'NDO Version'
            templateVars['version'] = variablesFromAPI(**kwargs)
            #templateVars['version'] = '3.7.1g'

        if templateVars['controller_type'] == 'apic': 
            site_wb = '%s_interface_selectors.xlsx' % (kwargs['site_name'])
            if not os.path.isfile(site_wb):
                kwargs['wb'].save(filename=site_wb)
                wb_wr = load_workbook(site_wb)
                ws_wr = wb_wr.get_sheet_names()
                for sheetName in ws_wr:
                    if sheetName not in ['Sites']:
                        sheetToDelete = wb_wr.get_sheet_by_name(sheetName)
                        wb_wr.remove_sheet(sheetToDelete)
                wb_wr.save(filename=site_wb)

        # Add Dictionary to easyDict
        kwargs['site_group'] = templateVars['site_id']
        templateVars['class_type'] = 'sites'
        templateVars['data_type'] = 'site_settings'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        kwargs['easyDict'] = OrderedDict(sorted(kwargs['easyDict'].items()))
        return kwargs['easyDict']

    #=============================================================================
    # Function - Site Settings
    #=============================================================================
    def site_settings(self, **kwargs):
        args = kwargs['args']
        easyDict = kwargs['easyDict']
        jsonData = kwargs['easy_jsonData']['components']['schemas']['easy_aci']['allOf'][1]['properties']
        templateVars = {}
        templateVars['annotation'] = 'orchestrator:terraform:easy-aci-v%s' % (jsonData['version'])
        templateVars['class_type'] = 'sites'
        for k, v in easyDict['sites']['site_settings'].items():
            site_name = v[0]['site_name']
            if v[0]['controller_type'] == 'apic':
                templateVars['apicHostname'] = v[0]['controller']
                templateVars['apic_version'] = v[0]['version']
                if v[0]['auth_type'] == 'username':
                    if not v[0]['login_domain'] == None:
                        login_domain = v[0]['login_domain']
                        username = v[0]['username']
                        templateVars['apicUser'] = f"apic#{login_domain}\\{username}"
                    else:
                        templateVars['apicUser'] = v[0]['username']
            else:
                templateVars['ndoHostname'] = v[0]['controller']
                templateVars['ndoUser'] = v[0]['username']
                templateVars['ndo_version'] = v[0]['version']
                templateVars['users'] = []
                if not v[0]['login_domain'] == None:
                    templateVars['ndoDomain'] = v[0]['login_domain']
            
            templateVars['template_type'] = 'variables'
            templateVars = OrderedDict(sorted(templateVars.items()))
            siteDirs = next(os.walk(os.path.join(args.dir, site_name)))[1]
            kwargs['auth_type'] = v[0]['auth_type']
            kwargs['controller_type'] = v[0]['controller_type']
            kwargs["initial_write"] = True
            kwargs['site_group'] = v[0]['site_id']
            kwargs["template_file"] = 'variables.jinja2'
            kwargs["tfvars_file"] = 'variables'
            for folder in siteDirs:
                kwargs["dest_dir"] = folder
                write_to_site(templateVars, **kwargs)

            if v[0]['run_location'] == 'tfc' and v[0]['configure_terraform_cloud'] == "true":
                terraform_cloud().create_terraform_workspaces(siteDirs, site_name, **kwargs)

        return kwargs['easyDict']

    #=============================================================================
    # Function - Site Groups
    #=============================================================================
    def group_id(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['site.Groups']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars['row_num'] = kwargs['row_num']
        templateVars['ws'] = kwargs['ws']
        validating.site_groups(**templateVars)
        templateVars.pop('row_num')
        templateVars.pop('ws')

        sites = []
        for x in range(1, 16):
            if not kwargs[f'site_{x}'] == None:
                sites.append(kwargs[f'site_{x}'])

        # Save the Site Information into Environment Variables
        os.environ[kwargs['site_group']] = '%s' % (templateVars)

        # Add Dictionary to easyDict
        templateVars = {
            'site_group':kwargs['site_group'],
            'sites':sites,
        }
        templateVars['class_type'] = 'sites'
        templateVars['data_type'] = 'site_groups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class system_settings(object):
    def __init__(self, type):
        self.type = type

    #=============================================================================
    # Function - APIC Connectivity Preference
    #=============================================================================
    def apic_preference(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.apicConnectivityPreference']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'apic_connectivity_preference'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - BGP Autonomous System Number
    #=============================================================================
    def bgp_asn(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.bgpASN']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'bgp_autonomous_system_number'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - BGP Route Reflectors
    #=============================================================================
    def bgp_rr(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.bgpRouteReflector']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Convert to Lists
        templateVars["node_list"] = vlan_list_full(templateVars["node_list"])

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'bgp_route_reflectors'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Global AES Passphrase Encryption Settings
    #=============================================================================
    def global_aes(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.globalAesEncryptionSettings']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'global_aes_encryption_settings'
        
        # If enable_encryption confirm aes_passphrase is set
        if kwargs['enable_encryption'] == 'true':
            templateVars['easyDict'] = kwargs['easyDict']
            templateVars['jsonData'] = jsonData
            templateVars["Variable"] = 'aes_passphrase'
            kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
            templateVars.pop('easyDict')
            templateVars.pop('jsonData')
            templateVars.pop('Variable')

        # Add Dictionary to easyDict
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class tenants(object):
    def __init__(self, type):
        self.type = type

    #=============================================================================
    # Function - APIC Inband Configuration
    #=============================================================================
    def apic_inb(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.apic.InbandMgmt']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['tenant'] = 'mgmt'
        templateVars['management_epg'] = templateVars['inband_epg']

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'apics_inband_mgmt_addresses'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Application Profiles
    #=============================================================================
    def app_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationProfiles']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['monitoring_policy'] = 'default'

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'application_profiles'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Bridge Domains
    #=============================================================================
    def bd_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bridgeDomains']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Attach the Bridge Domain General Policy Additional Attributes
        if kwargs['easyDict']['tenants']['bridge_domains_general'].get(templateVars['general_policy']):
            templateVars['general'] = kwargs['easyDict']['tenants']['bridge_domains_general'][templateVars['general_policy']]
        else:
            validating.error_policy_not_found('general_policy', **kwargs)

        # Attach the Bridge Domain L3 Configuration Policy Additional Attributes
        if kwargs['easyDict']['tenants']['bridge_domains_l3'].get(templateVars['l3_policy']):
            templateVars['l3_configurations'] = kwargs['easyDict']['tenants']['bridge_domains_l3'][templateVars['l3_policy']]
        else:
            validating.error_policy_not_found('l3_policy', **kwargs)
        
        # Move Variables to the Advanced/Troubleshooting Map
        atr = templateVars['l3_configurations']
        advanced_troubleshooting = {
            'disable_ip_data_plane_learning_for_pbr':atr['disable_ip_data_plane_learning_for_pbr'],
            'endpoint_clear':templateVars['endpoint_clear'],
            'first_hop_security_policy':atr['first_hop_security_policy'],
            'intersite_bum_traffic_allow':atr['intersite_bum_traffic_allow'],
            'intersite_l2_stretch':atr['intersite_l2_stretch'],
            'monitoring_policy':'default',
            'netflow_monitor_policies':atr['netflow_monitor_policies'],
            'optimize_wan_bandwidth':atr['optimize_wan_bandwidth'],
            'netflow_monitor_policies':atr['netflow_monitor_policies'],
            'rogue_coop_exception_list':atr['rogue_coop_exception_list'],
        }
        templateVars['advanced_troubleshooting'] = OrderedDict(sorted(advanced_troubleshooting.items()))
        
        # Move Variables to the General Map
        templateVars['general'].update({
            'alias':templateVars['alias'],
            'annotations':templateVars['annotations'],
            'description':templateVars['description'],
            'global_alias':templateVars['global_alias'],
            'vrf':templateVars['vrf'],
            'vrf_tenant':templateVars['vrf_tenant']
        })
        templateVars['general'] = OrderedDict(sorted(templateVars['general'].items()))

        # Move Variables to the L3 Configurations Map
        templateVars['l3_configurations'].update({
            'associated_l3outs':{
                'l3out':templateVars['l3out'],
                'link_local_ipv6_address':templateVars['link_local_ipv6_address'],
                'tenant':templateVars['vrf_tenant'],
                'route_profile':templateVars['l3_configurations']['route_profile']
            },
            'custom_mac_address':templateVars['custom_mac_address'],
            'subnets':{},
        })
        aa = templateVars['l3_configurations']['associated_l3outs']
        if aa['l3out'] == None and aa['tenant'] == None and aa['route_profile'] == None:
            templateVars['l3_configurations'].pop('associated_l3outs')
        templateVars['l3_configurations'] = OrderedDict(sorted(templateVars['l3_configurations'].items()))

        pop_list = [
            'alias',
            'annotations',
            'custom_mac_address',
            'description',
            'endpoint_clear',
            'general_policy',
            'global_alias',
            'l3out',
            'l3_policy',
            'link_local_ipv6_address',
            'vrf',
            'vrf_tenant'
        ]
        for i in pop_list:
            templateVars.pop(i)
        templateVars = OrderedDict(sorted(templateVars.items()))
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'bridge_domains'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Bridge Domains - General Policies
    #=============================================================================
    def bd_general(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bd.General']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'bridge_domains_general'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Bridge Domains - General Policies
    #=============================================================================
    def bd_l3(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bd.L3Configurations']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'bridge_domains_l3'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Bridge Domain - Subnets
    #=============================================================================
    def bd_subnet(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bd.Subnets']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Modify the templateVars scope and subnet_control
        templateVars['scope'] = {
            'advertise_externally':templateVars['advertise_externally'],
            'shared_between_vrfs':templateVars['shared_between_vrfs']
        }
        templateVars['subnet_control'] = {
            'neighbor_discovery':templateVars['neighbor_discovery'],
            'no_default_svi_gateway':templateVars['no_default_svi_gateway'],
            'querier_ip':templateVars['querier_ip']
        }
        pop_list = [
            'advertise_externally',
            'bridge_domain',
            'gateway_ip',
            'neighbor_discovery',
            'no_default_svi_gateway',
            'querier_ip',
            'shared_between_vrfs',
            'site_group',
            'tenant',
        ]
        for i in pop_list:
            templateVars.pop(i)
        
        bds = kwargs['easyDict']['tenants']['bridge_domains'][kwargs['site_group']]
        for bd in bds:
            if bd['name'] == kwargs['bridge_domain'] and bd['tenant'] == kwargs['tenant']:
                bd['l3_configurations']['subnets'].update({kwargs['gateway_ip']:templateVars})

        # Add Dictionary to easyDict
        return kwargs['easyDict']
        
    #=============================================================================
    # Function - L3Out - BGP Peer Connectivity Profile
    #=============================================================================
    def bgp_peer(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.bgpPeerConnectivityProfile']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Attach the BGP Peer Connectivity Policy Additional Attributes
        if kwargs['easyDict']['tenants']['bgp_peer_policies'].get(templateVars['bgp_peer_shared_policy']):
            templateVars.update(kwargs['easyDict']['tenants']['bgp_peer_policies'][templateVars['bgp_peer_shared_policy']])
        else:
            validating.error_policy_not_found('bgp_peer_shared_policy', **kwargs)

        templateVars.pop('bgp_peer_shared_policy')
        policy_dict = {kwargs['peer_address']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'bgp_peers'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']
        
    #=============================================================================
    # Function - Application Profiles
    #=============================================================================
    def bgp_pfx(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.policies.bgpPrefix']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'policies_bgp_peer_prefix'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3out - BGP Peer Connectivity Profile - Policy
    #=============================================================================
    def bgp_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.bgpPeerConnectivityProfile.Policy']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Modify the templateVars Address Type Controls, BGP Controls, Peer Controls, and Private AS Controls
        templateVars['address_type_controls'] = {
            'af_mcast':templateVars['af_mcast'],
            'af_ucast':templateVars['af_ucast']
        }
        templateVars['bgp_controls'] = {
            'allow_self_as':templateVars['allow_self_as'],
            'as_override':templateVars['as_override'],
            'disable_peer_as_check':templateVars['disable_peer_as_check'],
            'next_hop_self':templateVars['next_hop_self'],
            'send_community':templateVars['send_community'],
            'send_domain_path':templateVars['send_domain_path'],
            'send_extended_community':templateVars['send_extended_community']
        }
        templateVars['peer_controls'] = {
            'bidirectional_forwarding_detection':templateVars['bidirectional_forwarding_detection'],
            'disable_connected_check':templateVars['disable_connected_check']
        }
        templateVars['private_as_control'] = {
            'remove_all_private_as':templateVars['remove_all_private_as'],
            'remove_private_as':templateVars['remove_private_as'],
            'replace_private_as_with_local_as':templateVars['replace_private_as_with_local_as']
        }
        pop_list = [
            'af_mcast',
            'af_ucast',
            'allow_self_as',
            'as_override',
            'disable_peer_as_check',
            'next_hop_self',
            'send_community',
            'send_domain_path',
            'send_extended_community',
            'bidirectional_forwarding_detection',
            'disable_connected_check',
            'remove_all_private_as',
            'remove_private_as',
            'replace_private_as_with_local_as'
        ]
        for i in pop_list:
            templateVars.pop(i)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'bgp_peer_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Contracts
    #=============================================================================
    def contract_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.contract.Contracts']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['subjects'] = []

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'contracts'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Application Profiles
    #=============================================================================
    def contract_assign(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.contract.ContractAssignments']['allOf'][1]['properties']

        pop_list = []
        if 'external_epg' in kwargs['target_type']:
            pop_list = ['l3out', 'external_epgs']
            jsonData = required_args_add(pop_list, jsonData)
        elif 'epg' in kwargs['target_type']:
            pop_list = ['application_epgs', 'application_profile']
            jsonData = required_args_add(pop_list, jsonData)
        elif re.search('^(inb|oob)$', kwargs['target_type']):
            pop_list.append('application_epgs')
            jsonData = required_args_add(pop_list, jsonData)
        elif 'vrf' in kwargs['target_type']:
            pop_list.append('vrfs')
            jsonData = required_args_add(pop_list, jsonData)
        
        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Remove Items in the Pop List
        jsonData = required_args_remove(pop_list, jsonData)

        # Attach the Contract to the EPG and VRF Resource(s)
        if re.search('^(epg|inb|oob)$', kwargs['target_type']):
            easyDict = kwargs['easyDict']['tenants']['application_epgs']
            for i in kwargs['application_epgs'].split(','):
                item_count = 0
                for item in easyDict[kwargs['site_group']]:
                    if item['tenant'] == kwargs['target_tenant'] and item['application_profile'
                    ] == kwargs['application_profile'] and item['name'] == i:
                        contract = {
                            'contract_type':templateVars['contract_type'],
                            'name':templateVars['contract'],
                            'qos_class':templateVars['qos_class'],
                            'schema':templateVars['schema'],
                            'sites':templateVars['sites'],
                            'tenant':templateVars['tenant'],
                            'template':templateVars['template']
                        }
                        item['contracts'].append(contract)
                        item_count += 1
                if item_count == 0:
                    print(f'Did not find Application EPG {i}.  Exiting Script')
                    exit()
        elif 'external_epg' in kwargs['target_type']:
            easyDict = kwargs['easyDict']['tenants']['l3outs']
            for i in kwargs['external_epgs'].split(','):
                item_count = 0
                for item in easyDict[kwargs['site_group']]:
                    if item['tenant'] == kwargs['target_tenant'] and item['name'] == kwargs['l3out']:
                        for ext_epg in item['external_epgs']:
                            if ext_epg['name'] == i:
                                contract = {
                                    'contract_type':kwargs['contract_type'],
                                    'name':kwargs['contract'],
                                    'qos_class':kwargs['qos_class'],
                                    'schema':kwargs['schema'],
                                    'sites':kwargs['sites'],
                                    'tenant':kwargs['tenant'],
                                    'template':kwargs['template']
                                }
                                ext_epg['contracts'].append(contract)
                                item_count += 1
                if item_count == 0:
                    print(f'Did not find External EPG {i}.  Exiting Script')
                    exit()
        elif 'vrf' in kwargs['target_type']:
            easyDict = kwargs['easyDict']['tenants']['vrfs']
            tType = 'vrfs'
            for i in kwargs['vrfs'].split(','):
                item_count = 0
                for item in easyDict[kwargs['site_group']]:
                    if item['tenant'] == kwargs['target_tenant'] and item['name'] == i:
                        contract = {
                            'contract_type':templateVars['contract_type'],
                            'name':templateVars['contract'],
                            'qos_class':templateVars['qos_class'],
                            'schema':templateVars['schema'],
                            'sites':templateVars['sites'],
                            'tenant':templateVars['tenant'],
                            'template':templateVars['template']
                        }
                        item['epg_esg_collection_for_vrfs']['contracts'].append(contract)
                        item_count += 1
                if item_count == 0:
                    print(f'Did not find VRF {i}.  Exiting Script')
                    exit()
        
        # Return EasyDict
        return kwargs['easyDict']

    #=============================================================================
    # Function - Contracts - Add Subject
    #=============================================================================
    def contract_filters(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.contract.ContractFilters']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['directives'] = {
            'enable_policy_compression':templateVars['enable_policy_compression'],
            'log':templateVars['log_packets']
        }
        templateVars['filters'] = templateVars['filters_to_assign'].split(',')
        templateVars.pop('enable_policy_compression')
        templateVars.pop('filters_to_assign')
        templateVars.pop('log_packets')
        templateVars = OrderedDict(sorted(templateVars.items()))

        # Add Dictionary to Policy
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'contracts'
        templateVars['data_subtype'] = 'subjects'
        templateVars['policy_name'] = templateVars['contract_name']
        templateVars.pop('contract_name')
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - EIGRP Interface Policy
    #=============================================================================
    def dhcp_relay(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.policies.dhcpRelay']['allOf'][1]['properties']

        pop_list = []
        if 'external_epg' in kwargs['epg_type']: pop_list = ['l3out']
        else: pop_list = ['application_epg']
        jsonData = required_args_add(pop_list, jsonData)

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        jsonData = required_args_remove(pop_list, jsonData)

        pop_list = ['address', 'application_profile', 'epg', 'epg_type', 'l3out']
        templateVars['dhcp_relay_providers'] = {
            'address':templateVars['address'],
            'application_profile':templateVars['application_profile'],
            'epg':templateVars['epg'],
            'epg_type':templateVars['epg_type'],
            'l3out':templateVars['l3out'],
            'tenant':templateVars['tenant'],
        }

        for i in pop_list:
            templateVars.pop(i)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'policies_dhcp_relay'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - EIGRP Interface Policy
    #=============================================================================
    def eigrp_interface(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.policies.eigrpInterface']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'policies_eigrp_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - EIGRP Interface Profile
    #=============================================================================
    def eigrp_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.eigrpInterfaceProfile']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('profile_name')
        policy_dict = {kwargs['profile_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'eigrp_interface_profiles'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Application EPG
    #=============================================================================
    def epg_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationEpgs']['allOf'][1]['properties']

        # Attach the EPG Policy Additional Attributes
        if kwargs['easyDict']['tenants']['application_epg_policies'].get(kwargs['epg_policy']):
            epgpolicy = kwargs['easyDict']['tenants']['application_epg_policies'][kwargs['epg_policy']]
        else:
            validating.error_policy_not_found('epg_policy', **kwargs)

        pop_list = []
        if re.search('^(inb|oob)$', epgpolicy['epg_type']):
            pop_list.append('application_profile')
            if epgpolicy['epg_type'] == 'oob': pop_list.append('bridge_domain')
            jsonData = required_args_remove(pop_list, jsonData)
            if epgpolicy['epg_type'] == 'inb':
                jsonData = required_args_add(['vlans'], jsonData)
        
        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['monitoring_policy'] = 'default'

        if re.search('^(inb|oob)$', epgpolicy['epg_type']):
            jsonData = required_args_add(pop_list, jsonData)
            if epgpolicy['epg_type'] == 'inb':
                jsonData = required_args_remove(['vlans'], jsonData)

        domain_list = ['physical_domains', 'vmm_domains']
        for i in domain_list:
            if not templateVars[i] == None:
                    templateVars[i] = templateVars[i].split(',')
            else:
                templateVars[i] = []

        vmmpolicy = {}
        if len(templateVars['vmm_domains']) > 0:
            # Attach the EPG VMM Policy Additional Attributes
            if kwargs['easyDict']['tenants']['application_epg_vmm_policies'].get(templateVars['vmm_policy']):
                vmmpolicy.update(kwargs['easyDict']['tenants']['application_epg_vmm_policies'][templateVars['vmm_policy']])
            else:
                validating.error_policy_not_found('vmm_policy', **kwargs)

        templateVars = {**templateVars, **epgpolicy}
        templateVars['contracts'] = []
        templateVars['domains'] = []
        if not templateVars['physical_domains'] == None:
            for i in templateVars['physical_domains']:
                templateVars['domains'].append({'domain': i})
        if not templateVars['vmm_domains'] == None:
            if not templateVars['vmm_vlans'] == None:
                if ',' in  templateVars['vmm_vlans']:
                    templateVars['vmm_vlans'] = [int(s) for s in templateVars['vmm_vlans'].split(',')]
                else:
                     templateVars['vmm_vlans'] = [int( templateVars['vmm_vlans'])]
            for i in templateVars['vmm_domains']:
                templateVars['domains'].append({
                    'allow_micro_segmentation': vmmpolicy['allow_micro_segmentation'],
                    'custom_epg_name': templateVars['custom_epg_name'],
                    'delimiter': vmmpolicy['delimiter'],
                    'deploy_immediacy': vmmpolicy['deploy_immediacy'],
                    'domain': i,
                    'domain_type': 'vmm',
                    'number_of_ports': vmmpolicy['number_of_ports'],
                    'port_allocation': vmmpolicy['port_allocation'],
                    'port_binding': vmmpolicy['port_binding'],
                    'resolution_immediacy': vmmpolicy['resolution_immediacy'],
                    'security': {
                        'allow_promiscuous': vmmpolicy['allow_promiscuous'],
                        'forged_transmits': vmmpolicy['forged_transmits'],
                        'mac_changes': vmmpolicy['mac_changes']
                    },
                    'switch_provider': vmmpolicy['switch_provider'],
                    'vlan_mode': vmmpolicy['vlan_mode'],
                    'vlans': templateVars['vmm_vlans']
                })
        epg_to_aaeps = []
        if not templateVars['epg_to_aaeps'] == None:
            if not templateVars['vlans'] == None:
                if ',' in  str(templateVars['vlans']):
                    templateVars['vlans'] = [int(s) for s in templateVars['vlans'].split(',')]
                else:
                     templateVars['vlans'] = [int( templateVars['vlans'])]
            templateVars['epg_to_aaeps'] = templateVars['epg_to_aaeps'].split(',')
            for i in templateVars['epg_to_aaeps']:
                epg_to_aaeps.append({
                    'aaep': i,
                    'mode': templateVars['epg_to_aaep_mode'],
                    'vlans': templateVars['vlans']
                })
        templateVars['epg_to_aaeps'] = epg_to_aaeps
        if not len(templateVars['epg_to_aaeps']) > 0:
            templateVars.pop('epg_to_aaeps')
        if not len(templateVars['domains']) > 0:
            templateVars.pop('domains')
        if templateVars['epg_type'] == 'inb':
            templateVars['vlan'] = templateVars['vlans'].split(',')[0]

        pop_list = [
            'custom_epg_name',
            'epg_policy',
            'epg_to_aaep_mode',
            'physical_domains',
            'vlans',
            'vmm_domains',
            'vmm_policy',
            'vmm_vlans'
        ]
        for i in pop_list:
            if templateVars.get(i): templateVars.pop(i)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'application_epgs'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - EPG - Policy
    #=============================================================================
    def epg_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationEpg.Policy']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'application_epg_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - EPG - VMM Policy
    #=============================================================================
    def epg_vmm_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationEpg.VMMPolicy']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'application_epg_vmm_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out - Exteranl EPG
    #=============================================================================
    def ext_epg(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.externalEpg']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['contracts'] = []
        templateVars['subnets'] = []

        # Attach the External EPG Policy Additional Attributes
        if kwargs['easyDict']['tenants']['external_epg_policies'].get(templateVars['external_epg_shared_policy']):
            templateVars.update(kwargs['easyDict']['tenants']['external_epg_policies'][templateVars['external_epg_shared_policy']])
        else:
            validating.error_policy_not_found('external_epg_shared_policy', **kwargs)

        pop_list = ['external_epg_shared_policy']
        for i in pop_list:
            templateVars.pop(i)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'l3outs'
        templateVars['data_subtype'] = 'external_epgs'
        templateVars['policy_name'] = kwargs['l3out']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out - External EPG - Policy
    #=============================================================================
    def ext_epg_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.externalEpg.Policy']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'external_epg_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Bridge Domain - Subnets
    #=============================================================================
    def ext_epg_sub(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.externalEpg.Subnet']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Modify the templateVars aggregate, external_epg_classification, and route_control
        templateVars['aggregate'] = {
            'aggregate_export':templateVars['aggregate_export'],
            'aggregate_shared_routes':templateVars['aggregate_shared_routes']
        }
        templateVars['external_epg_classification'] = {
            'external_subnets_for_external_epg':templateVars['external_subnets_for_external_epg'],
            'shared_security_import_subnet':templateVars['shared_security_import_subnet']
        }
        templateVars['route_control'] = {
            'export_route_control_subnet':templateVars['export_route_control_subnet'],
            'shared_route_control_subnet':templateVars['shared_route_control_subnet']
        }
        pop_list = [
            'aggregate_export',
            'aggregate_shared_routes',
            'export_route_control_subnet',
            'external_subnets_for_external_epg',
            'shared_security_import_subnet',
            'shared_route_control_subnet'
        ]
        for i in pop_list:
            templateVars.pop(i)

        # Attach the Subnet to the External EPG
        if templateVars['site_group'] in kwargs['easyDict']['tenants']['l3outs'].keys():
            complete = False
            while complete == False:
                for item in kwargs['easyDict']['tenants']['l3outs'][templateVars['site_group']]:
                    if item['name'] == templateVars['l3out'] and item['tenant'] == templateVars['tenant']:
                        for i in item['external_epgs']:
                            if i['name'] == templateVars['external_epg']:
                                subnets = templateVars['subnets']
                                templateVars.pop('external_epg')
                                templateVars.pop('l3out')
                                templateVars.pop('tenant')
                                templateVars.pop('site_group')
                                templateVars.pop('subnets')
                                templateVars['subnets'] = subnets.split(',')
                                i['subnets'].append(templateVars)
                                complete = True
                                break
                    if complete == True: break

        # Return easyDict
        return kwargs['easyDict']

    #=============================================================================
    # Function - Contract Filter
    #=============================================================================
    def filter_add(self, **kwargs):
        # print(json.dumps(kwargs['easyDict']['tenants']['l3out_logical_node_profiles'], indent=4))
        # exit()

        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.contract.Filters']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['filter_entries'] = []

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'filters'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Contract Filter - Filter Entry
    #=============================================================================
    def filter_entry(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.contract.filterEntry']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'filters'
        templateVars['data_subtype'] = 'filter_entries'
        templateVars['policy_name'] = templateVars['filter_name']
        templateVars.pop('filter_name')
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out
    #=============================================================================
    def l3out_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.L3Out']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['external_epgs'] = []

        # Attach the L3Out Policy Additional Attributes
        if kwargs['easyDict']['tenants']['l3out_policies'].get(templateVars['l3out_shared_policy']):
            templateVars.update(kwargs['easyDict']['tenants']['l3out_policies'][templateVars['l3out_shared_policy']])
        else:
            validating.error_policy_not_found('l3out_shared_policy', **kwargs)

        # Attach the OSPF Routing Profile if defined
        if not templateVars['ospf_external_profile'] == None:
            if kwargs['easyDict']['tenants']['ospf_external_profiles'].get(templateVars['ospf_external_profile']):
                aa = kwargs['easyDict']['tenants']['ospf_external_profiles'][templateVars['ospf_external_profile']]
                templateVars['ospf_external_profile'] = aa
            else:
                validating.error_policy_not_found('ospf_external_profile', **kwargs)
        
        pop_list = [ 'l3out_shared_policy', 'ospf_external_profile']
        for i in pop_list:
            templateVars.pop(i)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'l3outs'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out - Policy
    #=============================================================================
    def l3out_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.L3Out.Policy']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Modify the templateVars Route Control Enforcement
        templateVars['route_control_enforcement'] = {
            'export':templateVars['export'],
            'import':templateVars['import']
        }
        pop_list = [
            'export',
            'import'
        ]
        for i in pop_list:
            templateVars.pop(i)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'l3out_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out - Exteranl EPG
    #=============================================================================
    def node_interface(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.logicalNodeInterfaceProfile']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'l3out_logical_node_profiles'
        templateVars['data_subtype'] = 'interface_profiles'

        # Attach the Node Interface Profile Additional Attributes
        if kwargs['easyDict']['tenants']['node_interface_profile_policies'].get(templateVars['interface_profile_shared_policy']):
            templateVars.update(kwargs['easyDict']['tenants']['node_interface_profile_policies'][templateVars['interface_profile_shared_policy']])
        else:
            validating.error_policy_not_found('interface_profile_shared_policy', **kwargs)

        # Attach the Interface Configuration
        if kwargs['easyDict']['tenants']['node_interface_configurations'].get(templateVars['interface_config_name']):
            templateVars.update(kwargs['easyDict']['tenants']['node_interface_configurations'][templateVars['interface_config_name']])
        else:
            validating.error_policy_not_found('interface_config_name', **kwargs)

        # Attach the BGP Peers if defined
        if not templateVars['bgp_peers'] == None:
            templateVars['bgp_peers'] = []
            for i in kwargs['bgp_peers'].split(','):
                if kwargs['easyDict']['tenants']['bgp_peers'].get(i):
                    aa = kwargs['easyDict']['tenants']['bgp_peers'][i]

                    # If BGP Password is Set, Check Environment for presence.
                    if re.search('^[0-5]$', str(aa['bgp_password'])):
                        templateVars['easyDict'] = kwargs['easyDict']
                        templateVars['jsonData'] = jsonData
                        templateVars["Variable"] = f'bgp_password_{aa["bgp_password"]}'
                        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
                        templateVars.pop('easyDict')
                        templateVars.pop('jsonData')
                        templateVars.pop('Variable')
                        aa['password'] = aa['bgp_password']
                        aa.pop('bgp_password')

                    aa = OrderedDict(sorted(aa.items()))
                    templateVars['bgp_peers'].append(aa)
                else:
                    validating.error_policy_not_found('bgp_peers', **kwargs)

        # Attach the EIGRP Interface Profile if defined
        if not templateVars['eigrp_interface_profile'] == None:
            if kwargs['easyDict']['tenants']['eigrp_interface_profiles'].get(templateVars['eigrp_interface_profile']):
                aa = kwargs['easyDict']['tenants']['eigrp_interface_profiles'][templateVars['eigrp_interface_profile']]
                templateVars['eigrp_interface_profile'] = aa
            else:
                validating.error_policy_not_found('eigrp_interface_profile', **kwargs)

        # Attach the OSPF Interface Profile if defined
        if not templateVars['ospf_interface_profile'] == None:
            if kwargs['easyDict']['tenants']['ospf_interface_profiles'].get(templateVars['ospf_interface_profile']):
                aa = kwargs['easyDict']['tenants']['ospf_interface_profiles'][templateVars['ospf_interface_profile']]
                templateVars['ospf_interface_profile'] = aa
                # If OSPF auth_key is Set, Check Environment for presence.
                if re.search('[0-9]', str(aa['key_id'])):
                    templateVars['easyDict'] = kwargs['easyDict']
                    templateVars['jsonData'] = jsonData
                    templateVars["Variable"] = f'ospf_key_{aa["key_id"]}'
                    kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
                    templateVars.pop('easyDict')
                    templateVars.pop('jsonData')
                    templateVars.pop('Variable')

            else:
                validating.error_policy_not_found('ospf_interface_profile', **kwargs)

        pop_list = [
            'interface_config_name',
            'interface_profile_shared_policy',
            'node_profile',
        ]
        for i in pop_list:
            templateVars.pop(i)


        # Add Dictionary to Policy
        templateVars['policy_name'] = kwargs['node_profile']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out - Logical Node Interface Profile - Interface Configuration
    #=============================================================================
    def node_intf_cfg(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.logicalNodeInterfaceProfile.InterfaceConfiguration']['allOf'][1]['properties']

        pop_list = []
        if re.search('^(l3-port|sub-interface)$', kwargs['interface_type']):
            pop_list.append('auto_state')
            if kwargs['interface_type'] == 'l3-port':
                pop_list.append('encap_scope')
                pop_list.append('mode')
                pop_list.append('vlan')
            jsonData = required_args_remove(pop_list, jsonData)
        
        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        if templateVars['interface_type'] == 'ext-svi':
            if not templateVars['link_local_addresses'] == None:
                if not len(templateVars['link_local_addresses'].split(',')) == 2:
                    validating.error_interface_address('link_local_addresses', **kwargs)
                link_local_a = templateVars['link_local_addresses'].split(',')[0]
                link_local_b = templateVars['link_local_addresses'].split(',')[1]
            else:
                link_local_a = None
                link_local_b = None
            if not templateVars['primary_preferred_addresses'] == None:
                if not len(templateVars['primary_preferred_addresses'].split(',')) == 2:
                    validating.error_interface_address('primary_preferred_addresses', **kwargs)
                primary_a = templateVars['primary_preferred_addresses'].split(',')[0]
                primary_b = templateVars['primary_preferred_addresses'].split(',')[1]
            else:
                primary_a = None
                primary_b = None
            if not templateVars['secondary_addresses'] == None:
                if not len(templateVars['secondary_addresses'].split(',')) % 2  == 0:
                    validating.error_interface_address('secondary_addresses', **kwargs)
                xsplit = templateVars['secondary_addresses'].split(',')
                half = len(xsplit)//2
                secondaries_a = xsplit[:half]
                secondaries_b = xsplit[half:]
            else:
                secondaries_a = None
                secondaries_b = None
            
            templateVars['svi_addresses'] = [
                {
                    'link_local_address':link_local_a,
                    'primary_preferred_address':primary_a,
                    'secondary_addresses':secondaries_a,
                    'side':'A'
                },
                {
                    'link_local_address':link_local_b,
                    'primary_preferred_address':primary_b,
                    'secondary_addresses':secondaries_b,
                    'side':'B'
                }
            ]
        else:
            if not templateVars['link_local_addresses'] == None:
                if not len(templateVars['link_local_addresses'].split(',')) == 1:
                    validating.error_interface_address('link_local_addresses', **kwargs)
            templateVars['link_local_address'] = templateVars['link_local_addresses']
            if not templateVars['primary_preferred_addresses'] == None:
                if not len(templateVars['primary_preferred_addresses'].split(',')) == 1:
                    validating.error_interface_address('primary_preferred_addresses', **kwargs)
            templateVars['primary_preferred_address'] = templateVars['primary_preferred_addresses']
            if not templateVars['secondary_addresses'] == None:
                templateVars['secondary_addresses'] = templateVars['secondary_addresses'].split(',')

        templateVars.pop('link_local_addresses')
        templateVars.pop('primary_preferred_addresses')
        templateVars.pop('policy_name')
        templateVars = OrderedDict(sorted(templateVars.items()))
        policy_dict = {kwargs['policy_name']:templateVars}

        if re.search('^(l3-port|sub-interface)$', kwargs['interface_type']):
            jsonData = required_args_add(pop_list, jsonData)

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'node_interface_configurations'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out - Logical Node Interface Profile - Policy
    #=============================================================================
    def node_intf_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.logicalNodeInterfaceProfile.Policy']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'node_interface_profile_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - L3Out - Logical Node Profile
    #=============================================================================
    def node_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.logicalNodeProfile']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['interface_profiles'] = []
        templateVars['node_router_ids'] = templateVars['node_router_ids'].split(',')
        templateVars['node_list'] = [int(s) for s in str(templateVars['node_list']).split(',')]
        templateVars['nodes'] = []
        for x in range(0, len(templateVars['node_list'])):
            node = {
                'node_id':templateVars['node_list'][x],
                'router_id':templateVars['node_router_ids'][x],
                'use_router_id_as_loopback':templateVars['use_router_id_as_loopback']
            }
            templateVars['nodes'].append(node)

        # Remove Arguments
        templateVars.pop('node_list')
        templateVars.pop('node_router_ids')
        templateVars.pop('use_router_id_as_loopback')

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'l3out_logical_node_profiles'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Application Profiles
    #=============================================================================
    def ospf_interface(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.policies.ospfInterface']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'policies_ospf_interface'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - OSPF Interface Profile
    #=============================================================================
    def ospf_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.ospfInterfaceProfile']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        
        templateVars['name'] = templateVars['profile_name']
        templateVars.pop('profile_name')
        policy_dict = {kwargs['profile_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'ospf_interface_profiles'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - OSPF Routing Profile
    #=============================================================================
    def ospf_routing(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.l3out.ospfRoutingProfile']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Modify the templateVars OSPF Area Control
        templateVars['ospf_area_control'] = {
            'originate_summary_lsa':templateVars['originate_summary_lsa'],
            'send_redistribution_lsas_into_nssa_area':templateVars['send_redistribution_lsas_into_nssa_area'],
            'suppress_forwarding_address':templateVars['suppress_forwarding_address']
        }
        pop_list = [
            'originate_summary_lsa',
            'send_redistribution_lsas_into_nssa_area',
            'suppress_forwarding_address'
        ]
        for i in pop_list:
            templateVars.pop(i)

        templateVars.pop('profile_name')
        policy_dict = {kwargs['profile_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'ospf_external_profiles'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Tenants
    #=============================================================================
    def tenant_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.Tenants']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['monitoring_policy'] = 'default'
        templateVars['tenant'] = templateVars['name']

        if re.search(r'\d{1,16}', kwargs['site_group']):
            if kwargs['easyDict']['tenants']['sites'].get(kwargs['site_group']):
                templateVars['sites'] = []
                templateVars['users'] = []
                for i in kwargs['easyDict']['tenants']['sites'][kwargs['site_group']]:
                    if i['tenant'] == templateVars['name']:
                        for x in i['users'].split(','):
                            if not x in templateVars['users']:
                                templateVars['users'].append(x)
                        templateVars['sites'].append(i)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'tenants'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Tenants
    #=============================================================================
    def tenant_site(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.Sites']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'sites'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - VRFs
    #=============================================================================
    def vrf_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.Vrfs']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['communities'] = []
        templateVars['epg_esg_collection_for_vrfs'] = dict(
            contracts = [],
            match_type = kwargs['label_match_criteria']
        )
        templateVars.pop('label_match_criteria')

        # Attach the VRF Policy Additional Attributes
        if kwargs['easyDict']['tenants']['vrf_policies'].get(templateVars['vrf_policy']):
            templateVars.update(kwargs['easyDict']['tenants']['vrf_policies'][templateVars['vrf_policy']])
        else:
            validating.error_policy_not_found('vrf_policy', **kwargs)

        # Add the ESG Collection Argument


        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'vrfs'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']
        
    #=============================================================================
    # Function - VRF - Communities
    #=============================================================================
    def vrf_community(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.vrf.Community']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'vrfs'
        templateVars['data_subtype'] = 'communities'

        # Check if the SNMP Community is in the Environment.  If not Add it.
        templateVars['easyDict'] = kwargs['easyDict']
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'vrf_snmp_community_{kwargs["community_variable"]}'
        kwargs['easyDict'] = sensitive_var_site_group(**templateVars)
        templateVars.pop('easyDict')
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['policy_name'] = templateVars['vrf']
        templateVars.pop('vrf')
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - VRF - Policy
    #=============================================================================
    def vrf_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.vrf.Policy']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        per_list = ['bgp_timers_per_address_family', 'eigrp_timers_per_address_family', 'ospf_timers_per_address_family']
        for i in per_list:
            if not templateVars[i] == None:
                dict_list = []
                for v in templateVars[i].split(','):
                    if '_' in v: dict_list.append({ 'address_family': v.split('_')[0], 'policy': v.split('_')[1] })
                templateVars[i] = dict_list

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'vrf_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Class.
#=====================================================================================
class terraform_cloud(object):
    def __init__(self):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(template_path + 'terraform/'))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)

    #=============================================================================
    # Function - Create Terraform Cloud Workspaces
    #=============================================================================
    def create_terraform_workspaces(self, folders, site, **kwargs):
        easyDict = kwargs['easyDict']
        jsonData = kwargs['easy_jsonData']['components']['schemas']['site.Identifiers']['allOf'][1]['properties']
        site_group = kwargs['site_group']
        tfcb_config = []
        valid = False
        while valid == False:
            templateVars = {}
            templateVars["Description"] = f'Terraform Cloud Workspaces for Site {site}'
            templateVars["varInput"] = f'Do you want to Proceed with creating Workspaces in Terraform Cloud?'
            templateVars["varDefault"] = 'Y'
            templateVars["varName"] = 'Terraform Cloud Workspaces'
            runTFCB = varBoolLoop(**templateVars)
            valid = True
        if runTFCB == True:
            templateVars = { 'site_group':site_group }
            templateVars["terraform_cloud_token"] = terraform_cloud().terraform_token()
            
            # Obtain Terraform Cloud Organization
            if os.environ.get('tfc_organization') is None:
                templateVars["tfc_organization"] = terraform_cloud().tfc_organization(**templateVars)
                os.environ['tfc_organization'] = templateVars["tfc_organization"]
            else:
                templateVars["tfc_organization"] = os.environ.get('tfc_organization')
            tfcb_config.append({'tfc_organization':templateVars["tfc_organization"]})
            
            # Obtain Terraform Cloud Agent_Pool
            if os.environ.get('agentPoolId') is None:
                templateVars["agentPoolId"] = terraform_cloud().tfc_agent_pool(**templateVars)
                os.environ['agentPoolId'] = templateVars["tfc_organization"]
            else:
                templateVars["agentPoolId"] = os.environ.get('agentPoolId')
            tfcb_config.append({'agentPoolId':templateVars["agentPoolId"]})
            
            # Obtain Version Control Provider
            if os.environ.get('tfc_vcs_provider') is None:
                tfc_vcs_provider,templateVars["tfc_oath_token"] = terraform_cloud(
                ).tfc_vcs_providers(**templateVars)
                templateVars["tfc_vcs_provider"] = tfc_vcs_provider
                os.environ['tfc_vcs_provider'] = tfc_vcs_provider
                os.environ['tfc_oath_token'] = templateVars["tfc_oath_token"]
            else:
                templateVars["tfc_vcs_provider"] = os.environ.get('tfc_vcs_provider')
                templateVars["tfc_oath_token"] = os.environ['tfc_oath_token']

            # Obtain Version Control Base Repo
            if os.environ.get('vcsBaseRepo') is None:
                templateVars["vcsBaseRepo"] = terraform_cloud().tfc_vcs_repository(**templateVars)
                os.environ['vcsBaseRepo'] = templateVars["vcsBaseRepo"]
            else:
                templateVars["vcsBaseRepo"] = os.environ.get('vcsBaseRepo')
            
            # Set Some of the default Variables that user is not Prompted for
            templateVars["allowDestroyPlan"] = False
            templateVars["executionMode"] = 'agent'
            templateVars["queueAllRuns"] = False
            templateVars["speculativeEnabled"] = True
            templateVars["triggerPrefixes"] = []

            # Set the Terraform Version for the Workspace
            templateVars["terraformVersion"] = kwargs['easyDict']['latest_versions']['terraform_version']

            # Loop through the Site Folders
            folders.sort()
            for folder in folders:
                templateVars["autoApply"] = True
                templateVars["Description"] = f'Site {site} - {folder}'
                templateVars["globalRemoteState"] = False
                templateVars["workingDirectory"] = f'{site}/{folder}'

                templateVars["Description"] = f'Name of the Workspace to Create in Terraform Cloud for:\n'\
                    f'  - Site: "{site}"\n  - Folder: "{folder}"'
                templateVars["varDefault"] = f'{site}_{folder}'
                templateVars["varInput"] = f'Terraform Cloud Workspace Name. [{site}_{folder}]: '
                templateVars["varName"] = f'Workspace Name'
                templateVars["maximum"] = 90
                templateVars["minimum"] = 1
                templateVars["pattern"] = '^[a-zA-Z0-9\\-\\_]+$'
                templateVars["workspaceName"] = varStringLoop(**templateVars)
                tfcb_config.append({folder:templateVars["workspaceName"]})
                # templateVars["vcsBranch"] = ''

                # Create Terraform Cloud Workspace
                templateVars['workspace_id'] = terraform_cloud().tfcWorkspace(**templateVars)

                #==============================================
                # Add Sensitive Variables to Workspace
                #==============================================
                site_list = easyDict['sensitive_vars'].keys()
                var_list = []
                for s in site_list:
                    if re.search('Grp_', s):
                        if site_group in easyDict['sites']['site_groups'][s][0]['sites']:
                            if easyDict['sensitive_vars'][s].get(folder):
                                var_list = var_list + easyDict['sensitive_vars'][s][folder]
                    elif str(s) == str(site_group):
                        if easyDict['sensitive_vars'][s].get(folder):
                            var_list = var_list + easyDict['sensitive_vars'][s][folder]

                if kwargs['controller_type'] == 'apic' and kwargs['auth_type'] == 'username':
                   var_list.append('apicPass')
                elif kwargs['controller_type'] == 'apic':
                   var_list.append('certName')
                   var_list.append('privateKey')
                else:
                   var_list.append('ndoPass')
                var_list.sort()
                for var in var_list:
                    if 'cert' in var or 'private' in var:
                        templateVars["Multi_Line_Input"] = True
                    print(f'* Adding {var} to {templateVars["workspaceName"]}')
                    templateVars['class_type'] = 'tfcVariables'
                    templateVars["Description"] = ''
                    templateVars["easyDict"] = easyDict
                    templateVars['jsonData'] = jsonData
                    templateVars["Variable"] = var
                    templateVars["varId"] = var
                    templateVars["varKey"] = var
                    sensitive_var_site_group(**templateVars)
                    templateVars["Sensitive"] = True
                    terraform_cloud().tfcVariables(**templateVars)

        else:
            print(f'\n-------------------------------------------------------------------------------------------\n')
            print(f'  Skipping Step to Create Terraform Cloud Workspaces.')
            print(f'\n-------------------------------------------------------------------------------------------\n')
     
    #=============================================================================
    # Function - Terraform Cloud - API Token
    #=============================================================================
    def terraform_token(self):
        # -------------------------------------------------------------------------------------------------------------------------
        # Check to see if the TF_VAR_terraform_cloud_token is already set in the Environment, and if not prompt the user for Input
        #--------------------------------------------------------------------------------------------------------------------------
        if os.environ.get('TF_VAR_terraform_cloud_token') is None:
            print(f'\n----------------------------------------------------------------------------------------\n')
            print(f'  The Run or State Location was set to Terraform Cloud.  To Store the Data in Terraform')
            print(f'  Cloud we will need a User or Org Token to authenticate to Terraform Cloud.  If you ')
            print(f'  have not already obtained a token see instructions in how to obtain a token Here:\n')
            print(f'   - https://www.terraform.io/docs/cloud/users-teams-organizations/api-tokens.html')
            print(f'\n----------------------------------------------------------------------------------------\n')

            while True:
                user_response = input('press enter to continue: ')
                if user_response == '':
                    break

            # Request the TF_VAR_terraform_cloud_token Value from the User
            while True:
                try:
                    secure_value = stdiomask.getpass(prompt=f'Enter the value for the Terraform Cloud Token: ')
                    break
                except Exception as e:
                    print('Something went wrong. Error received: {}'.format(e))

            # Add the TF_VAR_terraform_cloud_token to the Environment
            os.environ['TF_VAR_terraform_cloud_token'] = '%s' % (secure_value)
            terraform_cloud_token = secure_value
        else:
            terraform_cloud_token = os.environ.get('TF_VAR_terraform_cloud_token')

        return terraform_cloud_token

    #=============================================================================
    # Function - Terraform Cloud - VCS Repository
    #=============================================================================
    def tfc_agent_pool(self, **templateVars):
        #-------------------------------
        # Configure the Variables URL
        #-------------------------------
        url = 'https://app.terraform.io/api/v2/organizations/%s/agent-pools' % (templateVars['tfc_organization'])
        tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
        tf_header = {'Authorization': tf_token,
                'Content-Type': 'application/vnd.api+json'
        }

        #----------------------------------------------------------------------------------
        # Get the Contents of the Workspace to Search for the Variable
        #----------------------------------------------------------------------------------
        status,json_data = tfc_get(url, tf_header, 'Get Agent Pools')

        #--------------------------------------------------------------
        # Parse the JSON Data to see if the Variable Exists or Not.
        #--------------------------------------------------------------

        if status == 200:
            # print(json.dumps(json_data, indent = 4))
            json_data = json_data['data']
            pool_list = []
            pool_dict = {}
            for item in json_data:
                pool_list.append(item['attributes']['name'])
                pool_dict.update({item['attributes']['name']:item['id']})

            # print(vcsProvider)
            templateVars["multi_select"] = False
            templateVars["var_description"] = "Terraform Cloud Agent Pools:"
            templateVars["jsonVars"] = sorted(pool_list)
            templateVars["varType"] = 'Agent Pools'
            templateVars["defaultVar"] = ''
            agentPool = variablesFromAPI(**templateVars)

            agentPool = pool_dict[agentPool]
            return agentPool
        else:
            print(status)

    #=============================================================================
    # Function - Terraform Cloud - Organization
    #=============================================================================
    def tfc_organization(self, **templateVars):
        #-------------------------------
        # Configure the Variables URL
        #-------------------------------
        url = 'https://app.terraform.io/api/v2/organizations'
        tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
        tf_header = {'Authorization': tf_token,
                'Content-Type': 'application/vnd.api+json'
        }

        #----------------------------------------------------------------------------------
        # Get the Contents of the Workspace to Search for the Variable
        #----------------------------------------------------------------------------------
        status,json_data = tfc_get(url, tf_header, 'Get Terraform Cloud Organizations')

        #--------------------------------------------------------------
        # Parse the JSON Data to see if the Variable Exists or Not.
        #--------------------------------------------------------------

        if status == 200:
            # print(json.dumps(json_data, indent = 4))
            json_data = json_data['data']
            tfcOrgs = []
            for item in json_data:
                for k, v in item.items():
                    if k == 'id':
                        tfcOrgs.append(v)

            # print(tfcOrgs)
            templateVars["multi_select"] = False
            templateVars["var_description"] = 'Terraform Cloud Organizations:'
            templateVars["jsonVars"] = tfcOrgs
            templateVars["varType"] = 'Terraform Cloud Organization'
            templateVars["defaultVar"] = ''
            tfc_organization = variablesFromAPI(**templateVars)
            return tfc_organization
        else:
            print(status)

    #=============================================================================
    # Function - Terraform Cloud - VCS Repository
    #=============================================================================
    def tfc_vcs_repository(self, **templateVars):
        #-------------------------------
        # Configure the Variables URL
        #-------------------------------
        oauth_token = templateVars["tfc_oath_token"]
        url = 'https://app.terraform.io/api/v2/oauth-tokens/%s/authorized-repos?oauth_token_id=%s' % (oauth_token, oauth_token)
        tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
        tf_header = {'Authorization': tf_token,
                'Content-Type': 'application/vnd.api+json'
        }

        #----------------------------------------------------------------------------------
        # Get the Contents of the Workspace to Search for the Variable
        #----------------------------------------------------------------------------------
        status,json_data = tfc_get(url, tf_header, 'Get VCS Repos')

        #--------------------------------------------------------------
        # Parse the JSON Data to see if the Variable Exists or Not.
        #--------------------------------------------------------------

        if status == 200:
            # print(json.dumps(json_data, indent = 4))
            json_data = json_data['data']
            repo_list = []
            for item in json_data:
                for k, v in item.items():
                    if k == 'id':
                        repo_list.append(v)

            # print(vcsProvider)
            templateVars["multi_select"] = False
            templateVars["var_description"] = "Terraform Cloud VCS Base Repository:"
            templateVars["jsonVars"] = sorted(repo_list)
            templateVars["varType"] = 'VCS Base Repository'
            templateVars["defaultVar"] = ''
            vcsBaseRepo = variablesFromAPI(**templateVars)

            return vcsBaseRepo
        else:
            print(status)

    #=============================================================================
    # Function - Terraform Cloud - VCS Providers
    #=============================================================================
    def tfc_vcs_providers(self, **templateVars):
        #-------------------------------
        # Configure the Variables URL
        #-------------------------------
        url = 'https://app.terraform.io/api/v2/organizations/%s/oauth-clients' % (templateVars["tfc_organization"])
        tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
        tf_header = {'Authorization': tf_token,
                'Content-Type': 'application/vnd.api+json'
        }

        #----------------------------------------------------------------------------------
        # Get the Contents of the Workspace to Search for the Variable
        #----------------------------------------------------------------------------------
        status,json_data = tfc_get(url, tf_header, 'Get VCS Repos')

        #--------------------------------------------------------------
        # Parse the JSON Data to see if the Variable Exists or Not.
        #--------------------------------------------------------------

        if status == 200:
            # print(json.dumps(json_data, indent = 4))
            json_data = json_data['data']
            vcsProvider = []
            vcsAttributes = []
            for item in json_data:
                for k, v in item.items():
                    if k == 'id':
                        vcs_id = v
                    elif k == 'attributes':
                        vcs_name = v['name']
                    elif k == 'relationships':
                        oauth_token = v["oauth-tokens"]["data"][0]["id"]
                vcsProvider.append(vcs_name)
                vcs_repo = {
                    'id':vcs_id,
                    'name':vcs_name,
                    'oauth_token':oauth_token
                }
                vcsAttributes.append(vcs_repo)

            # print(vcsProvider)
            templateVars["multi_select"] = False
            templateVars["var_description"] = "Terraform Cloud VCS Provider:"
            templateVars["jsonVars"] = vcsProvider
            templateVars["varType"] = 'VCS Provider'
            templateVars["defaultVar"] = ''
            vcsRepoName = variablesFromAPI(**templateVars)

            for i in vcsAttributes:
                if i["name"] == vcsRepoName:
                    tfc_oauth_token = i["oauth_token"]
                    vcsBaseRepo = i["id"]
            # print(f'vcsBaseRepo {vcsBaseRepo} and tfc_oauth_token {tfc_oauth_token}')
            return vcsBaseRepo,tfc_oauth_token
        else:
            print(status)

    #=============================================================================
    # Function - Terraform Cloud - GET Workspaces
    #=============================================================================
    def tfcWorkspace(self, **templateVars):
        #-------------------------------
        # Configure the Workspace URL
        #-------------------------------
        url = 'https://app.terraform.io/api/v2/organizations/%s/workspaces/%s' %  (templateVars['tfc_organization'], templateVars['workspaceName'])
        tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
        tf_header = {'Authorization': tf_token,
                'Content-Type': 'application/vnd.api+json'
        }

        #----------------------------------------------------------------------------------
        # Get the Contents of the Organization to Search for the Workspace
        #----------------------------------------------------------------------------------
        status,json_data = tfc_get(url, tf_header, 'workspace_check')

        #--------------------------------------------------------------
        # Parse the JSON Data to see if the Workspace Exists or Not.
        #--------------------------------------------------------------
        key_count = 0
        workspace_id = ''
        # print(json.dumps(json_data, indent = 4))
        if status == 200:
            if json_data['data']['attributes']['name'] == templateVars['workspaceName']:
                workspace_id = json_data['data']['id']
                key_count =+ 1

        #--------------------------------------------
        # If the Workspace was not found Create it.
        #--------------------------------------------

        opSystem = platform.system()
        if opSystem == 'Windows':
            workingDir = templateVars["workingDirectory"]
            templateVars["workingDirectory"] = workingDir.replace('\\', '/')

        if re.search(r'\/', templateVars["workingDirectory"]):
            workingDir = templateVars["workingDirectory"]
            templateVars["workingDirectory"] = workingDir[1 : ]
        
        if not key_count > 0:
            #-------------------------------
            # Get Workspaces the Workspace URL
            #-------------------------------
            url = 'https://app.terraform.io/api/v2/organizations/%s/workspaces/' %  (templateVars['tfc_organization'])
            tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
            tf_header = {'Authorization': tf_token,
                    'Content-Type': 'application/vnd.api+json'
            }

            # Define the Template Source
            template_file = 'workspace.jinja2'
            template = self.templateEnv.get_template(template_file)

            # Create the Payload
            payload = template.render(templateVars)

            if print_payload:
                print(payload)

            # Post the Contents to Terraform Cloud
            json_data = tfc_post(url, payload, tf_header, template_file)

            # Get the Workspace ID from the JSON Dump
            workspace_id = json_data['data']['id']
            key_count =+ 1

        else:
            #-----------------------------------
            # Configure the PATCH Variables URL
            #-----------------------------------
            url = 'https://app.terraform.io/api/v2/workspaces/%s/' %  (workspace_id)
            tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
            tf_header = {'Authorization': tf_token,
                    'Content-Type': 'application/vnd.api+json'
            }

            # Define the Template Source
            template_file = 'workspace.jinja2'
            template = self.templateEnv.get_template(template_file)

            # Create the Payload
            payload = template.render(templateVars)

            if print_payload:
                print(payload)

            # PATCH the Contents to Terraform Cloud
            json_data = tfc_patch(url, payload, tf_header, template_file)
            # Get the Workspace ID from the JSON Dump
            workspace_id = json_data['data']['id']
            key_count =+ 1

        if not key_count > 0:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'\n   Unable to Determine the Workspace ID for "{templateVars["workspaceName"]}".')
            print(f'\n   Exiting...')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

        # print(json.dumps(json_data, indent = 4))
        return workspace_id

    #=============================================================================
    # Function - Terraform Cloud - Workspace Remove
    #=============================================================================
    def tfcWorkspace_remove(self, **templateVars):
        #-------------------------------
        # Configure the Workspace URL
        #-------------------------------
        url = 'https://app.terraform.io/api/v2/organizations/%s/workspaces/%s' %  (templateVars['tfc_organization'], templateVars['workspaceName'])
        tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
        tf_header = {'Authorization': tf_token,
                'Content-Type': 'application/vnd.api+json'
        }

        #----------------------------------------------------------------------------------
        # Delete the Workspace of the Organization to Search for the Workspace
        #----------------------------------------------------------------------------------
        response = delete(url, headers=tf_header)
        # print(response)

        #--------------------------------------------------------------
        # Parse the JSON Data to see if the Workspace Exists or Not.
        #--------------------------------------------------------------
        del_count = 0
        workspace_id = ''
        # print(json.dumps(json_data, indent = 4))
        if response.status_code == 200:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'    Successfully Deleted Workspace "{templateVars["workspaceName"]}".')
            print(f'\n-----------------------------------------------------------------------------\n')
            del_count =+ 1
        elif response.status_code == 204:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'    Successfully Deleted Workspace "{templateVars["workspaceName"]}".')
            print(f'\n-----------------------------------------------------------------------------\n')
            del_count =+ 1

        if not del_count > 0:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'    Unable to Determine the Workspace ID for "{templateVars["workspaceName"]}".')
            print(f'\n-----------------------------------------------------------------------------\n')

    #=============================================================================
    # Function - Terraform Cloud - Workspace Variables
    #=============================================================================
    def tfcVariables(self, **templateVars):
        #-------------------------------
        # Configure the Variables URL
        #-------------------------------
        url = 'https://app.terraform.io/api/v2/workspaces/%s/vars' %  (templateVars['workspace_id'])
        tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
        tf_header = {'Authorization': tf_token,
                'Content-Type': 'application/vnd.api+json'
        }

        #----------------------------------------------------------------------------------
        # Get the Contents of the Workspace to Search for the Variable
        #----------------------------------------------------------------------------------
        status,json_data = tfc_get(url, tf_header, 'variable_check')

        #--------------------------------------------------------------
        # Parse the JSON Data to see if the Variable Exists or Not.
        #--------------------------------------------------------------
        # print(json.dumps(json_data, indent = 4))
        json_text = json.dumps(json_data)
        key_count = 0
        var_id = ''
        if 'id' in json_text:
            for keys in json_data['data']:
                if keys['attributes']['key'] == templateVars['Variable']:
                    var_id = keys['id']
                    key_count =+ 1

        #--------------------------------------------
        # If the Variable was not found Create it.
        # If it is Found Update the Value
        #--------------------------------------------
        if not key_count > 0:
            # Define the Template Source
            template_file = 'variables.jinja2'
            template = self.templateEnv.get_template(template_file)

            # Create the Payload
            payload = template.render(templateVars)

            if print_payload:
                print(payload)

            # Post the Contents to Terraform Cloud
            json_data = tfc_post(url, payload, tf_header, template_file)

            # Get the Workspace ID from the JSON Dump
            var_id = json_data['data']['id']
            key_count =+ 1

        else:
            #-----------------------------------
            # Configure the PATCH Variables URL
            #-----------------------------------
            url = 'https://app.terraform.io/api/v2/workspaces/%s/vars/%s' %  (templateVars['workspace_id'], var_id)
            tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
            tf_header = {'Authorization': tf_token,
                    'Content-Type': 'application/vnd.api+json'
            }

            # Define the Template Source
            template_file = 'variables.jinja2'
            template = self.templateEnv.get_template(template_file)

            # Create the Payload
            templateVars.pop('varId')
            payload = template.render(templateVars)

            if print_payload:
                print(payload)

            # PATCH the Contents to Terraform Cloud
            json_data = tfc_patch(url, payload, tf_header, template_file)
            # Get the Workspace ID from the JSON Dump
            var_id = json_data['data']['id']
            key_count =+ 1

        if not key_count > 0:
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f"\n   Unable to Determine the Variable ID for {templateVars['Variable']}.")
            print(f"\n   Exiting...")
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

        # print(json.dumps(json_data, indent = 4))
        return var_id
