#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from collections import OrderedDict
from easy_functions import countKeys, findKeys, findVars
from easy_functions import easyDict_append, easyDict_append_policy, easyDict_append_subtype
from easy_functions import interface_selector_workbook, post, process_kwargs
from easy_functions import required_args_add, required_args_remove
from easy_functions import sensitive_var_site_group, stdout_log, validate_args
from easy_functions import variablesFromAPI, vlan_list_full
from openpyxl import load_workbook
import ast
import jinja2
import json
import os
import pkg_resources
import re
import requests
import sys
import validating
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Global path to main Template directory
json_path = pkg_resources.resource_filename('classes', 'templates/')

class LoginFailed(Exception):
    pass

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Function.
#=====================================================================================
class access(object):
    def __init__(self, type):
        self.type = type

    #======================================================
    # Function - Global Policies - AAEP Profiles
    #======================================================
    def aep_profile(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.global.attachableAccessEntityProfile']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        domain_list = ['physical_domains', 'l3_domains', 'vmm_domains']
        for i in domain_list:
            if not templateVars[f'{i}'] == None:
                if ',' in templateVars[f'{i}']:
                    templateVars[f'{i}'] = templateVars[f'{i}'].split(',')

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'aaep_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Policies - CDP
    #======================================================
    def cdp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.cdpInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'cdp_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Policies - Fibre Channel
    #======================================================
    def fibre_channel(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.fibreChannelInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'fibre_channel_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Policy Groups - Interface Policies
    # Shared Policies with Access and Bundle Poicies Groups
    #======================================================
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

    #======================================================
    # Function - Interface Policies - L2 Interfaces
    #======================================================
    def l2_interface(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.L2Interface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'l2_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Domain - Layer 3
    #======================================================
    def l3_domain(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.domains.Layer3']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'layer3_domains'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Leaf Policy Group
    #======================================================
    def leaf_pg(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.switches.leafPolicyGroup']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'leaf_policy_groups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Policies - Link Level (Speed)
    #======================================================
    def link_level(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.linkLevel']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'link_level_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Policies - LLDP
    #======================================================
    def lldp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.lldpInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'lldp_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Policies - Mis-Cabling Protocol
    #======================================================
    def mcp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.mcpInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'mcp_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Policy Group - Access
    #======================================================
    def pg_access(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.leafAccessPort']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        if ',' in templateVars['netflow_monitor_policies']:
            templateVars['netflow_monitor_policies'] = templateVars['netflow_monitor_policies'].split(',')

        # Attach the Interface Policy Additional Attributes
        if kwargs['easyDict']['access']['interface_policies'].get(templateVars['interface_policy']):
            templateVars.update(kwargs['easyDict']['access']['interface_policies'][templateVars['interface_policy']])
        else:
            validating.error_policy_not_found('interface_policy', **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'leaf_port_group_access'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Policy Group - VPC/Port-Channel
    #======================================================
    def pg_bundle(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.leafBundle']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        if ',' in templateVars['netflow_monitor_policies']:
            templateVars['netflow_monitor_policies'] = templateVars['netflow_monitor_policies'].split(',')

        # Attach the Interface Policy Additional Attributes
        if kwargs['easyDict']['access']['interface_policies'].get(templateVars['interface_policy']):
            templateVars.update(kwargs['easyDict']['access']['interface_policies'][templateVars['interface_policy']])
        else:
            validating.error_policy_not_found('interface_policy', **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'leaf_port_group_bundle'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Policy Group - Breakout
    #======================================================
    def pg_breakout(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.leafBreakOut']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'leaf_port_group_breakout'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Policy Group - Spine
    #======================================================
    def pg_spine(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policyGroups.spineAccessPort']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'spine_port_group_access'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Domains - Physical
    #======================================================
    def phys_domain(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.domains.Physical']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'physical_domains'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Policies - Port Channel
    #======================================================
    def port_channel(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.PortChannel']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'port_channel_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Port Conversion
    #======================================================
    def port_cnvt(self, **kwargs):
        print('hello')
        
    #======================================================
    # Function - Interface Policies - Port Security
    #======================================================
    def port_security(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.portSecurity']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'port_security_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Spine Policy Group
    #======================================================
    def spine_pg(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.switches.spinePolicyGroup']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'spine_policy_groups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Policies - Spanning Tree
    #======================================================
    def stp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.policies.spanningTreeInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'spanning_tree_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - VLAN Pools
    #======================================================
    def vlan_pool(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.pools.Vlan']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'vlan_pools'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Virtual Networking - Controllers
    #======================================================
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

    #======================================================
    # Function - Virtual Networking - Credentials
    #======================================================
    def vmm_creds(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['access.vmm.Credentials']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'vmm_password_{kwargs["password"]}'
        sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'virtual_networking'
        templateVars['data_subtype'] = 'credentials'
        templateVars['policy_name'] = kwargs['domain_name']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Virtual Networking - Domains
    #======================================================
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

    #======================================================
    # Function - Virtual Networking - Controllers
    #======================================================
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

    #======================================================
    # Function - Virtual Networking - Controllers
    #======================================================
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

    #======================================================
    # Function - Configuration Backup - Export Policies
    #======================================================
    def export_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['admin.exportPolicy']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        Additions = {
            'configuration_export': [],
        }
        templateVars.update(Additions)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Global Security Settings
    #======================================================
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

    #======================================================
    # Function - Configuration Backup  - Remote Host
    #======================================================
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

    #======================================================
    # Function - RADIUS Authentication
    #======================================================
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

        # Check if the secert is in the Environment.  If not Add it.
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'radius_key_{kwargs["key"]}'
        sensitive_var_site_group(**templateVars)
        if templateVars['server_monitoring'] == 'enabled':
            # Check if the Password is in the Environment.  If not Add it.
            templateVars["Variable"] = f'radius_monitoring_password_{kwargs["monitoring_password"]}'
            sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Reset jsonData
        if not kwargs['server_monitoring'] == 'disabled':
            jsonData = required_args_remove(['monitoring_password', 'username'], jsonData)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'radius'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

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

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Reset jsonData
        if kwargs['authentication_type'] == 'usePassword':
            jsonData = required_args_remove(['username'], jsonData)
        
        # Convert to Lists
        if ',' in templateVars["remote_hosts"]:
            templateVars["remote_hosts"] = templateVars["remote_hosts"].split(',')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'configuration_backups'
        templateVars['data_subtype'] = 'configuration_export'
        templateVars['policy_name'] = kwargs['scheduler_name']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Global Security Settings
    #======================================================
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

    #======================================================
    # Function - TACACS+ Authentication
    #======================================================
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

        # Check if the secert is in the Environment.  If not Add it.
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'tacacs_key_{kwargs["key"]}'
        sensitive_var_site_group(**templateVars)
        if templateVars['server_monitoring'] == 'enabled':
            # Check if the Password is in the Environment.  If not Add it.
            templateVars["Variable"] = f'tacacs_monitoring_password_{kwargs["monitoring_password"]}'
            sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Reset jsonData
        if not kwargs['server_monitoring'] == 'disabled':
            jsonData = required_args_remove(['monitoring_password', 'username'], jsonData)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'admin'
        templateVars['data_type'] = 'tacacs'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Function.
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

    #======================================================
    # Function - DNS Profiles
    #======================================================
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

        Additions = {
            'name':'default',
        }
        templateVars.update(Additions)
        
        # Convert to Lists
        if ',' in templateVars["dns_providers"]:
            templateVars["dns_providers"] = templateVars["dns_providers"].split(',')

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'dns_profiles'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Date and Time Policy - NTP Servers
    #======================================================
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

    #======================================================
    # Function - Date and Time Policy - NTP Keys
    #======================================================
    def ntp_key(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.NtpKeys']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Check if the NTP Key is in the Environment.  If not Add it.
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'ntp_key_{kwargs["key_id"]}'
        sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'date_and_time'
        templateVars['data_subtype'] = 'authentication_keys'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Smart CallHome Policy
    #======================================================
    def smart_callhome(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.smartCallHome']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        Additions = {
            'name':'default',
            'smtp_server': [],
            'smart_destinations': [],
        }
        templateVars.update(Additions)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'smart_callhome'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Smart CallHome Policy - Smart Destinations
    #======================================================
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

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Check if the Smart CallHome SMTP Password is in the Environment and if not add it.
        if 'true' in kwargs['secure_smtp']:
            templateVars['jsonData'] = jsonData
            templateVars["Variable"] = f'smtp_password'
            sensitive_var_site_group(**templateVars)
            templateVars.pop('jsonData')
            templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'smart_callhome'
        templateVars['data_subtype'] = 'smtp_server'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - Client Groups
    #======================================================
    def snmp_clgrp(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpClientGroups']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Convert to Lists
        if ',' in templateVars["clients"]:
            templateVars["clients"] = templateVars["clients"].split(',')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'fabric'
        templateVars['data_type'] = 'snmp_policies'
        templateVars['data_subtype'] = 'snmp_client_groups'
        templateVars['policy_name'] = 'default'
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - SNMP Policy - Communities
    #======================================================
    def snmp_community(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpCommunities']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
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

    #======================================================
    # Function - SNMP Policy
    #======================================================
    def snmp_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.snmpPolicy']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
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

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Syslog Policy
    #======================================================
    def syslog(self, **kwargs):
       # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['fabric.Syslog']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Syslog Policy - Syslog Destinations
    #======================================================
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
# for detailed information on the Arguments used by this Function.
#=====================================================================================
class switches(object):
    def __init__(self, type):
        self.type = type
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(json_path + 'switches/'))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)
    #======================================================
    # Function - Interface Selectors
    #======================================================
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
        if len(templateVars['port'].split(',')) > 2:
            templateVars['sub_port'] = 'true'
        else:
            templateVars['sub_port'] = 'false'
        templateVars.pop('access_or_native_vlan')
        templateVars.pop('description')
        templateVars.pop('interface_profile')
        templateVars.pop('interface_selector')
        templateVars.pop('node_id')
        templateVars.pop('pod_id')
        templateVars.pop('switchport_mode')
        templateVars.pop('trunk_port_allowed_vlans')
        templateVars['class_type'] = 'switches'
        templateVars['data_type'] = 'switch_profiles'
        templateVars['data_subtype'] = 'interfaces'
        templateVars['policy_name'] = kwargs['interface_profile']
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Interface Selectors
    #======================================================
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
                templateVars['jsonData'] = jsonData
                templateVars["Variable"] = 'apicPass'
                sensitive_var_site_group(**templateVars)
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
                        post(controller, payload, cookies, uri, template_file)
        if re.search('Grp_', templateVars['site_group']):
            group_id = '%s' % (templateVars['site_group'])
            site_group = ast.literal_eval(os.environ[group_id])
            for x in range(1, 16):
                if not site_group[f'site_{x}'] == None:
                    site_id = 'site_id_%s' % (site_group[f'site_{x}'])
                    site_dict = ast.literal_eval(os.environ[site_id])

                    # Process the Site Port Conversions
                    process_site(site_dict, templateVars, **kwargs)
        else:
            site_id = 'site_id_%s' % (templateVars['site_group'])
            site_dict = ast.literal_eval(os.environ[site_id])

            # Process the Site Port Conversions
            process_site(site_dict, templateVars, **kwargs)

        return kwargs['easyDict']

    #======================================================
    # Function - Switch Inventory
    #======================================================
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
        site_id = 'site_id_%s' % (templateVars['site_group'])
        site_dict = ast.literal_eval(os.environ[site_id])
        kwargs['excel_workbook'] = '%s_interface_selectors.xlsx' % (site_dict['site_name'])
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

        # kwargs['wb_sw'].save(kwargs['excel_workbook'])

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

    #======================================================
    # Function - Interface Policies - Spanning Tree
    #======================================================
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
            node_list = templateVars['node_list'].split(',')
        else:
            node_list = [templateVars['node_list']]
        templateVars.pop('node_list')
 
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'access'
        templateVars['data_type'] = 'spine_modules'
        for node in node_list:
            templateVars['node_id'] = node
            kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

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

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'apic_connectivity_preference'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - BGP Autonomous System Number
    #======================================================
    def bgp_asn(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.bgpASN']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'bgp_asn'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - BGP Route Reflectors
    #======================================================
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
        templateVars['data_type'] = 'bgp_rr'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Global AES Passphrase Encryption Settings
    #======================================================
    def global_aes(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['system.globalAesEncryptionSettings']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        if kwargs['enable_encryption'] == 'true':
            templateVars["Variable"] = 'aes_passphrase'
            templateVars['jsonData'] = jsonData
            sensitive_var_site_group(**templateVars)
            templateVars.pop('jsonData')
            templateVars.pop('Variable')
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'system_settings'
        templateVars['data_type'] = 'global_aes_encryption_settings'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
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

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        kwargs["multi_select"] = False
        jsonVars = kwargs['easy_jsonData']['components']['schemas']['easy_aci']['allOf'][1]['properties']
        # Prompt User for the Version of the Controller
        if templateVars['controller_type'] == 'apic':
            # APIC Version
            kwargs["var_description"] = f"Select the Version that Most Closely matches "\
                f'your version for the Site "{templateVars["site_name"]}".'
            kwargs["jsonVars"] = jsonVars['apic_versions']['enum']
            kwargs["defaultVar"] = jsonVars['apic_versions']['default']
            kwargs["varType"] = 'APIC Version'
            templateVars['version'] = variablesFromAPI(**kwargs)
        else:
            # NDO Version
            kwargs["var_description"] = f'Select the Version that Most Closely matches '\
                f'your version for the Site "{templateVars["site_name"]}".'
            kwargs["jsonVars"] = jsonVars['easyDict']['latest_versions']['ndo_versions']['enum']
            kwargs["defaultVar"] = jsonVars['easyDict']['latest_versions']['ndo_versions']['default']
            kwargs["varType"] = 'NDO Version'
            templateVars['version'] = variablesFromAPI(**kwargs)

        # Save the Site Information into Environment Variables
        site_id = 'site_id_%s' % (kwargs['site_id'])
        os.environ[site_id] = '%s' % (templateVars)

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
        # # If the state_location is tfc configure workspaces in the cloud
        # if kwargs['run_location'] == 'tfc' and kwargs['configure_terraform_cloud'] == 'true':
        #     # Initialize the Class
        #     class_init = '%s()' % ('lib_terraform.Terraform_Cloud')
        # 
        #     # Get terraform_cloud_token
        #     terraform_cloud().terraform_token()
        # 
        #     # Get workspace_ids
        #     easy_jsonData = kwargs['easy_jsonData']
        #     terraform_cloud().create_terraform_workspaces(easy_jsonData, folder_list, kwargs["site_name"])
        # 
        #     if kwargs['auth_type'] == 'user_pass' and kwargs["controller_type"] == 'apic':
        #         var_list = ['apicUrl', 'aciUser', 'aciPass']
        #     elif kwargs["controller_type"] == 'apic':
        #         var_list = ['apicUrl', 'certName', 'privateKey']
        #     else:
        #         var_list = ['ndoUrl', 'ndoDomain', 'ndoUser', 'ndoPass']
        # 
        #     # Get var_ids
        #     tf_var_dict = {}
        #     for folder in folder_list:
        #         folder_id = 'site_id_%s_%s' % (kwargs['site_id'], folder)
        #         # kwargs['workspace_id'] = workspace_dict[folder_id]
        #         kwargs['description'] = ''
        #         # for var in var_list:
        #         #     tf_var_dict = tf_variables(class_init, folder, var, tf_var_dict, **kwargs)
        # 

        # Return Dictionary
        kwargs['easyDict'] = OrderedDict(sorted(kwargs['easyDict'].items()))
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
        kwargs['easyDict'] = OrderedDict(sorted(kwargs['easyDict'].items()))
        return kwargs['easyDict']
