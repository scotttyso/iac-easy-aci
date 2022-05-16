#!/usr/bin/env python3

#======================================================
# Source Modules
#======================================================
from easy_functions import countKeys, easyDict_append_policy, findVars
from easy_functions import easyDict_append, easyDict_append_subtype
from easy_functions import process_kwargs, process_workbook
from easy_functions import sensitive_var_site_group, validate_args
import jinja2
import json
import pkg_resources
import re
import validating

aci_template_path = pkg_resources.resource_filename('class_tenants', 'templates/')

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
class tenants(object):
    def __init__(self, type):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(aci_template_path + '%s/') % (type))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)
        self.type = type

    #======================================================
    # Function - APIC Inband Configuration
    #======================================================
    def apic_inb(self, **kwargs):
        # Dicts for required and optional args
        required_args = {
            'site_group': '',
            'name': '',
            'node_id': '',
            'pod_id': '',
            'Inband_EPG': ''
        }
        optional_args = {
            'Inband_IPv4': '',
            'Inband_GWv4': '',
            'Inband_IPv6': '',
            'Inband_GWv6': '',
        }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        # Configure the Generic Template Variables
        templateVars['Device_Type'] = 'apic'
        templateVars['Type'] = 'in_band'
        templateVars['EPG'] = templateVars['Inband_EPG']
        templateVars['IPv4'] = templateVars['Inband_IPv4']
        templateVars['GWv4'] = templateVars['Inband_GWv4']
        templateVars['IPv6'] = templateVars['Inband_IPv6']
        templateVars['GWv6'] = templateVars['Inband_GWv6']

        # Initialize the Class
        lib_aci_ref = 'Access_Policies'
        class_init = '%s(ws)' % (lib_aci_ref)

        # Assign the APIC Inband Management IP's
        eval("%s.%s(wb, ws, row_num, **templateVars)" % (class_init, 'mgmt_static'))

    #======================================================
    # Function - Application Profiles
    #======================================================
    def app_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationProfile']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['monitoring_policy'] = 'default'

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'application_profiles'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Bridge Domains
    #======================================================
    def bd_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bridgeDomains']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
        templateVars['advanced_troubleshooting'] = {
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
        
        # Move Variables to the General Map
        templateVars['general'].update({
            'alias':templateVars['alias'],
            'annotations':templateVars['annotations'],
            'description':templateVars['description'],
            'global_alias':templateVars['global_alias'],
            'vrf':templateVars['vrf'],
            'vrf_tenant':templateVars['vrf_tenant']
        })

        # Move Variables to the L3 Configurations Map
        templateVars['l3_configurations'].update({
            'associated_l3outs':{
                'l3out':templateVars['l3out'],
                'l3out_tenant':templateVars['vrf_tenant'],
                'route_profile':templateVars['l3_configurations']['route_profile']
            },
            'custom_mac_address':templateVars['custom_mac_address'],
            'description':templateVars['description'],
            'global_alias':templateVars['global_alias'],
            'subnets':[],
            'vrf':templateVars['vrf'],
            'vrf_tenant':templateVars['vrf_tenant']
        })

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'bridge_domains'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Bridge Domains - General Policies
    #======================================================
    def bd_general(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bd.General']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'bridge_domains_general'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Bridge Domains - General Policies
    #======================================================
    def bd_l3(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bd.L3Configurations']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'bridge_domains_l3'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - BGP Peers
    #======================================================
    def bgp_peer(self, **kwargs):
        # Open the Network Policies Worksheet
        ws_net = kwargs['wb']['Network Policies']
        rows = ws_net.max_row

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'Path_Policy_Name': '',
                         'Peer_Interface': '',
                         'Peer_Address': '',
                         'Remote_ASN': '',
                         'eBGP_Multihop_TTL': '',
                         'Weight': '',
                         'Local_ASN_Config': '',
                         'Admin_State': '',
                         'BGP_Peer_Policy': '',
                         'tenant': '',
                         'L3Out': '',
                         'Profile_Name': '',
                         'Node_Profile': '',
                         'Interface_Profile': '',
                         'Interface_Type': '',
                         'Pod_ID': '',
                         'Node1_ID': '',
                         'Interface_or_PG': '',
                         'allow_self_as': '',
                         'as_override': '',
                         'disable_peer_as_check': '',
                         'next_hop_self': '',
                         'send_community': '',
                         'send_ext_community': '',
                         'allowed_self_as_count': '',
                         'bfd': '',
                         'disable_connected_check': '',
                         'AF_Mcast': '',
                         'AF_Ucast': '',
                         'remove_all_private_as': '',
                         'remove_private_as': '',
                         'private_to_local': ''}
        optional_args = {'description': '',
                         'BGP_Password': '',
                         'BGP_Peer_Prefix_Policy': '',
                         'Local_ASN': '',
                         'Node2_ID': '',
                         'Policy_Name': '',
                         'Prefix_tenant': '',
                         'PFX_description': '',
                         'Action': '',
                         'Maximum_Prefixes': '',
                         'Restart_Time': '',
                         'Threshold': ''}

        # Get the Node Policies from the Network Policies Tab
        rows = kwargs['ws'].max_row
        func = 'l3out_path'
        count = countKeys(kwargs['ws'], func)
        row_path = ''
        var_dict = findVars(kwargs['ws'], func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('Policy_Name') == kwargs.get('Path_Policy_Name'):
                row_path = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        # Get the Node Policies from the Network Policies Tab
        func = 'bgp_profile'
        bgp_count = countKeys(ws_net, func)
        row_bgp = ''
        var_dict = findVars(ws_net, func, rows, bgp_count)
        for pos in var_dict:
            if var_dict[pos].get('Profile_Name') == kwargs.get('BGP_Peer_Policy'):
                row_bgp = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        if not kwargs.get('BGP_Peer_Prefix_Policy') == None:
            # Get the Node Policies from the Network Policies Tab
            func = 'pfx_policy'
            pfx_count = countKeys(ws_net, func)
            row_pfx = ''
            var_dict = findVars(ws_net, func, rows, pfx_count)
            for pos in var_dict:
                if var_dict[pos].get('Policy_Name') == kwargs.get('BGP_Peer_Prefix_Policy'):
                    row_pfx = var_dict[pos]['row']
                    del var_dict[pos]['row']
                    kwargs = {**kwargs, **var_dict[pos]}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.values('Peer_Interface', templateVars['Peer_Interface'], ['Interface', 'Loopback'])
            validating.ip_address('Peer_Address', templateVars['Peer_Address'])
            validating.number_check('Remote_ASN', templateVars['Remote_ASN'], 1, 4294967295)
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
            validating.number_check('eBGP_Multihop_TTL', templateVars['eBGP_Multihop_TTL'], 1, 255)
            validating.number_check('Weight', templateVars['Weight'], 0, 65535)
            if not templateVars['BGP_Peer_Prefix_Policy'] == None:
                validating.validator('BGP_Peer_Prefix_Policy', templateVars['BGP_Peer_Prefix_Policy'])
            validating.values('Local_ASN_Config', templateVars['Local_ASN_Config'], ['dual-as', 'no-prepend', 'none', 'replace-as'])
            if not templateVars['Local_ASN'] == None:
                validating.number_check('Local_ASN', templateVars['Local_ASN'], 1, 4294967295)
            validating.values('Admin_State', templateVars['Admin_State'], ['disabled', 'enabled'])
            validating.validator('BGP_Peer_Policy', templateVars['BGP_Peer_Policy'])

            validating.validator(row_path, kwargs['ws'], 'tenant', templateVars['tenant'])
            validating.validator(row_path, kwargs['ws'], 'L3Out', templateVars['L3Out'])
            validating.validator(row_path, kwargs['ws'], 'Node_Profile', templateVars['Node_Profile'])
            validating.validator(row_path, kwargs['ws'], 'Interface_Profile', templateVars['Interface_Profile'])
            validating.values(row_path, kwargs['ws'], 'Interface_Type', templateVars['Interface_Type'], ['ext-svi', 'l3-port', 'sub-interface'])
            validating.number_check(row_path, kwargs['ws'], 'Pod_ID', templateVars['Pod_ID'], 1, 15)
            validating.number_check(row_path, kwargs['ws'], 'Node1_ID', templateVars['Node1_ID'], 101, 4001)
            if not templateVars['Node2_ID'] == None:
                validating.number_check(row_path, kwargs['ws'], 'Node2_ID', templateVars['Node2_ID'], 101, 4001)

            validating.number_check(row_bgp, ws_net, 'allowed_self_as_count', templateVars['allowed_self_as_count'], 1, 10)
            validating.values(row_bgp, ws_net, 'allow_self_as', templateVars['allow_self_as'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'as_override', templateVars['as_override'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'disable_peer_as_check', templateVars['disable_peer_as_check'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'next_hop_self', templateVars['next_hop_self'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'send_community', templateVars['send_community'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'send_ext_community', templateVars['send_ext_community'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'bfd', templateVars['bfd'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'AF_Mcast', templateVars['AF_Mcast'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'AF_Ucast', templateVars['AF_Ucast'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'remove_all_private_as', templateVars['remove_all_private_as'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'remove_private_as', templateVars['remove_private_as'], ['no', 'yes'])
            validating.values(row_bgp, ws_net, 'private_to_local', templateVars['private_to_local'], ['no', 'yes'])

            if not templateVars['BGP_Peer_Prefix_Policy'] == None:
                validating.validator(row_pfx, ws_net, 'Prefix_tenant', templateVars['Prefix_tenant'])
                validating.number_check(row_pfx, ws_net, 'Maximum_Prefixes', templateVars['Maximum_Prefixes'], 1, 300000)
                validating.number_check(row_pfx, ws_net, 'Restart_Time', templateVars['Restart_Time'], 1, 65535)
                validating.number_check(row_pfx, ws_net, 'Threshold', templateVars['Threshold'], 1, 100)
                validating.values(row_pfx, ws_net, 'Action', templateVars['Action'], ['log', 'reject', 'restart', 'shut'])
                if not templateVars['PFX_description'] == None:
                    validating.validator(row_pfx, ws_net, 'PFX_description', templateVars['PFX_description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        if re.search(r'\.', templateVars['Peer_Address']):
            templateVars['Peer_Address_'] = templateVars['Peer_Address'].replace('.', '-')
        else:
            templateVars['Peer_Address_'] = templateVars['Peer_Address'].replace(':', '-')

        ctrl_count = 0
        Ctrl = ''
        if templateVars['allow_self_as'] == 'yes':
            Ctrl = 'allow-self-as'
            ctrl_count =+ 1
        if templateVars['as_override'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'as-override'
            ctrl_count =+ 1
        elif templateVars['as_override'] == 'yes':
            Ctrl = 'as-override'
            ctrl_count =+ 1
        if templateVars['disable_peer_as_check'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'dis-peer-as-check'
            ctrl_count =+ 1
        elif templateVars['disable_peer_as_check'] == 'yes':
            Ctrl = 'dis-peer-as-check'
            ctrl_count =+ 1
        if templateVars['next_hop_self'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'nh-self'
            ctrl_count =+ 1
        elif templateVars['next_hop_self'] == 'yes':
            Ctrl = 'nh-self'
            ctrl_count =+ 1
        if templateVars['send_community'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'send-com'
            ctrl_count =+ 1
        elif templateVars['send_community'] == 'yes':
            Ctrl = 'send-com'
            ctrl_count =+ 1
        if templateVars['send_ext_community'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'send-ext-com'
            ctrl_count =+ 1
        elif templateVars['send_ext_community'] == 'yes':
            Ctrl = 'send-ext-com'
            ctrl_count =+ 1
        if ctrl_count > 0:
            templateVars['Ctrl'] = '%s' % (Ctrl)
        else:
            templateVars['Ctrl'] = ''

        ctrl_count = 0
        Ctrl = ''
        if templateVars['bfd'] == 'yes':
            Ctrl = 'bfd'
            ctrl_count =+ 1
        if templateVars['disable_connected_check'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'dis-conn-check'
            ctrl_count =+ 1
        elif templateVars['disable_connected_check'] == 'yes':
            Ctrl = 'dis-conn-check'
            ctrl_count =+ 1
        if ctrl_count > 0:
            templateVars['Peer_Ctrl'] = '%s' % (Ctrl)
        else:
            templateVars['Peer_Ctrl'] = ''

        ctrl_count = 0
        Ctrl = ''
        if templateVars['AF_Mcast'] == 'yes':
            Ctrl = 'af-mcast'
            ctrl_count =+ 1
        if templateVars['AF_Ucast'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'af-ucast'
            ctrl_count =+ 1
        elif templateVars['AF_Ucast'] == 'yes':
            Ctrl = 'af-ucast'
            ctrl_count =+ 1
        if ctrl_count > 0:
            templateVars['Address_Fam_Ctrl'] = '%s' % (Ctrl)
        else:
            templateVars['Address_Fam_Ctrl'] = ''

        ctrl_count = 0
        Ctrl = ''
        if templateVars['remove_all_private_as'] == 'yes':
            Ctrl = 'remove-all'
            ctrl_count =+ 1
        if templateVars['remove_private_as'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'remove-exclusive'
            ctrl_count =+ 1
        elif templateVars['remove_private_as'] == 'yes':
            Ctrl = 'remove-exclusive'
            ctrl_count =+ 1
        if templateVars['private_to_local'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + 'replace-as'
            ctrl_count =+ 1
        elif templateVars['private_to_local'] == 'yes':
            Ctrl = 'replace-as'
            ctrl_count =+ 1
        if ctrl_count > 0:
            templateVars['Private_AS_Ctrl'] = '%s' % (Ctrl)
        else:
            templateVars['Private_AS_Ctrl'] = ''

        if not templateVars['BGP_Password'] == None:
            x = templateVars['BGP_Password'].split('r')
            key_number = x[1]
            templateVars['sensitive_var'] = 'BGP_Password%s' % (key_number)

            # Define the Template Source
            template_file = "variables.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'variable_%s.tf' % (templateVars['sensitive_var'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            

            sensitive_var_site_group(dest_dir, dest_file, template, **templateVars)

        # Create Global Variables for First Template
        if not templateVars['Node2_ID'] == None:
            templateVars['PATH'] = 'protpaths-%s-%s' % (templateVars['Node1_ID'], templateVars['Node2_ID'])
        else:
            templateVars['PATH'] = 'paths-%s' % (templateVars['Node1_ID'])

        # Define the Template Source
        if templateVars['PATH'] == 'Loopback':
            template_file = "bgp_peer_connectivity_profile.jinja2"
        else:
            if templateVars['Local_ASN_Config'] == 'none':
                templateVars['Local_ASN_Config'] == None
            template_file = "bgp_peer_interface.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        if not templateVars['BGP_Peer_Prefix_Policy'] == None:

            # Define the Template Source
            if templateVars['Restart_Time'] == 65535:
                templateVars['Restart_Time'] = 'infinite'
            template_file = "bgp_peer_prefix.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'BGP_Peer_Prefix_%s.tf' % (templateVars['Policy_Name'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            print(f'dest file {dest_file} and dest_dir {dest_dir} and tenant = {templateVars["tenant"]}')
            

            if not templateVars['tenant'] == templateVars['Prefix_tenant']:
                # Define the Template Source
                template_file = "data_tenant.jinja2"
                template = self.templateEnv.get_template(template_file)

                temp_tenant = templateVars['tenant']
                templateVars['tenant'] = templateVars['Prefix_tenant']
                # Process the template through the Sites
                dest_file = 'data_tenant_%s.tf' % (templateVars['tenant'])
                dest_dir = 'tenant_%s' % (temp_tenant)
                

                templateVars['tenant'] = temp_tenant

                # Define the Template Source
                template_file = "data_bgp_peer_prefix.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'data_BGP_Peer_Prefix_%s.tf' % (templateVars['Policy_Name'])
                dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - Application Profiles
    #======================================================
    def bgp_pfx(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.policies.bgpPrefix']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'bgp_peer_prefix_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Contracts
    #======================================================
    def contract_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.contracts']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('tenant', templateVars['tenant'])
            validating.validator('contract', templateVars['contract'])
            validating.values('target_dscp', jsonData, **kwargs)
            validating.values('qos_class', jsonData, **kwargs)
            validating.values('contract_type', templateVars['contract_type'], ['OOB', 'Standard', 'Taboo'])
            validating.values('scope', templateVars['scope'], ['application-profile', 'context', 'global', 'tenant'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Define the Template Source
        if templateVars['contract_type'] == 'OOB':
            template_file = "contract_oob.jinja2"
        elif templateVars['contract_type'] == 'Standard':
            template_file = "contract.jinja2"
        elif templateVars['contract_type'] == 'Taboo':
            template_file = "contract_taboo.jinja2"
        dest_file = 'contract_type_%s_%s.tf' % (templateVars['contract_type'], templateVars['contract'])
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - Add Contract to EPG
    #======================================================
    def contract_to_epg(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'contract_tenant': '',
                         'contract_type': '',
                         'contract': '',
                         'tenant': '',
                         'App_Profile': '',
                         'EPG': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('contract_tenant', templateVars['contract_tenant'])
            validating.validator('contract', templateVars['contract'])
            validating.validator('tenant', templateVars['tenant'])
            validating.validator('App_Profile', templateVars['App_Profile'])
            validating.validator('EPG', templateVars['EPG'])
            validating.values('contract_type', templateVars['contract_type'], ['consumer', 'provider'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "epg_to_contract.jinja2"
        dest_file = 'App_Profile_%s_EPG_%s.tf' % (templateVars['App_Profile'], templateVars['EPG'])
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - Application Profiles
    #======================================================
    def eigrp_interface(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.policies.eigrpInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'eigrp_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Application EPG
    #======================================================
    def epg_add(self, **kwargs):
        # Open the Network Policies Worksheet
        ws_net = kwargs['wb']['Network Policies']
        rows = ws_net.max_row

        # Get the EPG Policies from the Network Policies Tab
        func = 'epg'
        count = countKeys(ws_net, func)
        row_net = ''
        var_dict = findVars(ws_net, func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('policy_name') == kwargs.get('epg_policy'):
                row_net = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        if kwargs['custom_qos'] == 'default':
            kwargs['custom_qos'] = 'uni/tn-common/qoscustom-default'
        if kwargs['data_plane_policer'] == 'default':
            kwargs['data_plane_policer'] = 'uni/tn-common/qosdpppol-default'
        if kwargs['fhs_trust_control_policy'] == 'default':
            kwargs['fhs_trust_control_policy'] = 'uni/tn-common/trustctrlpol-default'

        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationEpgs']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('application_profile', **kwargs)
            validating.validator('bridge_domain', **kwargs)
            validating.validator('name', **kwargs)
            validating.validator('tenant', **kwargs)
            if not kwargs['alias'] == None:
                validating.validator('alias', **kwargs)
            if not kwargs['description'] == None:
                validating.validator('description', **kwargs)
            if not kwargs['annotations'] == None:
                validating.validator_array('annotations', **kwargs)
            if not kwargs['physical_domains'] == None:
                if re.match(',', kwargs['physical_domains']):
                    validating.validator_list('physical_domains', **kwargs)
                else:
                    validating.validator('physical_domains', **kwargs)
            if not templateVars['vmm_domains'] == None:
                if re.match(',', templateVars['vmm_domains']):
                    validating.validator_list('vmm_domains', **kwargs)
                else:
                    validating.validator('vmm_domains', **kwargs)
            if not templateVars['epg_to_aaeps'] == None:
                validating.validator('epg_to_aaeps', templateVars['epg_to_aep'])
            if not templateVars['vlans'] == None:
                validating.vlans('vlans', templateVars['vlans'])
            row_num = kwargs['row_num']
            ws = kwargs['ws']
            kwargs['row_num'] = row_net
            kwargs['ws'] = ws_net
            validating.values('qos_class', jsonData, **kwargs)
            validating.values('epg_admin_state', jsonData, **kwargs)
            validating.values('flood_in_encapsulation', jsonData, **kwargs)
            validating.values('intra_epg_isolation', jsonData, **kwargs)
            validating.values('label_match_criteria', jsonData, **kwargs)
            validating.values('preferred_group_member', jsonData, **kwargs)
            validating.values('useg_epg', jsonData, **kwargs)
            kwargs['row_num'] = row_num
            kwargs['ws'] = ws
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        wb = kwargs['wb']
        if not templateVars['VLAN'] == None:
            # Define the Template Source
            template_file = "static_path.jinja2"
            template = self.templateEnv.get_template(template_file)

            dest_file = 'App_Profile_%s_EPG_%s.tf' % (templateVars['App_Profile'], templateVars['EPG'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            process_workbook('a+', dest_dir, dest_file, template, **templateVars)

        if not templateVars['epg_to_aep'] == None:
            if re.search(',', templateVars['epg_to_aep']):
                # Define the Template Source
                aep_list = templateVars['epg_to_aep'].split(',')
                for aep in aep_list:
                    templateVars['AAEP'] = aep

                    # Define the Template Source
                    template_file = "policies_global_aep_generic.jinja2"
                    template = self.templateEnv.get_template(template_file)

                    # Process the template through the Sites
                    dest_file = 'Policies_Global_AEP_%s_generic.tf' % (templateVars['AAEP'])
                    dest_dir = 'Access'
                    

                    # Define the Template Source
                    template_file = "data_access_generic.jinja2"
                    template = self.templateEnv.get_template(template_file)

                    # Process the template through the Sites
                    dest_file = 'data_AEP_%s.tf' % (templateVars['AAEP'])
                    dest_dir = 'tenant_%s' % (templateVars['tenant'])
                    

                    # Define the Template Source
                    template_file = "epgs_using_function.jinja2"
                    template = self.templateEnv.get_template(template_file)

                    # Process the template through the Sites
                    dest_file = 'App_Profile_%s_EPG_%s.tf' % (templateVars['App_Profile'], templateVars['EPG'])
                    dest_dir = 'tenant_%s' % (templateVars['tenant'])
                    
            else:
                templateVars['AAEP'] = templateVars['EPG_to_AAEP']
                # Define the Template Source
                template_file = "policies_global_aep_generic.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'Policies_Global_AEP_%s_generic.tf' % (templateVars['AAEP'])
                dest_dir = 'Access'
                

                # Define the Template Source
                template_file = "data_access_generic.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'data_AEP_%s.tf' % (templateVars['AAEP'])
                dest_dir = 'tenant_%s' % (templateVars['tenant'])
                

                # Define the Template Source
                template_file = "epgs_using_function.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'App_Profile_%s_EPG_%s.tf' % (templateVars['App_Profile'], templateVars['EPG'])
                dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - L3Out - External EPG
    #======================================================
    def ext_epg(self, **kwargs):
        # Open the Network Policies Worksheet
        ws_net = kwargs['wb']['Network Policies']
        rows = ws_net.max_row

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'tenant': '',
                         'L3Out': '',
                         'Ext_EPG': '',
                         'Ext_EPG_Policy': '',
                         'Subnets': '',
                         'Ext_Subnet_Policy': '',
                         'qos_class': '',
                         'target_dscp': '',
                         'pref_gr_memb': '',
                         'match_t': '',
                         'flood': '',
                         'export-rtctrl': '',
                         'import-rtctrl': '',
                         'import-security': '',
                         'shared-security': '',
                         'shared-rtctrl': '',
                         'agg-export': '',
                         'agg-import': '',
                         'agg-shared': ''}
        optional_args = {'alias': '',
                         'description': '',
                         'annotation': '',
                         'cons_vzBrCP': '',
                         'vzCPIf': '',
                         'Master_fvEPg': '',
                         'prov_vzBrCP': '',
                         'vzTaboo': '',
                         'exception_tag': '',
                         'rtctrlProfile': '',
                         'sub_rtctrlProfile': '',
                         'rtsumARtSummPol': ''}


        # Get the L3Out Policies from the Network Policies Tab
        func = 'ext_epg'
        count = countKeys(ws_net, func)
        row_epg = ''
        var_dict = findVars(ws_net, func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('Policy_Name') == kwargs.get('Ext_EPG_Policy'):
                row_epg = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        func = 'ext_subnet'
        count = countKeys(ws_net, func)
        row_sub = ''
        var_dict = findVars(ws_net, func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('Policy_Name') == kwargs.get('Ext_Subnet_Policy'):
                row_sub = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('tenant', templateVars['tenant'])
            if not templateVars['Subnets'] == None:
                if re.search(',', templateVars['Subnets']):
                    sx = templateVars['Subnets'].split(',')
                    for x in sx:
                        validating.ip_address('Subnets', x)
                else:
                    validating.ip_address('Subnets', templateVars['Subnets'])
            validating.dscp(row_epg, ws_net, 'target_dscp', templateVars['target_dscp'])
            validating.match_t(row_epg, ws_net, 'match_t', templateVars['match_t'])
            validating.values(row_epg, ws_net, 'qos_class', templateVars['qos_class'])
            validating.values(row_epg, ws_net, 'flood', templateVars['flood'], ['disabled', 'enabled'])
            validating.values(row_epg, ws_net, 'pref_gr_memb', templateVars['pref_gr_memb'], ['exclude', 'include'])
            validating.values(row_sub, ws_net, 'agg-export', templateVars['agg-export'], ['no', 'yes'])
            validating.values(row_sub, ws_net, 'agg-import', templateVars['agg-import'], ['no', 'yes'])
            validating.values(row_sub, ws_net, 'agg-shared', templateVars['agg-shared'], ['no', 'yes'])
            validating.values(row_sub, ws_net, 'export-rtctrl', templateVars['export-rtctrl'], ['no', 'yes'])
            validating.values(row_sub, ws_net, 'import-rtctrl', templateVars['import-rtctrl'], ['no', 'yes'])
            validating.values(row_sub, ws_net, 'import-security', templateVars['import-security'], ['no', 'yes'])
            validating.values(row_sub, ws_net, 'shared-security', templateVars['shared-security'], ['no', 'yes'])
            validating.values(row_sub, ws_net, 'shared-rtctrl', templateVars['shared-rtctrl'], ['no', 'yes'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Create aggregate templateVars
        aggregate = ''
        aggregate_count = 0
        if templateVars['agg-export'] == 'yes':
            aggregate = '"agg-export"'
            aggregate_count =+ 1
        if templateVars['agg-import'] == 'yes':
            if aggregate_count == 0:
                aggregate = '"agg-import"'
                aggregate_count =+ 1
            else:
                aggregate = aggregate + ', ' + '"agg-import"'
                aggregate_count =+ 1
        if templateVars['agg-shared'] == 'yes':
            if aggregate_count == 0:
                aggregate = '"agg-import"'
                aggregate_count =+ 1
            else:
                aggregate = aggregate + ', ' + '"agg-shared"'
                aggregate_count =+ 1

        if aggregate_count == 0:
            templateVars['aggregate'] = None
        else:
            templateVars['aggregate'] = '[%s]' % (aggregate)

        # Create scope templateVars
        scope = ''
        scope_count = 0
        if templateVars['export-rtctrl'] == 'yes':
            scope = '"export-rtctrl"'
            scope_count =+ 1
        if templateVars['import-rtctrl'] == 'yes':
            if scope_count == 0:
                scope = '"import-rtctrl"'
                scope_count =+ 1
            else:
                scope = scope + ', ' + '"import-rtctrl"'
                scope_count =+ 1
        if templateVars['import-security'] == 'yes':
            if scope_count == 0:
                scope = '"import-security"'
                scope_count =+ 1
            else:
                scope = scope + ', ' + '"import-security"'
                scope_count =+ 1
        if templateVars['shared-security'] == 'yes':
            if scope_count == 0:
                scope = '"shared-security"'
                scope_count =+ 1
            else:
                scope = scope + ', ' + '"shared-security"'
                scope_count =+ 1
        if templateVars['shared-rtctrl'] == 'yes':
            if scope_count == 0:
                scope = '"shared-rtctrl"'
                scope_count =+ 1
            else:
                scope = scope + ', ' + '"shared-rtctrl"'
                scope_count =+ 1

        if scope_count == 0:
            templateVars['scope'] = None
        else:
            templateVars['scope'] = '[%s]' % (scope)

        # Define the Template Source
        template_file = "epg_ext.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'L3Out_%s_EXTERNAL_EPG_%s.tf' % (templateVars['L3Out'], templateVars['Ext_EPG'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        if re.search(',', templateVars['Subnets']):
            sx = templateVars['Subnets'].split(',')
            for x in sx:
                templateVars['Subnet'] = x
                if re.search(':', x):
                    templateVars['Subnet_'] = x.replace(':', '-')
                    templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')
                else:
                    templateVars['Subnet_'] = x.replace('.', '-')
                    templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')

                # Define the Template Source
                template_file = "ext_subnet.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'L3Out_%s_EXTERNAL_EPG_%s.tf' % (templateVars['L3Out'], templateVars['Ext_EPG'])
                dest_dir = 'tenant_%s' % (templateVars['tenant'])
                
        else:
            templateVars['Subnet'] = templateVars['Subnets']
            if re.search(':', templateVars['Subnet']):
                templateVars['Subnet_'] = templateVars['Subnet'].replace(':', '-')
                templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')
            else:
                templateVars['Subnet_'] = templateVars['Subnet'].replace('.', '-')
                templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')

            # Define the Template Source
            template_file = "ext_subnet.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'L3Out_%s_EXTERNAL_EPG_%s.tf' % (templateVars['L3Out'], templateVars['Ext_EPG'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - L3Out - Out-of-Band External EPG
    #======================================================
    def ext_epg_oob(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationProfile']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Open the Network Policies Worksheet
        ws_net = kwargs['wb']['Network Policies']
        rows = ws_net.max_row

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'Ext_EPG': '',
                         'qos_class': '',
                         'Subnets': ''}
        optional_args = {'annotation': '',
                         'consumed_contracts': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('Ext_EPG', templateVars['Ext_EPG'])
            validating.values('qos_class', jsonData, **kwargs)
            if not templateVars['annotation'] == None:
                if re.search(',', templateVars['annotation']):
                    for x in templateVars['annotation'].split(','):
                        validating.validator('annotation', x)
                else:
                    validating.validator('annotation', templateVars['annotation'])
            if not templateVars['consumed_contracts'] == None:
                if re.search(',', templateVars['consumed_contracts']):
                    templateVars['provide_count'] =+ 1
                    for x in templateVars['consumed_contracts'].split(','):
                        validating.validator('consumed_contracts', x)
                else:
                    validating.validator('consumed_contracts', templateVars['consumed_contracts'])
            if not templateVars['Subnet'] == None:
                if re.search(',', templateVars['Subnets']):
                    sx = templateVars['Subnets'].split(',')
                    for x in sx:
                        validating.ip_address('Subnets', x)
                else:
                    validating.ip_address('Subnets', templateVars['Subnets'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "epg_ext_oob.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'EPG_Mgmt_OOB_External_EPG_%s.tf' % (templateVars['Ext_EPG'])
        dest_dir = 'tenant_mgmt'
        

        if re.search(',', templateVars['Subnets']):
            sx = templateVars['Subnets'].split(',')
            for x in sx:
                templateVars['Subnet'] = x
                if re.search(':', x):
                    templateVars['Subnet_'] = x.replace(':', '-')
                    templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')
                else:
                    templateVars['Subnet_'] = x.replace('.', '-')
                    templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')

                # Define the Template Source
                template_file = "ext_subnet_oob.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'EPG_Mgmt_OOB_External_EPG_%s.tf' % (templateVars['Ext_EPG'])
                dest_dir = 'tenant_mgmt'
                
        else:
            templateVars['Subnet'] = templateVars['Subnets']
            if re.search(':', templateVars['Subnet']):
                templateVars['Subnet_'] = templateVars['Subnet'].replace(':', '-')
                templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')
            else:
                templateVars['Subnet_'] = templateVars['Subnet'].replace('.', '-')
                templateVars['Subnet_'] = templateVars['Subnet_'].replace('/', '_')

            # Define the Template Source
            template_file = "ext_subnet_oob.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'EPG_Mgmt_OOB_External_EPG_%s.tf' % (templateVars['Ext_EPG'])
            dest_dir = 'tenant_mgmt'

    #======================================================
    # Function - Contract Filter
    #======================================================
    def filter_add(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'tenant': '',
                         'Filter': ''}
        optional_args = {'description': '',
                         'alias': '',
                         'annotation': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('Filter', templateVars['Filter'])
            validating.validator('tenant', templateVars['tenant'])
            if not templateVars['alias'] == None:
                validating.validator('alias', templateVars['alias'])
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "contract_filter.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'contract_Filter_%s.tf' % (templateVars['Filter'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - Contract Filter - Filter Entry
    #======================================================
    def filter_entry(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.filterEntry']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        try:
            # Validate Required Arguments
            validating.filter_ports('destination_from', **kwargs)
            validating.filter_ports('destination_to', **kwargs)
            validating.filter_ports('source_from', **kwargs)
            validating.filter_ports('source_to', **kwargs)
            validating.site_group('site_group', **kwargs)
            validating.validator('filter_name', **kwargs)
            validating.validator('name', **kwargs)
            validating.validator('tenant', **kwargs)
            validating.values('acknowledgement', jsonData, **kwargs)
            validating.values('arp_flag', jsonData, **kwargs)
            validating.values('established', jsonData, **kwargs)
            validating.values('ethertype', jsonData, **kwargs)
            validating.values('finish', jsonData, **kwargs)
            validating.values('icmpv4_type', jsonData, **kwargs)
            validating.values('icmpv6_type', jsonData, **kwargs)
            validating.values('ip_protocol', jsonData, **kwargs)
            validating.values('match_dscp', jsonData, **kwargs)
            validating.values('match_only_frags', jsonData, **kwargs)
            validating.values('reset', jsonData, **kwargs)
            validating.values('stateful', jsonData, **kwargs)
            validating.values('synchronize', jsonData, **kwargs)
            if not templateVars['alias'] == None:
                validating.validator('alias', **kwargs)
            if not templateVars['description'] == None:
                validating.validator('description', **kwargs)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Add Dictionary to easyDict
        for items in kwargs['easyDict']['tenants']['filters']:
            for k, v in items.items():
                if k == kwargs['site_group']:
                    for i in v:
                        i['filter_entries'].append(templateVars)

        # Return Dictionary
        return kwargs['easyDict']

    #======================================================
    # Function - L3Out
    #======================================================
    def l3out_add(self, **kwargs):
        # Open the Network Policies Worksheet
        ws_net = kwargs['wb']['Network Policies']
        rows = ws_net.max_row

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'tenant': '',
                         'L3Out': '',
                         'VRF_tenant': '',
                         'VRF': '',
                         'L3_Domain': '',
                         'target_dscp': '',
                         'Run_BGP': '',
                         'export': '',
                         'import': '',
                         'target_dscp': ''}
        optional_args = {'description': '',
                         'alias': '',
                         'annotation': '',
                         'EIGRP_Routing_Policy': '',
                         'OSPF_Routing_Policy': '',
                         'leak_rtctrlProfile': '',
                         'damp_rtctrlProfile': '',
                         'fvBDPublicSubnetHolder': ''}


        # Get the L3Out Policies from the Network Policies Tab
        func = 'l3Out'
        count = countKeys(ws_net, func)
        row_l3out = ''
        var_dict = findVars(ws_net, func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('Policy_Name') == kwargs.get('L3Out_Policy'):
                row_l3out = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.dscp(row_l3out, ws_net, 'target_dscp', templateVars['target_dscp'])
            validating.validator('tenant', templateVars['tenant'])
            validating.validator('L3Out', templateVars['L3Out'])
            validating.validator('VRF', templateVars['VRF'])
            validating.validator('VRF_tenant', templateVars['VRF_tenant'])
            validating.values('export', templateVars['export'], ['no', 'yes'])
            validating.values('import', templateVars['import'], ['no', 'yes'])
            validating.values('Run_BGP', templateVars['Run_BGP'], ['no', 'yes'])
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Process the template through the Sites
        template_file = "data_domain_l3_profile.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'data_domain_l3_profile_%s.tf' % (templateVars['L3_Domain'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        ctrl_count = 0
        Ctrl = ''
        if templateVars['export'] == 'yes':
            Ctrl = '"export"'
            ctrl_count =+ 1
        if templateVars['import'] == 'yes' and ctrl_count > 0:
            Ctrl = Ctrl + ',' + '"import"'
            ctrl_count =+ 1
        elif templateVars['import'] == 'yes':
            Ctrl = '"import"'
            ctrl_count =+ 1
        if ctrl_count > 0:
            templateVars['enforce_rtctrl'] = '[%s]' % (Ctrl)
        else:
            templateVars['enforce_rtctrl'] = '["unspecified"]'

        # Define the Template Source
        template_file = "l3out.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'L3Out_%s.tf' % (templateVars['L3Out'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        if templateVars['Run_BGP'] == 'yes':
            # Define the Template Source
            template_file = "bgp_external_policy.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'L3Out_%s.tf' % (templateVars['L3Out'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            

        #--------------------------
        # EIGRP Routing Policy
        #--------------------------
        if not templateVars['EIGRP_Routing_Policy'] == None:

            # Dicts for required and optional args
            required_args = {'site_group': '',
                            'tenant': '',
                            'L3Out': '',
                            'Policy_Name': '',
                            'AS_Number': ''}
            optional_args = { }
            # Get the L3Out Policies from the Network Policies Tab
            func = 'eigrp_routing'
            count = countKeys(ws_net, func)
            row_eigrp = ''
            var_dict = findVars(ws_net, func, rows, count)
            for pos in var_dict:
                if var_dict[pos].get('Policy_Name') == kwargs.get('EIGRP_Routing_Policy'):
                    row_eigrp = var_dict[pos]['row']
                    del var_dict[pos]['row']
                    kwargs = {**kwargs, **var_dict[pos]}

            # Validate inputs, return dict of template vars
            templateVars = process_kwargs(required_args, optional_args, **kwargs)

            try:
                # Validate Required Arguments
                validating.number_check(row_eigrp, ws_net, 'AS_Number', templateVars['AS_Number'], 1, 65534)
            except Exception as err:
                errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws_net, row_eigrp)
                raise ErrException(errorReturn)

            # Define the Template Source
            template_file = "l3out_eigrp_external_policy.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'L3Out_%s.tf' % (templateVars['L3Out'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            

        #--------------------------
        # OSPF Routing Policy
        #--------------------------
        if not templateVars['OSPF_Routing_Policy'] == None:

            # Dicts for required and optional args
            required_args = {'site_group': '',
                            'tenant': '',
                            'L3Out': '',
                            'Policy_Name': '',
                            'Area_ID': '',
                            'Area_Type': '',
                            'Cost': '',
                            'Redistribute_NSSA': '',
                            'Originate_Summary': '',
                            'Suppress_FA': ''}
            optional_args = { }
            # Get the L3Out Policies from the Network Policies Tab
            func = 'ospf_routing'
            count = countKeys(ws_net, func)
            row_ospf = ''
            var_dict = findVars(ws_net, func, rows, count)
            for pos in var_dict:
                if var_dict[pos].get('Policy_Name') == kwargs.get('OSPF_Routing_Policy'):
                    row_ospf = var_dict[pos]['row']
                    del var_dict[pos]['row']
                    kwargs = {**kwargs, **var_dict[pos]}

            # Validate inputs, return dict of template vars
            templateVars = process_kwargs(required_args, optional_args, **kwargs)

            try:
                # Validate Required Arguments
                validating.number_check(row_ospf, ws_net, 'Cost', templateVars['Cost'], 0, 16777215)
                validating.values(row_ospf, ws_net, 'Area_Type', templateVars['Area_Type'], ['nssa', 'regular', 'stub'])
                validating.values(row_ospf, ws_net, 'Redistribute_NSSA', templateVars['Redistribute_NSSA'], ['no', 'yes'])
                validating.values(row_ospf, ws_net, 'Originate_Summary', templateVars['Originate_Summary'], ['no', 'yes'])
                validating.values(row_ospf, ws_net, 'Suppress_FA', templateVars['Suppress_FA'], ['no', 'yes'])
            except Exception as err:
                errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws_net, row_ospf)
                raise ErrException(errorReturn)

            ctrl_count = 0
            Ctrl = ''
            if templateVars['Redistribute_NSSA'] == 'yes':
                Ctrl = 'redistribute'
                ctrl_count =+ 1
            if templateVars['Originate_Summary'] == 'yes' and ctrl_count > 0:
                Ctrl = Ctrl + ',' + 'summary'
                ctrl_count =+ 1
            elif templateVars['Originate_Summary'] == 'yes':
                Ctrl = 'summary'
                ctrl_count =+ 1
            if templateVars['Suppress_FA'] == 'yes' and ctrl_count > 0:
                Ctrl = Ctrl + ',' + 'supress-fa'
                ctrl_count =+ 1
            elif templateVars['Suppress_FA'] == 'yes':
                Ctrl = 'supress-fa'
                ctrl_count =+ 1
            if ctrl_count > 0:
                templateVars['Ctrl'] = '%s' % (Ctrl)
            else:
                templateVars['Ctrl'] = 'unspecified'

            # Define the Template Source
            template_file = "l3out_ospf_external_policy.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'L3Out_%s.tf' % (templateVars['L3Out'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - Management EPG
    #======================================================
    def mgmt_epg(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationProfile']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Dicts for Bridge Domain required and optional args
        required_args = {'site_group': '',
                         'Type': '',
                         'EPG': '',
                         'qos_class': ''}
        optional_args = {'annotation': '',
                         'VLAN': '',
                         'Bridge_Domain': '',
                         'tenant': '',
                         'consumed_contracts': '',
                         'provided_contracts': '',
                         'match_t': '',
                         'contract_Interfaces': '',
                         'Taboo_contracts': '',
                         'Subnets': '',
                         'Static_Routes': '',}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        templateVars['consume_count'] = 1
        templateVars['provide_count'] = 1
        templateVars['interface_count'] = 1
        templateVars['taboo_count'] = 1
        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('EPG', templateVars['EPG'])
            validating.values('qos_class', jsonData, **kwargs)
            validating.values('Type', templateVars['Type'], ['in_band', 'out_of_band'])
            if templateVars['Type'] == 'in_band':
                validating.vlans('VLAN', templateVars['VLAN'])
                validating.validator('Bridge_Domain', templateVars['Bridge_Domain'])
            if not templateVars['tenant'] == None:
                validating.validator('tenant', templateVars['tenant'])
                if re.search(',', templateVars['provided_contracts']):
                    templateVars['provide_count'] =+ 1
                    for x in templateVars['provided_contracts'].split(','):
                        validating.validator('provided_contracts', x)
                else:
                    validating.validator('provided_contracts', templateVars['provided_contracts'])
                if templateVars['Type'] == 'in_band':
                    if not templateVars['consumed_contracts'] == None:
                        if re.search(',', templateVars['consumed_contracts']):
                            templateVars['provide_count'] =+ 1
                            for x in templateVars['consumed_contracts'].split(','):
                                validating.validator('consumed_contracts', x)
                        else:
                            validating.validator('consumed_contracts', templateVars['consumed_contracts'])
                    if not templateVars['contract_Interfaces'] == None:
                        if re.search(',', templateVars['contract_Interfaces']):
                            for x in templateVars['contract_Interfaces'].split(','):
                                templateVars['interface_count'] =+ 1
                                validating.not_empty('contract_Interfaces', x)
                        else:
                            validating.not_empty('contract_Interfaces', templateVars['contract_Interfaces'])
                    if not templateVars['Taboo_contracts'] == None:
                        if re.search(',', templateVars['Taboo_contracts']):
                            templateVars['taboo_count'] =+ 1
                            for x in templateVars['Taboo_contracts'].split(','):
                                validating.not_empty('Taboo_contracts', x)
                        else:
                            validating.not_empty('Taboo_contracts', templateVars['Taboo_contracts'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        if not templateVars['tenant'] == 'mgmt':
            dest_dir = 'tenant_mgmt'

            template_file = 'data_tenant.jinja2'
            template = self.templateEnv.get_template(template_file)
            dest_file = 'data_tenant_%s.tf' % (templateVars['tenant'])
            

            templateVars['contract_type'] = 'Standard'
            if not templateVars['consumed_contracts'] == None:
                if re.search(',', templateVars['consumed_contracts']):
                    for x in templateVars['consumed_contracts'].split(','):
                        templateVars['contract'] = x
                        template_file = 'data_contract.jinja2'
                        template = self.templateEnv.get_template(template_file)
                        dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], x)
                        
                else:
                    templateVars['contract'] = templateVars['consumed_contracts']
                    template_file = 'data_contract.jinja2'
                    template = self.templateEnv.get_template(template_file)
                    dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], templateVars['consumed_contracts'])
                    

            if not templateVars['contract_Interfaces'] == None:
                if re.search(',', templateVars['contract_Interfaces']):
                    for x in templateVars['contract_Interfaces'].split(','):
                        templateVars['contract'] = x
                        template_file = 'data_contract.jinja2'
                        template = self.templateEnv.get_template(template_file)
                        dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], x)
                        
                else:
                    templateVars['contract'] = templateVars['contract_Interfaces']
                    template_file = 'data_contract.jinja2'
                    template = self.templateEnv.get_template(template_file)
                    dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], templateVars['contract_Interfaces'])
                    

            if not templateVars['provided_contracts'] == None:
                if templateVars['Type'] == 'in_band':
                    if re.search(',', templateVars['provided_contracts']):
                        for x in templateVars['provided_contracts'].split(','):
                            templateVars['contract'] = x
                            template_file = 'data_contract.jinja2'
                            template = self.templateEnv.get_template(template_file)
                            dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], x)
                            
                    else:
                        templateVars['contract'] = templateVars['provided_contracts']
                        template_file = 'data_contract.jinja2'
                        template = self.templateEnv.get_template(template_file)
                        dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], templateVars['provided_contracts'])
                        

            templateVars['contract_type'] = 'Taboo'
            if not templateVars['Taboo_contracts'] == None:
                if re.search(',', templateVars['Taboo_contracts']):
                    for x in templateVars['Taboo_contracts'].split(','):
                        templateVars['contract'] == x
                        template_file = 'data_contract_taboo.jinja2'
                        template = self.templateEnv.get_template(template_file)
                        dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], x)
                        
                else:
                    templateVars['contract'] = templateVars['Taboo_contracts']
                    template_file = 'data_contract_taboo.jinja2'
                    template = self.templateEnv.get_template(template_file)
                    dest_file = 'data_contract_type_%s_%s.tf' % (templateVars['contract_type'], templateVars['contract_type'])
                    

        # Define the Template Source and Destination File
        template_file = "epg_mgmt.jinja2"
        template = self.templateEnv.get_template(template_file)
        dest_file = 'EPG_Mgmt_Type_%s_EPG_%s.tf' % (templateVars['Type'], templateVars['EPG'])

        # Process the template through the Sites
        dest_dir = 'tenant_mgmt'
        

        # Define the Template Source and Destination File
        template_file = "var_mgmt.jinja2"
        template = self.templateEnv.get_template(template_file)

        if templateVars['Type'] == 'in_band':
            templateVars['var_name'] = 'in_band'
            dest_file = 'var_Mgmt_EPG_%s.tf' % ('inb')
        else:
            templateVars['var_name'] = 'out_of_band'
            dest_file = 'var_Mgmt_EPG_%s.tf' % ('oob')

        # Process the template through the Sites
        dest_dir = 'Access'
        
        dest_dir = 'Admin'
        
        dest_dir = 'Fabric'

    #======================================================
    # Function - L3Out - Node Interface
    #======================================================
    def node_intf(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationProfile']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Open the Network Policies Worksheet
        ws_net = kwargs['wb']['Network Policies']
        rows = ws_net.max_row

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'tenant': '',
                         'L3Out': '',
                         'Node_Profile': '',
                         'Interface_Profile': '',
                         'qos_class': '',
                         'Node_Intf_Policy': '',
                         'Policy_Name': '',
                         'tag': ''}
        optional_args = {'description': '',
                         'alias': '',
                         'EIGRP_Intf_Profile': '',
                         'OSPF_Intf_Profile': '',
                         'ndIfPol': '',
                         'egress_qosDppPol': '',
                         'ingress_qosDppPol': '',
                         'qosCustomPol': '',
                         'igmp_policy': '',
                         'netflowMonitorPol': ''}

        # Get the Node Policies from the Network Policies Tab
        func = 'node_intf'
        count = countKeys(ws_net, func)
        row_node = ''
        var_dict = findVars(ws_net, func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('Policy_Name') == kwargs.get('Node_Intf_Policy'):
                row_node = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('tenant', templateVars['tenant'])
            validating.validator('L3Out', templateVars['L3Out'])
            validating.validator('Node_Profile', templateVars['Node_Profile'])
            validating.validator('Interface_Profile', templateVars['Interface_Profile'])
            validating.values('qos_class', jsonData, **kwargs)
            if not templateVars['alias'] == None:
                validating.validator('alias', templateVars['alias'])
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
            if not templateVars['EIGRP_Intf_Profile'] == None:
                validating.validator('EIGRP_Intf_Profile', templateVars['EIGRP_Intf_Profile'])
            if not templateVars['OSPF_Intf_Profile'] == None:
                validating.validator('OSPF_Intf_Profile', templateVars['OSPF_Intf_Profile'])
            if not templateVars['tag'] == None:
                validating.tag_check(row_node, ws_net, 'tag', templateVars['tag'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "logical_interface_profile.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        #--------------------------
        # OSPF Interface Profile
        #--------------------------
        if not templateVars['OSPF_Intf_Profile'] == None:

            # Dicts for required and optional args
                             # OSPF Interface Profile
            required_args = {'site_group': '',
                             'tenant': '',
                             'L3Out': '',
                             'Node_Profile': '',
                             'Interface_Profile': '',
                             'Auth_Type': '',
                             'Interface_Policy_Name': '',
                             # OSPF Interface Policy
                             'OSPF_Policy_Name': '',
                             'Policy_tenant': '',
                             'Network_Type': '',
                             'Priority': '',
                             'Cost': '',
                             'Advertise_Subnet': '',
                             'BFD': '',
                             'MTU_Ignore': '',
                             'Passive_Interface': '',
                             'Hello_Interval': '',
                             'Dead_Interval': '',
                             'Retransmit_Interval': '',
                             'Transmit_Delay': ''}
                             # OSPF Interface Profile
            optional_args = {'Auth_Key_ID': '',
                             'OSPF_Auth_Key': '',
                             'description': '',
                             # OSPF Interface Policy
                             'OSPF_description': ''}

            # Get the OSPF Profile Attributes from the Network Policies Tab
            func = 'ospf_profile'
            count = countKeys(ws_net, func)
            row_ospf = ''
            var_dict = findVars(ws_net, func, rows, count)
            for pos in var_dict:
                if var_dict[pos].get('Policy_Name') == kwargs.get('OSPF_Intf_Profile'):
                    row_ospf = var_dict[pos]['row']
                    del var_dict[pos]['row']
                    kwargs = {**kwargs, **var_dict[pos]}

           # Get the OSPF Policy Attributes from the Network Policies Tab
            func = 'ospf_policy'
            count = countKeys(ws_net, func)
            row_intf = ''
            var_dict = findVars(ws_net, func, rows, count)
            for pos in var_dict:
                if var_dict[pos].get('OSPF_Policy_Name') == kwargs.get('Interface_Policy_Name'):
                    row_intf = var_dict[pos]['row']
                    del var_dict[pos]['row']
                    kwargs = {**kwargs, **var_dict[pos]}
            # print(kwargs)
            # Validate inputs, return dict of template vars
            templateVars = process_kwargs(required_args, optional_args, **kwargs)

            try:
                # Validate OSPF Profile Required Arguments
                validating.values(row_ospf, ws_net, 'Auth_Type', templateVars['Auth_Type'], ['md5', 'none', 'simple'])
                if not templateVars['Auth_Key_ID'] == None:
                    validating.number_check(row_ospf, ws_net, 'Auth_Key_ID', templateVars['Auth_Key_ID'], 1, 255)
                if not templateVars['description'] == None:
                    validating.validator(row_ospf, ws_net, 'description', templateVars['description'])
                # Validate OSPF Policy Required Arguments
                validating.validator(row_intf, ws_net, 'OSPF_Policy_Name', templateVars['OSPF_Policy_Name'])
                validating.number_check(row_intf, ws_net, 'Priority', templateVars['Priority'], 0, 255)
                validating.number_check(row_intf, ws_net, 'Cost', templateVars['Cost'], 0, 65535)
                validating.number_check(row_intf, ws_net, 'Hello_Interval', templateVars['Hello_Interval'], 1, 65535)
                validating.number_check(row_intf, ws_net, 'Dead_Interval', templateVars['Dead_Interval'], 1, 65535)
                validating.number_check(row_intf, ws_net, 'Retransmit_Interval', templateVars['Retransmit_Interval'], 1, 65535)
                validating.number_check(row_intf, ws_net, 'Transmit_Delay', templateVars['Transmit_Delay'], 1, 65535)
                validating.values(row_intf, ws_net, 'Network_Type', templateVars['Network_Type'], ['broadcast', 'p2p', 'unspecified'])
                validating.values(row_intf, ws_net, 'Advertise_Subnet', templateVars['Advertise_Subnet'], ['no', 'yes'])
                validating.values(row_intf, ws_net, 'BFD', templateVars['BFD'], ['no', 'yes'])
                validating.values(row_intf, ws_net, 'MTU_Ignore', templateVars['MTU_Ignore'], ['no', 'yes'])
                validating.values(row_intf, ws_net, 'Passive_Interface', templateVars['Passive_Interface'], ['no', 'yes'])
                if not templateVars['OSPF_description'] == None:
                    validating.validator(row_intf, ws_net, 'OSPF_description', templateVars['OSPF_description'])
            except Exception as err:
                errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (SystemExit(err), ws_net, row_ospf)
                raise ErrException(errorReturn)

            if templateVars['Auth_Type'] == 'none':
                templateVars['Auth_Key_ID'] = 1

            # Define the Template Source
            template_file = "l3out_ospf_interface_profile.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            

            if not templateVars['OSPF_Auth_Key'] == None:
                x = templateVars['OSPF_Auth_Key'].split('r')
                key_number = x[1]
                templateVars['sensitive_var'] = 'OSPF_Auth_Key%s' % (key_number)

                # Define the Template Source
                template_file = "variables.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'variable_%s.tf' % (templateVars['sensitive_var'])
                dest_dir = 'tenant_%s' % (templateVars['tenant'])
                

                sensitive_var_site_group(dest_dir, dest_file, template, **templateVars)

            #--------------------------
            # OSPF Interface Policy
            #--------------------------

            ctrl_count = 0
            Ctrl = ''
            if templateVars['Advertise_Subnet'] == 'yes':
                Ctrl = '"advert-subnet"'
                ctrl_count =+ 1
            if templateVars['BFD'] == 'yes' and ctrl_count > 0:
                Ctrl = Ctrl + ', ' + '"bfd"'
                ctrl_count =+ 1
            elif templateVars['BFD'] == 'yes':
                Ctrl = '"bfd"'
                ctrl_count =+ 1
            if templateVars['MTU_Ignore'] == 'yes' and ctrl_count > 0:
                Ctrl = Ctrl + ', ' + '"mtu-ignore"'
                ctrl_count =+ 1
            elif templateVars['MTU_Ignore'] == 'yes':
                Ctrl = '"mtu-ignore"'
                ctrl_count =+ 1
            if templateVars['Passive_Interface'] == 'yes' and ctrl_count > 0:
                Ctrl = Ctrl + ', ' + '"passive"'
                ctrl_count =+ 1
            elif templateVars['Passive_Interface'] == 'yes':
                Ctrl = '"passive"'
                ctrl_count =+ 1
            if ctrl_count > 0:
                templateVars['Ctrl'] = '[%s]' % (Ctrl)
            else:
                templateVars['Ctrl'] = '[unspecified]'

            if templateVars['Cost'] == 0:
                templateVars['Cost'] = 'unspecified'

            # Define the Template Source
            template_file = "ospf_interface_policy.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'Policies_OSPF_Interface_%s.tf' % (templateVars['OSPF_Policy_Name'])
            dest_dir = 'tenant_%s' % (templateVars['Policy_tenant'])
            

            if not templateVars['tenant'] == templateVars['Policy_tenant']:
                # Define the Template Source
                template_file = "data_ospf_interface_policy.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'data_Policies_OSPF_Interface_%s.tf' % (templateVars['OSPF_Policy_Name'])
                dest_dir = 'tenant_%s' % (templateVars['tenant'])
                

                templateVars['L3Out_tenant'] = templateVars['tenant']
                templateVars['tenant'] = templateVars['Policy_tenant']
                # Define the Template Source
                template_file = "data_tenant.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'data_tenant_%s.tf' % (templateVars['tenant'])
                dest_dir = 'tenant_%s' % (templateVars['L3Out_tenant'])

    #======================================================
    # Function - L3Out - Node Path
    #======================================================
    def node_path(self, **kwargs):
        # Dicts for required and optional args
                         # Logical Interface Profile
        required_args = {'site_group': '',
                         'Path_Policy_Name': '',
                         'MTU': '',
                         'target_dscp': '',
                         'SideA_Address': '',
                         'SideA_IPv6_DAD': '',
                         # L3Out Path Profile
                         'tenant': '',
                         'L3Out': '',
                         'Policy_Name': '',
                         'Node_Profile': '',
                         'Interface_Profile': '',
                         'Interface_Type': '',
                         'Pod_ID': '',
                         'Node1_ID': '',
                         'Interface_or_PG': ''}
                         # Logical Interface Profile
        optional_args = {'Encap_scope': '',
                         'Mode': '',
                         'VLAN': '',
                         'description': '',
                         'Auto_State': '',
                         'MAC_Address': '',
                         'SideA_Secondary': '',
                         'SideA_Link_Local': '',
                         'SideB_Address': '',
                         'SideB_IPv6_DAD': '',
                         'SideB_Secondary': '',
                         'SideB_Link_Local': '',
                         'Node2_ID': ''}

        # Get the Node Policies from the Network Policies Tab
        rows = kwargs['ws'].max_row
        func = 'l3out_path'
        count = countKeys(kwargs['ws'], func)
        row_path = ''
        var_dict = findVars(kwargs['ws'], func, rows, count)
        for pos in var_dict:
            if var_dict[pos].get('Policy_Name') == kwargs.get('Path_Policy_Name'):
                row_path = var_dict[pos]['row']
                del var_dict[pos]['row']
                kwargs = {**kwargs, **var_dict[pos]}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments Logincal Interface Profile
            validating.site_group('site_group', **kwargs)
            validating.validator(row_path, kwargs['ws'], 'Path_Policy_Name', templateVars['Path_Policy_Name'])
            if not templateVars['Encap_scope'] == None:
                validating.values('Encap_scope', templateVars['Encap_scope'], ['ctx', 'local'])
            if not templateVars['Mode'] == None:
                validating.values('Mode', templateVars['Mode'], ['native', 'regular', 'untagged'])
            if not templateVars['VLAN'] == None:
                validating.vlans('VLAN', templateVars['VLAN'])
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
            if not templateVars['Auto_State'] == None:
                validating.values('Auto_State', templateVars['Auto_State'], ['disabled', 'enabled'])
            if not templateVars['MTU'] == 'inherit':
                validating.number_check(row_path, kwargs['ws'], 'MTU', templateVars['MTU'], 1300, 9216)
            if not templateVars['MAC_Address'] == None:
                validating.mac_address('MAC_Address', templateVars['MAC_Address'])
            validating.dscp('target_dscp', templateVars['target_dscp'])
            validating.ip_address('SideA_Address', templateVars['SideA_Address'])
            validating.values('SideA_IPv6_DAD', templateVars['SideA_IPv6_DAD'], ['disabled', 'enabled'])
            if not templateVars['SideA_Secondary'] == None:
                validating.ip_address('SideA_Secondary', templateVars['SideA_Secondary'])
            if not templateVars['SideA_Link_Local'] == None:
                validating.ip_address('SideA_Link_Local', templateVars['SideA_Link_Local'])
            if not templateVars['SideB_Address'] == None:
                validating.ip_address('SideB_Address', templateVars['SideB_Address'])
            if not templateVars['SideB_Address'] == None:
                validating.values('SideB_IPv6_DAD', templateVars['SideB_IPv6_DAD'], ['disabled', 'enabled'])
            if not templateVars['SideB_Secondary'] == None:
                validating.ip_address('SideB_Secondary', templateVars['SideB_Secondary'])
            if not templateVars['SideB_Link_Local'] == None:
                validating.ip_address('SideB_Link_Local', templateVars['SideB_Link_Local'])
            validating.validator(row_path, kwargs['ws'], 'tenant', templateVars['tenant'])
            validating.validator(row_path, kwargs['ws'], 'L3Out', templateVars['L3Out'])
            validating.validator(row_path, kwargs['ws'], 'Node_Profile', templateVars['Node_Profile'])
            validating.validator(row_path, kwargs['ws'], 'Interface_Profile', templateVars['Interface_Profile'])
            validating.values(row_path, kwargs['ws'], 'Interface_Type', templateVars['Interface_Type'], ['ext-svi', 'l3-port', 'sub-interface'])
            validating.number_check(row_path, kwargs['ws'], 'Pod_ID', templateVars['Pod_ID'], 1, 15)
            validating.number_check(row_path, kwargs['ws'], 'Node1_ID', templateVars['Node1_ID'], 101, 4001)
            if not templateVars['Node2_ID'] == None:
                validating.number_check(row_path, kwargs['ws'], 'Node2_ID', templateVars['Node2_ID'], 101, 4001)
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Create Global Variables for First Template
        if not templateVars['Node2_ID'] == None:
            templateVars['Address'] = None
            templateVars['IPv6_DAD'] = templateVars['SideA_IPv6_DAD']
            templateVars['Link_Local'] = None
            templateVars['PATH'] = 'protpaths-%s-%s' % (templateVars['Node1_ID'], templateVars['Node2_ID'])
        else:
            templateVars['Address'] = templateVars['SideA_Address']
            templateVars['IPv6_DAD'] = templateVars['SideA_IPv6_DAD']
            templateVars['Link_Local'] = templateVars['SideA_Link_Local']
            templateVars['PATH'] = 'paths-%s' % (templateVars['Node1_ID'])
        # Define the Template Source
        template_file = "l3out_path_attachment.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        # Create Global Variables for First Template
        if not templateVars['Node2_ID'] == None:
            templateVars['PATH'] = 'protpaths-%s-%s' % (templateVars['Node1_ID'], templateVars['Node2_ID'])

            templateVars['Address'] = templateVars['SideA_Address']
            templateVars['IPv6_DAD'] = templateVars['SideA_IPv6_DAD']
            templateVars['Link_Local'] = templateVars['SideA_Link_Local']
            templateVars['Side'] = 'A'

            # Define the Template Source
            template_file = "l3out_path_attach_vpc.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            

            templateVars['Address'] = templateVars['SideB_Address']
            templateVars['IPv6_DAD'] = templateVars['SideB_IPv6_DAD']
            templateVars['Link_Local'] = templateVars['SideB_Link_Local']
            templateVars['Side'] = 'B'

            # Define the Template Source
            template_file = "l3out_path_attach_vpc.jinja2"
            template = self.templateEnv.get_template(template_file)

            # Process the template through the Sites
            dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])
            

        # Define the Template Source
        if not templateVars['SideA_Secondary'] == None:
            if templateVars['Node2_ID'] == None:
                templateVars['Secondary'] = templateVars['SideA_Secondary']
                templateVars['IPv6_DAD'] = templateVars['SideA_IPv6_DAD']
                templateVars['PATH'] = 'paths-%s' % (templateVars['Node1_ID'])
                template_file = "l3out_path_attach_secondary.jinja2"
                template = self.templateEnv.get_template(template_file)

                # Process the template through the Sites
                dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
                dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - L3Out - Node Profile
    #======================================================
    def node_prof(self, **kwargs):
        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'tenant': '',
                         'L3Out': '',
                         'Node_Profile': '',
                         'target_dscp': '',
                         'Color_Tag': '',
                         'Pod_ID': '',
                         'Node1_ID': '',
                         'Node1_Router_ID': '',
                         'Node1_Loopback': ''}
        optional_args = {'alias': '',
                         'description': '',
                         'Node2_ID': '',
                         'Node2_Router_ID': '',
                         'Node2_Loopback': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.validator('tenant', templateVars['tenant'])
            validating.validator('L3Out', templateVars['L3Out'])
            validating.validator('Node_Profile', templateVars['Node_Profile'])
            validating.dscp('target_dscp', templateVars['target_dscp'])
            validating.tag_check('Color_Tag', templateVars['Color_Tag'])
            validating.number_check('Node1_ID', templateVars['Node1_ID'], 101, 4001)
            validating.ip_address('Node1_Router_ID', templateVars['Node1_Router_ID'])
            validating.values('Node1_Loopback', templateVars['Node1_Loopback'], ['no', 'yes'])
            if not templateVars['alias'] == None:
                validating.validator('alias', templateVars['alias'])
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
            if not templateVars['Node2_ID'] == None:
                validating.number_check('Node2_ID', templateVars['Node2_ID'], 101, 4001)
                validating.ip_address('Node2_Router_ID', templateVars['Node2_Router_ID'])
                validating.values('Node2_Loopback', templateVars['Node2_Loopback'], ['no', 'yes'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "logical_node_profile.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        # Define the Template Source
        template_file = "logical_node_to_fabric_node.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Modify Variables for Template
        templateVars['Node_ID'] = templateVars['Node1_ID']
        templateVars['rtr_id'] = templateVars['Node1_Router_ID']
        templateVars['rtr_id_loopback'] = templateVars['Node1_Loopback']

        # Process the template through the Sites
        dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])
        

        if not templateVars['Node2_ID'] == None:
            # Modify Variables for Template
            templateVars['Node_ID'] = templateVars['Node2_ID']
            templateVars['rtr_id'] = templateVars['Node2_Router_ID']
            templateVars['rtr_id_loop_back'] = templateVars['Node2_Loopback']

            # Process the template through the Sites
            dest_file = 'L3Out_%s_Node_Profile_%s.tf' % (templateVars['L3Out'], templateVars['Node_Profile'])
            dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - Application Profiles
    #======================================================
    def ospf_interface(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.policies.ospfInterface']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'ospf_interface_policies'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Contracts - Add Subject
    #======================================================
    def subject_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationProfile']['allOf'][1]['properties']

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Dicts for required and optional args
        required_args = {'site_group': '',
                         'tenant': '',
                         'Subject': '',
                         'contract_type': '',
                         'contract': '',
                         'Reverse_Filter_Ports': '',
                         'qos_class': '',
                         'target_dscp': '',
                         'Filters_to_Assign': ''}
        optional_args = {'description': '',
                         'alias': '',
                         'annotation': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        templateVars['filters_count'] = 1
        try:
            # Validate Required Arguments
            validating.site_group('site_group', **kwargs)
            validating.dscp('target_dscp', templateVars['target_dscp'])
            validating.validator('contract', templateVars['contract'])
            validating.validator('Subject', templateVars['Subject'])
            validating.validator('tenant', templateVars['tenant'])
            validating.values('qos_class', jsonData, **kwargs)
            validating.values('contract_type', templateVars['contract_type'], ['OOB', 'Standard', 'Taboo'])
            validating.values('Reverse_Filter_Ports', templateVars['Reverse_Filter_Ports'], ['no', 'yes'])
            if not templateVars['alias'] == None:
                validating.validator('alias', templateVars['alias'])
            if not templateVars['description'] == None:
                validating.validator('description', templateVars['description'])
            if not templateVars['Filters_to_Assign'] == None:
                if re.search(',', templateVars['Filters_to_Assign']):
                    templateVars['filters_count'] =+ 1
                    for x in templateVars['Filters_to_Assign'].split(','):
                        validating.validator('Filters_to_Assign', x)
                else:
                    validating.validator('Filters_to_Assign', templateVars['Filters_to_Assign'])
        except Exception as err:
            errorReturn = '%s\nError on Worksheet %s Row %s.  Please verify Input Information.' % (
                SystemExit(err), kwargs['ws'], kwargs['row_num'])
            raise ErrException(errorReturn)

        # Define the Template Source
        template_file = "contract_subject.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Define the Template Source
        if templateVars['contract_type'] == 'OOB':
            template_file = "contract_subject.jinja2"
        elif templateVars['contract_type'] == 'Standard':
            template_file = "contract_subject.jinja2"
        elif templateVars['contract_type'] == 'Taboo':
            template_file = "contract_subject.jinja2"
        template = self.templateEnv.get_template(template_file)

        # Process the template through the Sites
        dest_file = 'contract_type_%s_%s_Subj_%s.tf' % (templateVars['contract_type'], templateVars['contract'], templateVars['Subject'])
        dest_dir = 'tenant_%s' % (templateVars['tenant'])

    #======================================================
    # Function - Bridge Domain - Subnets
    #======================================================
    def subnet_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.bd.Subnets']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

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
            'shared_between_vrfs',
            'neighbor_discovery',
            'no_default_svi_gateway',
            'querier_ip'
        ]
        for i in pop_list:
            templateVars.pop(i)

        # Attach the Subnet to the Bridge Domain
        if templateVars['site_group'] in kwargs['easyDict']['tenants']['bridge_domains'].keys():
            for i in kwargs['easyDict']['tenants']['bridge_domains'][templateVars['site_group']]:
                if i['name'] == templateVars['bridge_domain'] and i['tenant'] == templateVars['tenant']:
                    templateVars.pop('bridge_domain')
                    templateVars.pop('site_group')
                    templateVars.pop('tenant')
                    i['l3_configurations']['subnets'].append(templateVars)
                    break

        # Return easyDict
        return kwargs['easyDict']

    #======================================================
    # Function - Tenants
    #======================================================
    def tenant_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.Tenants']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['monitoring_policy'] = 'default'

        if re.search(r'\d{1,16}', kwargs['site_group']):
            if kwargs['easyDict']['tenants']['sites'].get(kwargs['site_group']):
                templateVars['sites'] = []
                templateVars['users'] = []
                for i in kwargs['easyDict']['tenants']['sites'][kwargs['site_group']]:
                    if i['tenant'] == templateVars['name']:
                        i.pop('class_type')
                        i.pop('site_group')
                        i.pop('tenant')
                        for x in i['users'].split(','):
                            if not x in templateVars['users']:
                                templateVars['users'].append(x)
                        i.pop('users')
                        templateVars['sites'].append(i)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'tenants'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - Tenants
    #======================================================
    def tenant_site(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.Sites']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'sites'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - VRFs
    #======================================================
    def vrf_add(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.Vrfs']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)
        templateVars['communities'] = []

        # Attach the VRF Policy Additional Attributes
        if kwargs['easyDict']['tenants']['vrf_policies'].get(templateVars['vrf_policy']):
            templateVars.update(kwargs['easyDict']['tenants']['vrf_policies'][templateVars['vrf_policy']])
        else:
            validating.error_policy_not_found('vrf_policy', **kwargs)

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'vrfs'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
        return kwargs['easyDict']
        
    #======================================================
    # Function - VRF - Communities
    #======================================================
    def vrf_community(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.vrfCommunity']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Check if the SNMP Community is in the Environment.  If not Add it.
        templateVars['jsonData'] = jsonData
        templateVars["Variable"] = f'vrf_snmp_community_{kwargs["community_variable"]}'
        sensitive_var_site_group(**templateVars)
        templateVars.pop('jsonData')
        templateVars.pop('Variable')

        # Add Dictionary to Policy
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'vrfs'
        templateVars['data_subtype'] = 'communities'
        templateVars['policy_name'] = templateVars['vrf']
        templateVars.pop('vrf')
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #======================================================
    # Function - VRF - Policy
    #======================================================
    def vrf_policy(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.vrfPolicy']['allOf'][1]['properties']

        # Validate User Input
        validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        templateVars.pop('policy_name')
        policy_dict = {kwargs['policy_name']:templateVars}

        # Add Dictionary to easyDict
        policy_dict['class_type'] = 'tenants'
        policy_dict['data_type'] = 'vrf_policies'
        kwargs['easyDict'] = easyDict_append_policy(policy_dict, **kwargs)
        return kwargs['easyDict']
