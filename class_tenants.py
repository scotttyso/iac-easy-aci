#!/usr/bin/env python3

#=============================================================================
# Source Modules
#=============================================================================
from collections import OrderedDict
from easy_functions import easyDict_append_policy
from easy_functions import easyDict_append, easyDict_append_subtype
from easy_functions import process_kwargs
from easy_functions import required_args_add, required_args_remove
from easy_functions import sensitive_var_site_group, validate_args
import json
import re
import validating

#=====================================================================================
# Please Refer to the "Notes" in the relevant column headers in the input Spreadhseet
# for detailed information on the Arguments used by this Function.
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

        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'apics_inband'
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
                'l3out_tenant':templateVars['vrf_tenant'],
                'route_profile':templateVars['l3_configurations']['route_profile']
            },
            'custom_mac_address':templateVars['custom_mac_address'],
        })
        aa = templateVars['l3_configurations']['associated_l3outs']
        if aa['l3out'] == None and aa['l3out_tenant'] == None and aa['route_profile'] == None:
            templateVars['l3_configurations'].pop('associated_l3outs')
        templateVars['l3_configurations'] = OrderedDict(sorted(templateVars['l3_configurations'].items()))

        pop_list = [
            'description',
            'endpoint_clear',
            'general_policy',
            'l3out',
            'l3_policy',
            'vrf',
            'vrf_tenant'
        ]
        for i in pop_list:
            templateVars.pop(i)
        
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
            'shared_between_vrfs',
            'neighbor_discovery',
            'no_default_svi_gateway',
            'querier_ip'
        ]
        for i in pop_list:
            templateVars.pop(i)
        
        # Add Dictionary to easyDict
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'bridge_domain_subnets'
        kwargs['easyDict'] = easyDict_append(templateVars, **kwargs)
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
        templateVars['data_type'] = 'bgp_peer_prefix_policies'
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
        templateVars['peer_controls'] = {
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
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.applicationProfile']['allOf'][1]['properties']

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
    # Function - Contracts - Add Subject
    #=============================================================================
    def contract_filters(self, **kwargs):
        # Get Variables from Library
        jsonData = kwargs['easy_jsonData']['components']['schemas']['tenants.contract.ContractFilters']['allOf'][1]['properties']

        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        # Add Dictionary to Policy
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'contracts'
        templateVars['data_subtype'] = 'filters'
        templateVars['policy_name'] = templateVars['contract']
        templateVars.pop('contract')
        kwargs['easyDict'] = easyDict_append_subtype(templateVars, **kwargs)
        return kwargs['easyDict']

    #=============================================================================
    # Function - Application Profiles
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
        templateVars['data_type'] = 'eigrp_interface_policies'
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
        
        # Validate User Input
        kwargs = validate_args(jsonData, **kwargs)

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(jsonData['required_args'], jsonData['optional_args'], **kwargs)

        if re.search('^(inb|oob)$', epgpolicy['epg_type']):
            jsonData = required_args_add(pop_list, jsonData)

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
                print(templateVars['domains'])
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
            templateVars.pop(i)

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
                        templateVars['jsonData'] = jsonData
                        templateVars["Variable"] = f'bgp_password_{aa["bgp_password"]}'
                        sensitive_var_site_group(**templateVars)
                        templateVars.pop('jsonData')
                        templateVars.pop('Variable')

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
                    templateVars['jsonData'] = jsonData
                    templateVars["Variable"] = f'ospf_key_{aa["key_id"]}'
                    sensitive_var_site_group(**templateVars)
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
        templateVars['class_type'] = 'tenants'
        templateVars['data_type'] = 'l3out_logical_node_profiles'
        templateVars['data_subtype'] = 'interface_profiles'
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
        # split_list = [
        #     'primary_preferred_addresses',
        #     'link_locals',
        #     'mac_addresses',
        #     'secondary_addresses',
        # ]
        # for i in split_list:
        #     if not templateVars[i] == None:
        #         templateVars[i] = templateVars[i].split(',')

        templateVars.pop('policy_name')
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
        templateVars['data_type'] = 'ospf_interface_policies'
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
