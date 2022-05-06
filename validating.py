#!/usr/bin/env python3

import ipaddress
import phonenumbers
import re
import validators

# Error Messages
def error_enforce(row_num, vrf):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num}. VRF {vrf}, Enforcement was not defined in the')
    print(f'   VRF Worksheet.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_enforcement(row_num, epg, ws2, ws3):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws3}. Enforcement on the EPG {epg}')
    print(f'   is set to enforced but the VRF is unenforced in {ws2}.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_policy_names(row_num, ws, policy_1, policy_2):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws.title}. The Policy {policy_1} was ')
    print(f'   not the same as {policy_2}. Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_int_selector(row_num, ws, int_select):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws.title}. Interface Selector {int_select}')
    print(f'   was not found in the terraform state file.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_login_domain(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[f'{var}_realm']
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws.title}. When {var} is set to {varValue}')
    print(f'   The Login Domain cannot be blank.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_request(status, text):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error in Retreiving Terraform Cloud Organization Workspaces')
    print(f'   Exiting on Error {status} with the following output:')
    print(f'   {text}')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_snmp_community(row_num, var):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num}; community_var {var}, was not pre-defined')
    print(f'   in the snmp_community section of the worksheet.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_snmp_user(row_num, var):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num}; snmp_user {var}, was not pre-defined')
    print(f'   in the snmp_user section of the worksheet.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_switch(row_num, ws, switch_ipr):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws.title}. Interface Profile {switch_ipr}')
    print(f'   was not found in the terraform state file.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_tenant(row_num, tenant, ws1, ws2):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws2}. Tenant {tenant} was not found')
    print(f'   in the {ws1} Worksheet.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_tenant_users(**templateVars):
    site_group = templateVars['site_group']
    tenant = templateVars['tenant']
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error with {site_group} tenant {tenant} users was empty.')
    print(f'   For Nexus Dashbord Orchestrator users is required.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_vlan_to_epg(row_num, vlan, ws):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num}. Did not Find EPG corresponding to VLAN {vlan}')
    print(f'   in Worksheet {ws.title}.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_vrf(row_num, vrf):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num}. VRF {vrf} was not found in the VRF Worksheet.')
    print(f'   Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

# Validations
def domain(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not validators.domain(varValue):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Domain {varValue}')
        print(f'   is invalid.  Please Validate the domain and retry.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def dns_name(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    hostname = varValue
    valid_count = 0
    if len(hostname) > 255:
        valid_count =+ 1
    if hostname[-1] == ".":
        hostname = hostname[:-1] # strip exactly one dot from the right, if present
    allowed = re.compile("(?!-)[A-Z\d-]{1,63}(?<!-)$", re.IGNORECASE)
    if not all(allowed.match(x) for x in hostname.split(".")):
        valid_count =+ 1
    if not valid_count == 0:
        print(f'\n--------------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue} ')
        print(f'   is not a valid Hostname.  Confirm that you have entered the DNS Name Correctly.')
        print(f'   Exiting....')
        print(f'\n--------------------------------------------------------------------------------\n')
        exit()

def email(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not validators.email(varValue, whitelist=None):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Email address "{varValue}"')
        print(f'   is invalid.  Please Validate the email and retry.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def filter_ports(var, kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    valid_count = 0
    if re.match(r'\d', varValue):
        if not validators.between(int(varValue), min=1, max=65535):
            valid_count =+ 1
    elif re.match(r'[a-z]', varValue):
        if not re.search('^(dns|ftpData|http|https|pop3|rtsp|smtp|unspecified)$', varValue):
            valid_count =+ 1
    else:
        valid_count =+ 1
    if not valid_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title} Row {row_num}. {var} {varValue} did not')
        print(f'   match allowed values. {var} can be:')
        print(f'    - dns')
        print(f'    - ftpData')
        print(f'    - http')
        print(f'    - https')
        print(f'    - pop3')
        print(f'    - rtsp')
        print(f'    - smtp')
        print(f'    - unspecified')
        print(f'    - or between 1 and 65535')
        print(f'   Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def hostname(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not (re.search('^[a-zA-Z0-9\\-]+$', varValue) and validators.length(varValue, min=1, max=63)):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue} ')
        print(f'   is not a valid Hostname.  Be sure you are not using the FQDN.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def ip_address(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if re.search('/', varValue):
        x = varValue.split('/')
        address = x[0]
    else:
        address = varValue
    valid_count = 0
    if re.search(r'\.', address):
        if not validators.ip_address.ipv4(address):
            valid_count =+ 1
    else:
        if not validators.ip_address.ipv6(address):
            valid_count =+ 1
    if not valid_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title} Row {row_num}. {var} {varValue} is not ')
        print(f'   a valid IPv4 or IPv6 Address.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def length_and_regex(pattern, varName, varValue, minimum, maximum):
    invalid_count = 0
    if not validators.length(varValue, min=int(minimum), max=int(maximum)):
        invalid_count += 1
        print(f'\n--------------------------------------------------------------------------------------\n')
        print(f'   !!! {varName} value "{varValue}" is Invalid!!!')
        print(f'   Length Must be between {minimum} and {maximum} characters.')
        print(f'\n--------------------------------------------------------------------------------------\n')
    if not re.search(pattern, varValue):
        invalid_count += 1
        print(f'\n--------------------------------------------------------------------------------------\n')
        print(f'   !!! Invalid Characters in {varValue}.  The allowed characters are:')
        print(f'   - "{pattern}"')
        print(f'\n--------------------------------------------------------------------------------------\n')
    if invalid_count == 0:
        return True
    else:
        return False

def length_and_regex_sensitive(pattern, varName, varValue, minimum, maximum):
    invalid_count = 0
    if not validators.length(varValue, min=int(minimum), max=int(maximum)):
        invalid_count += 1
        print(f'\n--------------------------------------------------------------------------------------\n')
        print(f'   !!! {varName} is Invalid!!!')
        print(f'   Length Must be between {minimum} and {maximum} characters.')
        print(f'\n--------------------------------------------------------------------------------------\n')
    if 'hashtag' in varName:
        if re.search(pattern, varValue):
            invalid_count += 1
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   The Shared Secret cannot contain backslash, space or hashtag.')
            print(f'\n-----------------------------------------------------------------------------\n')
    elif not re.search(pattern, varValue):
        invalid_count += 1
        print(f'\n--------------------------------------------------------------------------------------\n')
        print(f'   !!! Invalid Characters in {varName}.  The allowed characters are:')
        print(f'   - "{pattern}"')
        print(f'\n--------------------------------------------------------------------------------------\n')
    if invalid_count == 0:
        return True
    else:
        return False

def list_values(var, jsonData, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varList = jsonData[var]['enum']
    varValue = kwargs[var]
    match_count = 0
    for x in varList:
        if x == varValue:
            match_count =+ 1
    if not match_count > 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. ')
        print(f'   {var} should be one of the following:')
        for x in varList:
            print(f'    - {x}')
        print(f'    Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def mac_address(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not validators.mac_address.mac_address(varValue):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title} Row {row_num}. {var} {varValue} is not ')
        print(f'   a valid MAC Address.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def match_current_gw(row_num, current_inb_gwv4, inb_gwv4):
    if not current_inb_gwv4 == inb_gwv4:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Line {row_num}.  Current inband = "{current_inb_gwv4}" and found')
        print(f'   "{inb_gwv4}".  The Inband Network should be the same on all APICs and Switches.')
        print(f'   A Different Gateway was found.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def mgmt_network(row_num, ws, var1, var1_value, var2, var2_value):
    x = var1_value.split('/')
    ip_add = x[0]
    valid_count = 0
    if re.search(r'\.', ip_add):
        mgmt_check_ip = ipaddress.IPv4Interface(var1_value)
        mgmt_network = mgmt_check_ip.network
        if not ipaddress.IPv4Address(var2_value) in ipaddress.IPv4Network(mgmt_network):
            valid_count =+ 1
    else:
        if not validators.ip_address.ipv6(ip_add):
            valid_count =+ 1
    if not valid_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num}.  {var1} Network')
        print(f'   does not Match {var2} Network.')
        print(f'   Mgmt IP/Prefix: "{var1_value}"')
        print(f'   Gateway IP: "{var2_value}"')
        print(f'   Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def modules(row_num, name, switch_role, modules):
    module_count = 0
    if switch_role == 'leaf' and int(modules) == 1:
        module_count += 1
    elif switch_role == 'leaf':
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Row {row_num}. {name} module count is not valid.')
        print(f'   A Leaf can only have one module.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    elif switch_role == 'spine' and int(modules) < 17:
        module_count += 1
    elif switch_role == 'spine':
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Row {row_num}. {name} module count is not valid.')
        print(f'   A Spine needs between 1 and 16 modules.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def number_check(var, jsonData, **kwargs):
    minimum = jsonData[var]['minimum']
    maximum = jsonData[var]['maximum']
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not (validators.between(int(varValue), min=int(minimum), max=int(maximum))):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. Valid Values ')
        print(f'   are between {minimum} and {maximum}.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def number_list(var, jsonData, **kwargs):
    minimum = jsonData[var]['minimum']
    maximum = jsonData[var]['maximum']
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    for x in varValue.split(','):
        if not (validators.between(int(x), min=int(minimum), max=int(maximum))):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {x}. Valid Values ')
            print(f'   are between {minimum} and {maximum}.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

def not_empty(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if varValue == None:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. This is a  ')
        print(f'   required variable.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def phone_number(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    phone_number = phonenumbers.parse(varValue, None)
    if not phonenumbers.is_possible_number(phone_number):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Phone Number "{phone_number}" ')
        print(f'   is invalid.  Make sure you are including the country code and the full phone number.')
        print(f'   Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def string_list(var, jsonData, **kwargs):
    # Get Variables from Library
    minimum = jsonData[var]['minimum']
    maximum = jsonData[var]['maximum']
    pattern = jsonData[var]['pattern']
    row_num = kwargs['row_num']
    varValues = kwargs[var]
    ws = kwargs['ws']
    for varValue in varValues.split(','):
        if not (re.fullmatch(pattern,  varValue) and validators.length(
            str(varValue), min=int(minimum), max=int(maximum))):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. ')
            print(f'   "{varValue}" is an invalid Value...')
            print(f'   It failed one of the complexity tests:')
            print(f'    - Min Length {maximum}')
            print(f'    - Max Length {maximum}')
            print(f'    - Regex {pattern}')
            print(f'    Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

def string_pattern(var, jsonData, **kwargs):
    # Get Variables from Library
    minimum = jsonData[var]['minimum']
    maximum = jsonData[var]['maximum']
    pattern = jsonData[var]['pattern']
    row_num = kwargs['row_num']
    varValue = kwargs[var]
    ws = kwargs['ws']
    if not (re.fullmatch(pattern,  varValue) and validators.length(
        str(varValue), min=int(minimum), max=int(maximum))):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. ')
        print(f'   "{varValue}" is an invalid Value...')
        print(f'   It failed one of the complexity tests:')
        print(f'    - Min Length {maximum}')
        print(f'    - Max Length {maximum}')
        print(f'    - Regex {pattern}')
        print(f'    Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def site_group(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if 'Grp_' in varValue:
        if not re.search('Grp_[A-F]', varValue):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, Site_Group "{varValue}"')
            print(f'   is invalid.  A valid Group Name is Grp_A thru Grp_F.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()
    elif re.search(r'\d+', varValue):
        if not validators.between(int(varValue), min=1, max=15):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, Site_Group "{varValue}"')
            print(f'   is invalid.  A valid Site ID is between 1 and 15.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()
    else:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, Site_Group "{varValue}"')
        print(f'   is invalid.  A valid Site_Group is either 1 thru 15 or Group_A thru Group_F.')
        print(f'   Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def timeout(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    timeout_count = 0
    if not validators.between(int(varValue), min=5, max=60):
        timeout_count += 1
    if not (int(varValue) % 5 == 0):
        timeout_count += 1
    if not timeout_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. ')
        print(f'   {var} should be between 5 and 60 and be a factor of 5.  "{varValue}" ')
        print(f'   does not meet this.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def url(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    uRL = f'https://{varValue}'
    if not validators.url(uRL):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. ')
        print(f'   {var} should be a valid URL.  The Following is not a valid URL:')
        print(f'    - {uRL}')
        print(f'    Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def validator(var, **kwargs):
    # Get Variables from Library
    jsonData = kwargs['easy_jsonData']['components']['schemas']['globalData']['allOf'][1]['properties']
    minimum = jsonData[var]['minimum']
    maximum = jsonData[var]['maximum']
    pattern = jsonData[var]['pattern']
    row_num = kwargs['row_num']
    varValue = kwargs[var]
    ws = kwargs['ws']
    if not (re.fullmatch(pattern,  varValue) and validators.length(
        str(varValue), min=int(minimum), max=int(maximum))):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. ')
        print(f'   "{varValue}" is an invalid Value...')
        print(f'   It failed one of the complexity tests:')
        print(f'    - Min Length {maximum}')
        print(f'    - Max Length {maximum}')
        print(f'    - Regex {pattern}')
        print(f'    Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def validator_array(var, **kwargs):
    # Get Variables from Library
    jsonData = kwargs['easy_jsonData']['components']['schemas']['globalData']['allOf'][1]['properties']
    minimum = jsonData[var]['minimum']
    maximum = jsonData[var]['maximum']
    pattern = jsonData[var]['pattern']
    row_num = kwargs['row_num']
    varValue = kwargs[var]
    ws = kwargs['ws']
    for i in varValue:
        for k, v in i.items():
            if not (re.search(pattern,  k) and validators.length(
                str(k), min=int(minimum), max=int(maximum))):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. ')
                print(f'   "{k}" is an invalid Value...')
                print(f'   It failed one of the complexity tests:')
                print(f'    - Min Length {maximum}')
                print(f'    - Max Length {maximum}')
                print(f'    - Regex {pattern}')
                print(f'    Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()
            if not (re.search(pattern,  v) and validators.length(
                str(v), min=int(minimum), max=int(maximum))):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. ')
                print(f'   "{v}" is an invalid Value...')
                print(f'   It failed one of the complexity tests:')
                print(f'    - Min Length {maximum}')
                print(f'    - Max Length {maximum}')
                print(f'    - Regex {pattern}')
                print(f'    Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()

def validator_list(var, **kwargs):
    # Get Variables from Library
    jsonData = kwargs['easy_jsonData']['components']['schemas']['globalData']['allOf'][1]['properties']
    minimum = jsonData[var]['minimum']
    maximum = jsonData[var]['maximum']
    pattern = jsonData[var]['pattern']
    row_num = kwargs['row_num']
    varValue = kwargs[var]
    ws = kwargs['ws']
    for i in varValue:
        if not (re.search(pattern,  i) and validators.length(
            str(i), min=int(minimum), max=int(maximum))):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. ')
            print(f'   "{i}" is an invalid Value...')
            print(f'   It failed one of the complexity tests:')
            print(f'    - Min Length {maximum}')
            print(f'    - Max Length {maximum}')
            print(f'    - Regex {pattern}')
            print(f'    Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

def values(var, jsonData, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    if re.search('^(provider_)?version$', var) and ws.title == 'Sites':
        ctype = kwargs['controller_type']
        varList = jsonData[f'{var}_{ctype}']['enum']
    else:
        varList = jsonData[var]['enum']
    varValue = kwargs[var]
    match_count = 0
    for x in varList:
        if x == varValue:
            match_count =+ 1
    if not match_count > 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. ')
        print(f'   {var} should be one of the following:')
        for x in varList:
            print(f'    - {x}')
        print(f'    Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def vlans(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if re.search(',', str(varValue)):
        vlan_split = varValue.split(',')
        for x in vlan_split:
            if re.search('\\-', x):
                dash_split = x.split('-')
                for z in dash_split:
                    if not validators.between(int(z), min=1, max=4095):
                        print(f'\n-----------------------------------------------------------------------------\n')
                        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
                        print(f'   between 1 and 4095.  "{z}" is not valid.  Exiting....')
                        print(f'\n-----------------------------------------------------------------------------\n')
                        exit()
            elif not validators.between(int(x), min=1, max=4095):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
                print(f'   between 1 and 4095.  "{x}" is not valid.  Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()
    elif re.search('\\-', str(varValue)):
        dash_split = varValue.split('-')
        for x in dash_split:
            if not validators.between(int(x), min=1, max=4095):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
                print(f'   between 1 and 4095.  "{x}" is not valid.  Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()
    elif not validators.between(int(varValue), min=1, max=4095):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
        print(f'   between 1 and 4095.  "{varValue}" is not valid.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
