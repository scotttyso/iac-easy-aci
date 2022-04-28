#!/usr/bin/env python3

import ipaddress
import phonenumbers
import re
import validators

# Validations
def bool(row_num, ws, var, var_value):
    if not (var_value == 'True' or var_value == 'False'):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {var_value}. ')
        print(f'   The Option should be True or False but recieved {var_value}.')
        print(f'    Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def brkout_pg(row_num, brkout_pg):
    if not re.search('(2x100g_pg|4x100g_pg|4x10g_pg|4x25g_pg|8x50g_pg)', brkout_pg):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Row {row_num}. Breakout Port Group is Invalid.  Valid Values are:')
        print(f'   2x100g_pg, 4x100g_pg, 4x10g_pg, 4x25g_pg, and 8x50g_pg.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def description(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not (re.search(r'^[a-zA-Z0-9\\!#$%()*,-./:;@ _{|}~?&+]+$',  varValue) and validators.length(str(varValue), min=0, max=128)):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. ')
        print(f'   The description is an invalid Value... It failed one of the complexity tests:')
        print(f'    - Min Length 0')
        print(f'    - Max Length 128')
        print('    - Regex [a-zA-Z0-9\\!#$%()*,-./:;@ _{|}~?&+]+')
        print(f'    Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

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

def dscp(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not re.search('^(AF[1-4][1-3]|CS[0-7]|EF|VA|unspecified)$', varValue):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid Values are:')
        print(f'   AF11, AF12, AF13, AF21, AF22, AF23, AF31, AF32, AF33, AF41, AF42, AF43,')
        print(f'   CS0, CS1, CS2, CS3, CS4, CS5, CS6, CS7, EF, VA or unspecified.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
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
    print(f'   Error on Row {row_num} of Worksheet {ws}. The Policy {policy_1} was ')
    print(f'   not the same as {policy_2}. Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_int_selector(row_num, ws, int_select):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws}. Interface Selector {int_select}')
    print(f'   was not found in the terraform state file.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_request(status, text):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error in Retreiving Terraform Cloud Organization Workspaces')
    print(f'   Exiting on Error {status} with the following output:')
    print(f'   {text}')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_switch(row_num, ws, switch_ipr):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num} of Worksheet {ws}. Interface Profile {switch_ipr}')
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
    print(f'   in Worksheet {ws}.  Exiting....')
    print(f'\n-----------------------------------------------------------------------------\n')
    exit()

def error_vrf(row_num, vrf):
    print(f'\n-----------------------------------------------------------------------------\n')
    print(f'   Error on Row {row_num}. VRF {vrf} was not found in the VRF Worksheet.')
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
    if not re.search(pattern, varValue):
        invalid_count += 1
        print(f'\n--------------------------------------------------------------------------------------\n')
        print(f'   !!! Invalid Characters in {varName}.  The allowed characters are:')
        print(f'   - "{pattern}"')
        print(f'\n--------------------------------------------------------------------------------------\n')
    if invalid_count == 0:
        return True
    else:
        return False

def name_complexity(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    login_domain_count = 0
    if not re.fullmatch('^([a-zA-Z0-9\\_]+)$', varValue):
        login_domain_count += 1
    elif not validators.length(varValue, min=1, max=10):
        login_domain_count += 1
    if not login_domain_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num}, {var}, {varValue}.  The Value')
        print(f'   must be between 1 and 10 characters.  The only non alphanumeric characters')
        print(f'   allowed is "_"; but it must not start with "_".  "{varValue}" did not')
        print(f'   meet these restrictions.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def name_lists(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    for i in varValue:
        if not (re.search(r'^[a-zA-Z0-9_-]+$',  i) and validators.length(str(i), min=0, max=63)):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {i}. ')
            print(f'   {i} is an invalid Value... It failed one of the complexity tests:')
            print(f'    - Min Length 0')
            print(f'    - Max Length 63')
            print(f'    - Regex [a-zA-Z0-9_-]+')
            print(f'    Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

def name_maps(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    for i in varValue:
        for k, v in i.items():
            if not (re.search(r'^[a-zA-Z0-9_-]+$',  k) and validators.length(str(k), min=0, max=63)):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {k}. ')
                print(f'   {k} is an invalid Value... It failed one of the complexity tests:')
                print(f'    - Min Length 0')
                print(f'    - Max Length 63')
                print(f'    - Regex [a-zA-Z0-9_-]+')
                print(f'    Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()
            if not (re.search(r'^[a-zA-Z0-9_-]+$',  v) and validators.length(str(v), min=0, max=63)):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {v}. ')
                print(f'   {v} is an invalid Value... It failed one of the complexity tests:')
                print(f'    - Min Length 0')
                print(f'    - Max Length 63')
                print(f'    - Regex [a-zA-Z0-9_-]+')
                print(f'    Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()

def name_rule(var, **kwargs):
    row_num = kwargs['row_num']
    ws = kwargs['ws']
    varValue = kwargs[var]
    if not (re.search(r'^[a-zA-Z0-9_-]+$',  varValue) and validators.length(str(varValue), min=0, max=63)):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {varValue}. ')
        print(f'   {var} is an invalid Value... It failed one of the complexity tests:')
        print(f'    - Min Length 0')
        print(f'    - Max Length 63')
        print(f'    - Regex [a-zA-Z0-9_-]+')
        print(f'    Exiting....')
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


## Should these remain

def encryption_key(row_num, ws, var, var_value):
    if not validators.length(str(var_value), min=16, max=32):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. The Encryption Key')
        print(f'   Length must be between 16 and 32 characters.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def filter_ports(row_num, ws, var, var_value):
    valid_count = 0
    if re.match(r'\d', var_value):
        if not validators.between(int(var_value), min=1, max=65535):
            valid_count =+ 1
    elif re.match(r'[a-z]', var_value):
        if not re.search('^(dns|ftpData|http|https|pop3|rtsp|smtp|unspecified)$', var_value):
            valid_count =+ 1
    else:
        valid_count =+ 1
    if not valid_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title} Row {row_num}. {var} {var_value} did not')
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

def link_level(row_num, ws, var, var_value):
    if not re.search('(_Auto|_NoNeg)$', var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, value {var_value}:')
        print(f'   Please Select a valid Link Level from the drop-down menu.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    if not re.search('^(100M_|1G_|(1|4|5)0G_|25G_|[1-2]00G_|400G_|inherit_)', var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, value {var_value}:')
        print(f'   Please Select a valid Link Level from the drop-down menu.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def log_level(row_num, ws, var, var_value):
    if var == 'Severity' or var == 'Local_Level' or var == 'Minimum_Level':
        if not re.match('(emergencies|alerts|critical|errors|warnings|notifications|information|debugging)', var_value):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title} Row {row_num}. Logging Level for "{var}"')
            print(f'   with "{var_value}" is not valid.  Logging Levels can be:')
            print(f'   [emergencies|alerts|critical|errors|warnings|notifications|information|debugging]')
            print(f'   Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()
    elif var == 'Console_Level':
        if not re.match('^(emergencies|alerts|critical)$', var_value):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title} Row {row_num}. Logging Level for "{var}"  with "{var_value}"')
            print(f'   is not valid.  Logging Levels can be: [emergencies|alerts|critical].  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

def login_type(row_num, ws, var1, var1_value, var2, var2_value):
    login_type_count = 0
    if var1_value == 'console':
        if not re.fullmatch('^(local|ldap|radius|tacacs|rsa)$', var2_value):
            login_type_count += 1
    elif var1_value == 'default':
        if not re.fullmatch('^(local|ldap|radius|tacacs|rsa|saml)$', var2_value):
            login_type_count += 1
    if not login_type_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error in Worksheet {ws.title} Row {row_num}.  The Login Domain Type should be')
        print(f'   one of the following:')
        if var1_value == 'console':
            print(f'       [local|ldap|radius|tacacs|rsa]')
        elif var1_value == 'default':
            print(f'       [local|ldap|radius|tacacs|rsa|saml]')
        print(f'   "{var2_value}" did not match one of these types.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def mac_address(row_num, ws, var, var_value):
    if not validators.mac_address.mac_address(var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title} Row {row_num}. {var} {var_value} is not ')
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

def mgmt_domain(row_num, ws, var, var_value):
    if var_value == 'oob':
        var_value = 'oob-default'
    elif var_value == 'inband':
        var_value = 'inb-default'
    else:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var} value {var_value}.')
        print(f'   The Management Domain Should be inband or oob.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    return var_value

def mgmt_epg(row_num, ws, var, var_value):
    if var_value == 'var_inb':
        var_value = 'in_band'
    elif var_value == 'var_oob':
        var_value = 'out_of_band'
    else:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var} value {var_value}.')
        print(f'   The Management EPG Should be var_inb or var_oob.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    return var_value

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

def policy_type(row_num, ws, var, policy_group):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {policy_group}. A required')
        print(f'   policy of type {var} was not found.  Please verify {policy_group} is.')
        print(f'   configured properly.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def port_count(row_num, name, switch_role, port_count):
    if not re.search('^(16|32|34|36|48|54|60|64|66|102|108)$', port_count):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Row {row_num}. {name} port count of {port_count} is not valid.')
        print(f'   Valid port counts are 16, 32, 34, 36, 48, 54, 60, 64, 66, 102, 108.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def sensitive_var(row_num, ws, var, var_value):
    if not re.search('^(sensitive_var[1-7])$', var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid Values are:')
        print(f'   sensitive_var[1-7].  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def secret(row_num, ws, var, var_value):
    if not validators.length(var_value, min=1, max=32):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num}, {var}, {var_value}')
        print(f'   The Shared Secret Length must be between 1 and 32 characters.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    if re.search('[\\\\ #]+', var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num}, {var}, {var_value}')
        print(f'   The Shared Secret cannot contain backslash, space or hashtag.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def snmp_auth(row_num, ws, priv_type, priv_key, auth_type, auth_key):
    if not (priv_type == None or priv_type == 'none' or priv_type == 'aes-128' or priv_type == 'des'):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'    Error on Worksheet {ws.title}, Row {row_num}. priv_type {priv_type} is not ')
        print(f'    [none|des|aes-128].  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    if not (priv_type == 'none' or priv_type == None):
        if not validators.length(priv_key, min=8, max=32):
            print(f'\n-----------------------------------------------------------------------------\n')
            print(f'   Error on Worksheet {ws.title}, Row {row_num}. priv_key does not ')
            print(f'   meet the minimum character count of 8 or the maximum of 32.  Exiting....')
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()
    if not (auth_type == 'md5' or auth_type == 'sha1'):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'    Error on Worksheet {ws.title}, Row {row_num}. priv_type {priv_type} is not ')
        print(f'    [md5|sha1].  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    if not validators.length(auth_key, min=8, max=32):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num}. auth_key does not ')
        print(f'   meet the minimum character count of 8 or the maximum of 32.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def snmp_mgmt(row_num, ws, var, var_value):
    if var_value == 'oob':
        var_value = 'Out-of-Band'
    elif var_value == 'inband':
        var_value = 'Inband'
    else:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var} value {var_value}.')
        print(f'   The Management Domain Should be inband or oob.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    return var_value

def snmp_string(row_num, ws, var, var_value):
    if not (validators.length(var_value, min=8, max=32) and re.fullmatch('^([a-zA-Z0-9\\-\\_\\.]+)$', var_value)):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, "{var_value}" is invalid.')
        print(f'   The community/username policy name must be a minimum of 8 and maximum of 32 ')
        print(f'   characters in length.  The name can contain only letters, numbers and the ')
        print(f'   special characters of underscore (_), hyphen (-), or period (.). The name ')
        print(f'   cannot contain the @ symbol.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def stp(row_num, ws, var, var_value):
    if not re.search('^(BPDU_)', var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, value {var_value}:')
        print(f'   Please Select a valid STP Policy from the drop-down menu.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
    if not re.search('(ft_and_gd|ft_or_gd|_ft|_gd)$', var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, value {var_value}:')
        print(f'   Please Select a valid STP Policy from the drop-down menu.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def sw_version(row_num, ws, var1, var_value):
    ver_count = 0
    if re.match('^n9000', var_value):
        regex = re.compile(r'^n9000\-\d{2}\.\d{1,2}\(\d{1,2}[a-z]\)$')
        if not re.fullmatch(regex, var_value):
            ver_count += 1
    else:
        regex = re.compile(r'^simsw-\d{1}\.\d{1,2}\(\d{1,2}[a-z]\)$')
        if not re.fullmatch(regex, var_value):
            ver_count += 1
    if not ver_count == 0:
        print(f"\n-----------------------------------------------------------------------------\n")
        print(f"   Error in Worksheet {ws.title} Row {row_num}.  The SW_Version {var_value}")
        print(f"   did not match against the required regex of:")
        print(f"    - {regex}.")
        print(f"   Exiting....")
        print(f"\n-----------------------------------------------------------------------------\n")
        exit()

def syslog_fac(row_num, ws, var, var_value):
    if not re.match("^local[0-7]$", var_value):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, "{var_value}" is invalid.')
        print(f'   Please verify Syslog Facility {var_value}.  Exiting...\n')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def tag_check(row_num, ws, var, var_value):
    tag_list = ['alice-blue', 'antique-white', 'aqua', 'aquamarine', 'azure', 'beige', 'bisque', 'black', 'blanched-almond', 'blue', 'blue-violet',
    'brown', 'burlywood', 'cadet-blue', 'chartreuse', 'chocolate', 'coral', 'cornflower-blue', 'cornsilk', 'crimson', 'cyan', 'dark-blue', 'dark-cyan',
    'dark-goldenrod', 'dark-gray', 'dark-green', 'dark-khaki', 'dark-magenta', 'dark-olive-green', 'dark-orange', 'dark-orchid', 'dark-red', 'dark-salmon',
    'dark-sea-green', 'dark-slate-blue', 'dark-slate-gray', 'dark-turquoise', 'dark-violet', 'deep-pink', 'deep-sky-blue', 'dim-gray', 'dodger-blue',
    'fire-brick', 'floral-white', 'forest-green', 'fuchsia', 'gainsboro', 'ghost-white', 'gold', 'goldenrod', 'gray', 'green', 'green-yellow', 'honeydew',
    'hot-pink', 'indian-red', 'indigo', 'ivory', 'khaki', 'lavender', 'lavender-blush', 'lawn-green', 'lemon-chiffon', 'light-blue', 'light-coral',
    'light-cyan', 'light-goldenrod-yellow', 'light-gray', 'light-green', 'light-pink', 'light-salmon', 'light-sea-green', 'light-sky-blue',
    'light-slate-gray', 'light-steel-blue', 'light-yellow', 'lime', 'lime-green', 'linen', 'magenta', 'maroon', 'medium-aquamarine', 'medium-blue',
    'medium-orchid', 'medium-purple', 'medium-sea-green', 'medium-slate-blue', 'medium-spring-green', 'medium-turquoise', 'medium-violet-red', 'midnight-blue',
    'mint-cream', 'misty-rose', 'moccasin', 'navajo-white', 'navy', 'old-lace', 'olive', 'olive-drab', 'orange', 'orange-red', 'orchid', 'pale-goldenrod',
    'pale-green', 'pale-turquoise', 'pale-violet-red', 'papaya-whip', 'peachpuff', 'peru', 'pink', 'plum', 'powder-blue', 'purple', 'red', 'rosy-brown',
    'royal-blue', 'saddle-brown', 'salmon', 'sandy-brown', 'sea-green', 'seashell', 'sienna', 'silver', 'sky-blue', 'slate-blue', 'slate-gray', 'snow',
    'spring-green', 'steel-blue', 'tan', 'teal', 'thistle', 'tomato', 'turquoise', 'violet', 'wheat', 'white', 'white-smoke', 'yellow', 'yellow-green' ]
    regx = re.compile(var_value)
    if not list(filter(regx.match, tag_list)):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, "{var_value}" is invalid.')
        print(f'   Valid Tag Values are:')
        print(f'   alice-blue, antique-white, aqua, aquamarine, azure, beige, bisque, black,')
        print(f'   blanched-almond, blue, blue-violet, brown, burlywood, cadet-blue, chartreuse,')
        print(f'   chocolate, coral, cornflower-blue, cornsilk, crimson, cyan, dark-blue, dark-cyan,')
        print(f'   dark-goldenrod, dark-gray, dark-green, dark-khaki, dark-magenta, dark-olive-green,')
        print(f'   dark-orange, dark-orchid, dark-red, dark-salmon, dark-sea-green, dark-slate-blue,')
        print(f'   dark-slate-gray, dark-turquoise, dark-violet, deep-pink, deep-sky-blue, dim-gray,')
        print(f'   dodger-blue, fire-brick, floral-white, forest-green, fuchsia, gainsboro, ghost-white,')
        print(f'   gold, goldenrod, gray, green, green-yellow, honeydew, hot-pink, indian-red, indigo,')
        print(f'   ivory, khaki, lavender, lavender-blush, lawn-green, lemon-chiffon, light-blue,')
        print(f'   light-coral, light-cyan, light-goldenrod-yellow, light-gray, light-green, light-pink,')
        print(f'   light-salmon, light-sea-green, light-sky-blue, light-slate-gray, light-steel-blue,')
        print(f'   light-yellow, lime, lime-green, linen, magenta, maroon, medium-aquamarine, medium-blue,')
        print(f'   medium-orchid, medium-purple, medium-sea-green, medium-slate-blue, medium-spring-green,')
        print(f'   medium-turquoise, medium-violet-red, midnight-blue, mint-cream, misty-rose, moccasin,')
        print(f'   navajo-white, navy, old-lace, olive, olive-drab, orange, orange-red, orchid,')
        print(f'   pale-goldenrod, pale-green, pale-turquoise, pale-violet-red, papaya-whip, peachpuff,')
        print(f'   peru, pink, plum, powder-blue, purple, red, rosy-brown, royal-blue, saddle-brown,')
        print(f'   salmon, sandy-brown, sea-green, seashell, sienna, silver, sky-blue, slate-blue,')
        print(f'   slate-gray, snow, spring-green, steel-blue, tan, teal, thistle, tomato, turquoise,')
        print(f'   violet, wheat, white, white-smoke, yellow, and yellow-green')
        print(f'   Exiting...\n')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()


def timeout(row_num, ws, var, var_value):
    timeout_count = 0
    if not validators.between(int(var_value), min=5, max=60):
        timeout_count += 1
    if not (int(var_value) % 5 == 0):
        timeout_count += 1
    if not timeout_count == 0:
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}, {var_value}. ')
        print(f'   {var} should be between 5 and 60 and be a factor of 5.  "{var_value}" ')
        print(f'   does not meet this.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()

def vlans(row_num, ws, var, var_value):
    if re.search(',', str(var_value)):
        vlan_split = var_value.split(',')
        for x in vlan_split:
            if re.search('\\-', x):
                dash_split = x.split('-')
                for z in dash_split:
                    if not validators.between(int(z), min=1, max=4095):
                        print(f'\n-----------------------------------------------------------------------------\n')
                        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
                        print(f'   between 1 and 4095.  Exiting....')
                        print(f'\n-----------------------------------------------------------------------------\n')
                        exit()
            elif not validators.between(int(x), min=1, max=4095):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
                print(f'   between 1 and 4095.  Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()
    elif re.search('\\-', str(var_value)):
        dash_split = var_value.split('-')
        for x in dash_split:
            if not validators.between(int(x), min=1, max=4095):
                print(f'\n-----------------------------------------------------------------------------\n')
                print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
                print(f'   between 1 and 4095.  Exiting....')
                print(f'\n-----------------------------------------------------------------------------\n')
                exit()
    elif not validators.between(int(var_value), min=1, max=4095):
        print(f'\n-----------------------------------------------------------------------------\n')
        print(f'   Error on Worksheet {ws.title}, Row {row_num} {var}. Valid VLAN Values are:')
        print(f'   between 1 and 4095.  Exiting....')
        print(f'\n-----------------------------------------------------------------------------\n')
        exit()
