import collections
import jinja2
import json
import os
import platform
import pkg_resources
import re
import requests
import stdiomask
import sys
import time
import validating
import urllib3
from easy_functions import policies_parse
from easy_functions import sensitive_var_value
from easy_functions import varBoolLoop
from easy_functions import variablesFromAPI
from easy_functions import varStringLoop
from io import StringIO
from lxml import etree
from requests.api import delete, request
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Global options for debugging
print_payload = False
print_response_always = False
print_response_on_fail = True

# Log levels 0 = None, 1 = Class only, 2 = Line
log_level = 2

# Global path to main Template directory
tf_template_path = pkg_resources.resource_filename('class_terraform', 'templates/')

# Exception Classes
class InsufficientArgs(Exception):
    pass

class ErrException(Exception):
    pass

class InvalidArg(Exception):
    pass

class LoginFailed(Exception):
    pass

class FabLogin(object):
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
                             object_pairs_hook=collections.OrderedDict)
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

# Terraform Cloud For Business - Policies
# Class must be instantiated with Variables
class terraform_cloud(object):
    def __init__(self):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(tf_template_path + 'terraform/'))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)

    def create_terraform_workspaces(easy_jsonData, folders, site):
        opSystem = platform.system()
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
            templateVars = {}
            templateVars["terraform_cloud_token"] = terraform_cloud().terraform_token()
            
            # Obtain Terraform Cloud Organization
            if os.environ.get('tfc_organization') is None:
                templateVars["tfc_organization"] = terraform_cloud().tfc_organization(**templateVars)
                os.environ['tfc_organization'] = templateVars["tfc_organization"]
            else:
                templateVars["tfc_organization"] = os.environ.get('tfc_organization')
            tfcb_config.append({'tfc_organization':templateVars["tfc_organization"]})
            
            # Obtain Version Control Provider
            if os.environ.get('tfc_vcs_provider') is None:
                tfc_vcs_provider,templateVars["tfc_oath_token"] = terraform_cloud().tfc_vcs_providers(**templateVars)
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
            
            templateVars["agentPoolId"] = ''
            templateVars["allowDestroyPlan"] = False
            templateVars["executionMode"] = 'remote'
            templateVars["queueAllRuns"] = False
            templateVars["speculativeEnabled"] = True
            templateVars["triggerPrefixes"] = []

            # Query the Terraform Versions from the Release URL
            terraform_versions = []
            url = f'https://releases.hashicorp.com/terraform/'
            r = requests.get(url)
            html = r.content.decode("utf-8")
            parser = etree.HTMLParser()
            tree = etree.parse(StringIO(html), parser=parser)
            # This will get the anchor tags <a href...>
            refs = tree.xpath("//a")
            links = [link.get('href', '') for link in refs]
            for i in links:
                if re.search(r'/terraform/[1-2]\.[0-9]+\.[0-9]+/', i):
                    tf_version = re.search(r'/terraform/([1-2]\.[0-9]+\.[0-9]+)/', i).group(1)
                    terraform_versions.append(tf_version)

            # Removing Deprecated Versions from the List
            deprecatedVersions = ["1.1.0", "1.1.1"]
            for depver in deprecatedVersions:
                verCount = 0
                for Version in terraform_versions:
                    if str(depver) == str(Version):
                        terraform_versions.pop(verCount)
                    verCount += 1
            
            # Assign the Terraform Version from the Terraform Release URL Above
            templateVars["multi_select"] = False
            templateVars["var_description"] = "Terraform Version for Workspaces:"
            templateVars["jsonVars"] = terraform_versions
            templateVars["varType"] = 'Terraform Version'
            templateVars["defaultVar"] = ''

            # Obtain Terraform Workspace Version
            if os.environ.get('terraformVersion') is None:
                templateVars["terraformVersion"] = variablesFromAPI(**templateVars)
                os.environ['terraformVersion'] = templateVars["terraformVersion"]
            else:
                templateVars["terraformVersion"] = os.environ.get('terraformVersion')

            repoFoldercheck = False
            while repoFoldercheck == False:
                if not os.environ.get('tfWorkDir') is None:
                    tfDir = os.environ.get('tfWorkDir')
                else:
                    if os.environ.get('TF_DEST_DIR') is None:
                        tfDir = 'Intersight'
                        os.environ['tfWorkDir'] = 'Intersight'
                    else:
                        tfDir = os.environ.get('TF_DEST_DIR')
                if (opSystem == 'Windows' and re.search(r'(^\\|^\.\\)', tfDir)) or re.search(r'(^\/|^\.\.)', tfDir):
                    print(f'\n-------------------------------------------------------------------------------------------\n')
                    print(f'  Within Terraform Cloud, the Workspace will be configured with the directory where the ')
                    print(f'  configuration files are stored in the repo: {templateVars["vcsBaseRepo"]}.')
                    print(f'  For Example if the shortpath was "Intersight", The Repo URL would end up like:\n')
                    for folder in folders:
                        if opSystem == 'Windows':
                            print(f'    - {templateVars["vcsBaseRepo"]}\\{site}\\{folder}')
                        else:
                            print(f'    - {templateVars["vcsBaseRepo"]}/{site}/{folder}')
                    print(f'  The Destination Directory has been entered as:\n')
                    print(f'  {tfDir}\n')
                    print(f'  Which looks to be a system path instead of a Repository Directory.')
                    print(f'  Please confirm the Path Below is the short Path to the Repository Directory.')
                    print(f'\n-------------------------------------------------------------------------------------------\n')
                if opSystem == 'Windows' and re.search(r'(^\\|^\.\\)', tfDir):
                    question = input(f'Enter Value to Make Corrections: [Press Enter to Leave Base Path Empty]: ')
                    if question == '':
                        tfDir = ''
                        os.environ['tfWorkDir'] = tfDir
                        repoFoldercheck = True
                    else:
                        tfDir = question
                        os.environ['tfWorkDir'] = tfDir
                        repoFoldercheck = True
                elif re.search(r'(^\/|^\.\.)', tfDir):
                    dirLength = len(tfDir.split('/'))
                    question = input(f'Enter Value to Make Corrections: [Press Enter to Leave Base Path Empty]: ')
                    if question == '':
                        tfDir = ''
                        os.environ['tfWorkDir'] = tfDir
                        repoFoldercheck = True
                    else:
                        tfDir = question
                        os.environ['tfWorkDir'] = tfDir
                        repoFoldercheck = True
                else:
                    repoFoldercheck = True

            if opSystem == 'Windows':
                folder_list = [
                    f'{tfDir}\\{site}\\policies',
                    f'{tfDir}\\{site}\\pools',
                    f'{tfDir}\\{site}\\profiles',
                    f'{tfDir}\\{site}\\ucs_domain_profiles'
                ]
            else:
                folder_list = [
                    f'{tfDir}/{site}/policies',
                    f'{tfDir}/{site}/pools',
                    f'{tfDir}/{site}/profiles',
                    f'{tfDir}/{site}/ucs_domain_profiles'
                ]

            for folder in folder_list:
                if opSystem == 'Windows':
                    folder_length = len(folder.split('\\'))
                else:
                    folder_length = len(folder.split('/'))

                templateVars["autoApply"] = True
                if opSystem == 'Windows':
                    templateVars["Description"] = f'Site {site} - %s' % (folder.split('\\')[folder_length -2])
                else:
                    templateVars["Description"] = f'Site {site} - %s' % (folder.split('/')[folder_length -2])
                if opSystem == 'Windows':
                    fSplit = folder.split('\\')[folder_length -1]
                else:
                    fSplit = folder.split('/')[folder_length -1]
                if re.search('(pools|policies|ucs_domain_profiles)', fSplit):
                    templateVars["globalRemoteState"] = True
                else:
                    templateVars["globalRemoteState"] = False
                templateVars["workingDirectory"] = folder

                if opSystem == 'Windows':
                    fSplit = folder.split("\\")[folder_length -1]
                else:
                    fSplit = folder.split("/")[folder_length -1]
                templateVars["Description"] = f'Name of the {fSplit} Workspace to Create in Terraform Cloud'
                templateVars["varDefault"] = f'{site}_{fSplit}'
                templateVars["varInput"] = f'Terraform Cloud Workspace Name. [{site}_{fSplit}]: '
                templateVars["varName"] = f'Workspace Name'
                templateVars["varRegex"] = '^[a-zA-Z0-9\\-\\_]+$'
                templateVars["minLength"] = 1
                templateVars["maxLength"] = 90
                templateVars["workspaceName"] = varStringLoop(**templateVars)
                if opSystem == 'Windows':
                    tfcb_config.append({folder.split('\\')[folder_length -1]:templateVars["workspaceName"]})
                else:
                    tfcb_config.append({folder.split('/')[folder_length -1]:templateVars["workspaceName"]})
                # templateVars["vcsBranch"] = ''

                templateVars['workspace_id'] = terraform_cloud().tfcWorkspace(**templateVars)
                vars = [
                    'apikey.Intersight API Key',
                    'secretkey.Intersight Secret Key'
                ]
                for var in vars:
                    print(f'* Adding {var.split(".")[1]} to {templateVars["workspaceName"]}')
                    templateVars["Variable"] = var.split('.')[0]
                    if 'secret' in var:
                        templateVars["Multi_Line_Input"] = True
                    templateVars["Description"] = var.split('.')[1]
                    templateVars["varId"] = var.split('.')[0]
                    templateVars["varKey"] = var.split('.')[0]
                    templateVars["varValue"] = sensitive_var_value(easy_jsonData, **templateVars)
                    templateVars["Sensitive"] = True
                    if 'secret' in var and opSystem == 'Windows':
                        if os.path.isfile(templateVars["varValue"]):
                            f = open(templateVars["varValue"])
                            templateVars["varValue"] = f.read().replace('\n', '\\n')
                    terraform_cloud().tfcVariables(**templateVars)

                if opSystem == 'Windows':
                    folderSplit = folder.split("\\")[folder_length -1]
                else:
                    folderSplit = folder.split("/")[folder_length -1]
                if folderSplit == 'policies':
                    templateVars["Multi_Line_Input"] = False
                    vars = [
                        'ipmi_over_lan_policies.ipmi_key',
                        'iscsi_boot_policies.password',
                        'ldap_policies.binding_password',
                        'local_user_policies.local_user_password',
                        'persistent_memory_policies.secure_passphrase',
                        'snmp_policies.access_community_string',
                        'snmp_policies.password',
                        'snmp_policies.trap_community_string',
                        'virtual_media_policies.vmedia_password'
                    ]
                    sensitive_vars = []
                    for var in vars:
                        policy_type = 'policies'
                        policy = '%s' % (var.split('.')[0])
                        policies,json_data = policies_parse(site, policy_type, policy)
                        y = var.split('.')[0]
                        z = var.split('.')[1]
                        if y == 'persistent_memory_policies':
                            if len(policies) > 0:
                                sensitive_vars.append(z)
                        else:
                            for keys, values in json_data.items():
                                for key, value in values.items():
                                    for k, v in value.items():
                                        if k == z:
                                            if not v == 0:
                                                if y == 'iscsi_boot_policies':
                                                    varValue = 'iscsi_boot_password'
                                                else:
                                                    varValue = '%s_%s' % (k, v)
                                                sensitive_vars.append(varValue)
                                        elif k == 'binding_parameters':
                                            for ka, va in v.items():
                                                if ka == 'bind_method':
                                                    if va == 'ConfiguredCredentials':
                                                        sensitive_vars.append('binding_parameters_password')
                                        elif k == 'users' or k == 'vmedia_mappings':
                                            for ka, va in v.items():
                                                for kb, vb in va.items():
                                                    if kb == 'password':
                                                        varValue = '%s_%s' % (z, vb)
                                                        sensitive_vars.append(varValue)
                                        elif k == 'snmp_users' and z == 'password':
                                            for ka, va in v.items():
                                                for kb, vb in va.items():
                                                    if kb == 'auth_password':
                                                        varValue = 'snmp_auth_%s_%s' % (z, vb)
                                                        sensitive_vars.append(varValue)
                                                    elif kb == 'privacy_password':
                                                        varValue = 'snmp_privacy_%s_%s' % (z, vb)
                                                        sensitive_vars.append(varValue)
                    for var in sensitive_vars:
                        templateVars["Variable"] = var
                        if 'ipmi_key' in var:
                            templateVars["Description"] = 'IPMI over LAN Encryption Key'
                        elif 'iscsi' in var:
                            templateVars["Description"] = 'iSCSI Boot Password'
                        elif 'local_user' in var:
                            templateVars["Description"] = 'Local User Password'
                        elif 'access_comm' in var:
                            templateVars["Description"] = 'SNMP Access Community String'
                        elif 'snmp_auth' in var:
                            templateVars["Description"] = 'SNMP Authorization Password'
                        elif 'snmp_priv' in var:
                            templateVars["Description"] = 'SNMP Privacy Password'
                        elif 'trap_comm' in var:
                            templateVars["Description"] = 'SNMP Trap Community String'
                        templateVars["varValue"] = sensitive_var_value(easy_jsonData, **templateVars)
                        templateVars["varId"] = var
                        templateVars["varKey"] = var
                        templateVars["Sensitive"] = True
                        print(f'* Adding {templateVars["Description"]} to {templateVars["workspaceName"]}')
                        terraform_cloud().tfcVariables(**templateVars)

            # tfcb_config.append({'backend':'remote','org':org})
            # name_prefix = 'dummy'
            # type = 'pools'
            # policies_p1(name_prefix, org, type).intersight(easy_jsonData, tfcb_config)
            # type = 'policies'
            # policies_p1(name_prefix, org, type).intersight(easy_jsonData, tfcb_config)
            # type = 'profiles'
            # policies_p1(name_prefix, org, type).intersight(easy_jsonData, tfcb_config)
            # type = 'ucs_domain_profiles'
            # policies_p1(name_prefix, org, type).intersight(easy_jsonData, tfcb_config)
        else:
            valid = False
            while valid == False:
                templateVars = {}
                templateVars["Description"] = f'Will You be utilizing Local or Terraform Cloud'
                templateVars["varInput"] = f'Will you be utilizing Terraform Cloud?'
                templateVars["varDefault"] = 'Y'
                templateVars["varName"] = 'Terraform Type'
                runTFCB = varBoolLoop(**templateVars)

                if runTFCB == False:
                    tfcb_config.append({'backend':'local','site':site,'tfc_organization':'default'})
                    tfcb_config.append({'policies':'','pools':'','ucs_domain_profiles':''})

                    name_prefix = 'dummy'
                    type = 'pools'
                    # policies_p1(name_prefix, site, type).intersight(easy_jsonData, tfcb_config)
                    # type = 'policies'
                    # policies_p1(name_prefix, site, type).intersight(easy_jsonData, tfcb_config)
                    # type = 'profiles'
                    # policies_p1(name_prefix, site, type).intersight(easy_jsonData, tfcb_config)
                    # type = 'ucs_domain_profiles'
                    # policies_p1(name_prefix, site, type).intersight(easy_jsonData, tfcb_config)
                    valid = True
                else:
                    valid = True

            print(f'\n-------------------------------------------------------------------------------------------\n')
            print(f'  Skipping Step to Create Terraform Cloud Workspaces.')
            print(f'  Moving to last step to Confirm the Intersight Organization Exists.')
            print(f'\n-------------------------------------------------------------------------------------------\n')
     
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
        status,json_data = get(url, tf_header, 'Get Terraform Cloud Organizations')

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
        status,json_data = get(url, tf_header, 'Get VCS Repos')

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
        status,json_data = get(url, tf_header, 'Get VCS Repos')

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
        status,json_data = get(url, tf_header, 'workspace_check')

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
        # for key in json_data['data']:
        #     print(key['attributes']['name'])
        #     if key['attributes']['name'] == templateVars['Workspace_Name']:
        #         workspace_id = key['id']
        #         key_count =+ 1

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
            json_data = post(url, payload, tf_header, template_file)

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
            json_data = patch(url, payload, tf_header, template_file)
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
            # exit()

        # print(json.dumps(json_data, indent = 4))

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
        status,json_data = get(url, tf_header, 'variable_check')

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
            json_data = post(url, payload, tf_header, template_file)

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
            json_data = patch(url, payload, tf_header, template_file)
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

# Function to get contents from URL
def get(url, site_header, section=''):
    r = ''
    while r == '':
        try:
            r = requests.get(url, headers=site_header)
            status = r.status_code

            # Use this for Troubleshooting
            if print_response_always:
                print(status)
                print(r.text)

            if status == 200 or status == 404:
                json_data = r.json()
                return status,json_data
            else:
                validating.error_request(r.status_code, r.text)

        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: %s" % (e))
            time.sleep(5)
        except Exception as e:
            print("Method %s Failed. Exception: %s" % (section[:-5], e))
            exit()

# Function to PATCH Contents to URL
def patch(url, payload, site_header, section=''):
    r = ''
    while r == '':
        try:
            r = requests.patch(url, data=payload, headers=site_header)

            # Use this for Troubleshooting
            if print_response_always:
                print(r.status_code)
                # print(r.text)

            if r.status_code != 200:
                validating.error_request(r.status_code, r.text)

            json_data = r.json()
            return json_data

        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: %s" % (e))
            time.sleep(5)
        except Exception as e:
            print("Method %s Failed. Exception: %s" % (section[:-5], e))
            exit()

# Function to POST Contents to URL
def post(url, payload, site_header, section=''):
    r = ''
    while r == '':
        try:
            r = requests.post(url, data=payload, headers=site_header)

            # Use this for Troubleshooting
            if print_response_always:
                print(r.status_code)
                # print(r.text)

            if r.status_code != 201:
                validating.error_request(r.status_code, r.text)

            json_data = r.json()
            return json_data

        except requests.exceptions.ConnectionError as e:
            print("Connection error, pausing before retrying. Error: %s" % (e))
            time.sleep(5)
        except Exception as e:
            print("Method %s Failed. Exception: %s" % (section[:-5], e))
            exit()
