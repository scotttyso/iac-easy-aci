import getpass
import jinja2
import json
import os, re, sys
import pkg_resources
import requests
import time
import validating

# Global options for debugging
print_payload = False
print_response_always = True
print_response_on_fail = True

# Log levels 0 = None, 1 = Class only, 2 = Line
log_level = 2

# Global path to main Template directory
tf_template_path = pkg_resources.resource_filename('lib_terraform', 'Terraform/')

# Exception Classes
class InsufficientArgs(Exception):
    pass

class ErrException(Exception):
    pass

class InvalidArg(Exception):
    pass

class LoginFailed(Exception):
    pass

# Terraform Cloud For Business - Policies
# Class must be instantiated with Variables
class Terraform_Cloud(object):
    def __init__(self):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(tf_template_path + 'Terraform_Cloud/'))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def terraform_token(self):
        # -------------------------------------------------------------------------------------------------------------------------
        # Check to see if the TF_VAR_terraform_cloud_token is already set in the Environment, and if not prompt the user for Input.
        #--------------------------------------------------------------------------------------------------------------------------
        if os.environ.get('TF_VAR_terraform_cloud_token') is None:
            print(f'\n----------------------------------------------------------------------------------------\n')
            print(f'  The Run or State Location was set to Terraform_Cloud.  To Store the Data in Terraform')
            print(f'  Cloud we will need a User or Org Token to authenticate to Terraform Cloud.  If you ')
            print(f'  have not already obtained a token see instructions in how to obtain a token Here:\n')
            print(f'   - https://www.terraform.io/docs/cloud/users-teams-organizations/api-tokens.html\n')
            print(f'  Please Select "C" to Continue or "Q" to Exit:')
            print(f'\n----------------------------------------------------------------------------------------\n')

            while True:
                user_response = input('  Please Enter ["C" or "Q"]: ')
                if re.search('^C$', user_response):
                    break
                elif user_response == 'Q':
                    exit()
                else:
                    print(f'\n-----------------------------------------------------------------------------\n')
                    print(f'  A Valid Response is either "C" or "Q"...')
                    print(f'\n-----------------------------------------------------------------------------\n')

            # Request the TF_VAR_terraform_cloud_token Value from the User
            while True:
                try:
                    secure_value = getpass.getpass(prompt=f'Enter the value for the Terraform Cloud Token: ')
                    break
                except Exception as e:
                    print('Something went wrong. Error received: {}'.format(e))

            # Add the TF_VAR_terraform_cloud_token to the Environment
            os.environ['TF_VAR_terraform_cloud_token'] = '%s' % (secure_value)
            terraform_cloud_token = secure_value
        else:
            terraform_cloud_token = os.environ.get('TF_VAR_terraform_cloud_token')

        return terraform_cloud_token

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def oath_token(self, **kwargs):
        # Dicts for Application Profile; required and optional args
        # Dicts for required and optional args
        required_args = {'VCS_Base_Repo': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        # -------------------------------------------------------------------------------------------------------------------------
        # Check to see if the TF_VAR_terraform_cloud_token is already set in the Environment, and if not prompt the user for Input.
        #--------------------------------------------------------------------------------------------------------------------------
        if os.environ.get('TF_VAR_terraform_oath_token') is None:
            print(f'\n----------------------------------------------------------------------------------------\n')
            print(f'  The Run or State Location was set to Terraform_Cloud.  The Script will Create Workspaces')
            print(f'  using the repo {templateVars["VCS_Base_Repo"]}.  You will have needed to register this ')
            print(f'  repository in your workspace prior to running this script and obtain an OAuth Token ID:\n')
            print(f'   - https://www.terraform.io/docs/cloud/vcs/github.html\n')
            print(f'  Please Select "C" to Continue or "Q" to Exit:')
            print(f'\n----------------------------------------------------------------------------------------\n')
            while True:
                user_response = input('  Please Enter ["C" or "Q"]: ')
                if re.search('^C$', user_response):
                    break
                elif user_response == 'Q':
                    exit()
                else:
                    print(f'\n-----------------------------------------------------------------------------\n')
                    print(f'  A Valid Response is either "C", "Q"...')
                    print(f'\n-----------------------------------------------------------------------------\n')

                # Request the TF_VAR_terraform_oath_token Value from the User
            while True:
                try:
                    secure_value = getpass.getpass(prompt=f'Enter the value for the Terraform OAuth Token ID: ')
                    break
                except Exception as e:
                    print('Something went wrong. Error received: {}'.format(e))

            # Add the TF_VAR_terraform_oath_token to the Environment
            os.environ['TF_VAR_terraform_oath_token'] = '%s' % (secure_value)
            terraform_oath_token = secure_value
        else:
            # Obtain the TF_VAR_terraform_oath_token from the User Environment
            terraform_oath_token = os.environ.get('TF_VAR_terraform_oath_token')

        return terraform_oath_token


    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def tf_workspace(self, **kwargs):
        # Dicts for required and optional args:
        required_args = {'Terraform_Cloud_Org': '',
                         'Terraform_Version': '',
                         'Terraform_Agent_Pool_ID': '',
                         'terraform_cloud_token': '',
                         'terraform_oath_token': '',
                         'Working_Directory': '',
                         'Workspace_Name': '',
                         'VCS_Base_Repo': '',}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

        #-------------------------------
        # Configure the Workspace URL
        #-------------------------------
        url = 'https://app.terraform.io/api/v2/organizations/%s/workspaces/%s' %  (templateVars['Terraform_Cloud_Org'], templateVars['Workspace_Name'])
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
            if json_data['data']['attributes']['name'] == templateVars['Workspace_Name']:
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
        if not key_count > 0:
            #-------------------------------
            # Configure the Workspace URL
            #-------------------------------
            url = 'https://app.terraform.io/api/v2/organizations/%s/workspaces/' %  (templateVars['Terraform_Cloud_Org'])
            tf_token = 'Bearer %s' % (templateVars['terraform_cloud_token'])
            tf_header = {'Authorization': tf_token,
                    'Content-Type': 'application/vnd.api+json'
            }

            # Define the Template Source
            template_file = 'workspace_post.json'
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
            template_file = 'workspace_patch.json'
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
            print(f"\n   Unable to Determine the Workspace ID for {templateVars['Workspace_Name']}.")
            print(f"\n   Exiting...")
            print(f'\n-----------------------------------------------------------------------------\n')
            exit()

        os.environ[templateVars['Workspace_Name']] = '%s' % (workspace_id)

        # print(json.dumps(json_data, indent = 4))
        return workspace_id

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def tf_variable(self, **kwargs):
        # Dicts for required and optional args:
        required_args = {'Terraform_Cloud_Org': '',
                         'Terraform_Version': '',
                         'terraform_cloud_token': '',
                         'workspace_id': '',
                         'Variable': '',
                         'Var_Value': '',
                         'HCL': '',
                         'Sensitive': '',}
        optional_args = {'Description': ''}

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)

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
            template_file = 'var_post.jinja2'
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
            template_file = 'var_patch.jinja2'
            template = self.templateEnv.get_template(template_file)

            # Create the Payload
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

    # Method must be called with the following kwargs.
    # Please Refer to the Input Spreadsheet "Notes" in the relevant column headers
    # for Detailed information on the Arguments used by this Method.
    def var_value(self, **kwargs):
        # Dicts for required and optional args:
        required_args = {'Run_Location': '',
                         'State_Location': '',
                         'Variable': ''}
        optional_args = { }

        # Validate inputs, return dict of template vars
        templateVars = process_kwargs(required_args, optional_args, **kwargs)
        sensitive_var = 'TF_VAR_%s' % (templateVars['Variable'])

        # -------------------------------------------------------------------------------------------------------------------------
        # Check to see if the Variable is already set in the Environment, and if not prompt the user for Input.
        #--------------------------------------------------------------------------------------------------------------------------
        if os.environ.get(sensitive_var) is None and templateVars['State_Location'] == 'Local':
            print(f"\n---------------------------------------------------------------------------------------\n")
            print(f"  The State Location is set to {templateVars['State_Location']}, which means that sensitive ")
            print(f"  variablesneed to be stored locally in the Environment Variables.  The Script did not")
            print(f"  find {templateVars['Variable']} as an 'environment' variable.  To not be prompted for the")
            print(f"  value of {sensitive_var} each time add the following to your local environemnt:\n")
            print(f"   - export {sensitive_var}='{templateVars['Variable']}_value'")
            print(f"\n---------------------------------------------------------------------------------------\n")
        elif os.environ.get(sensitive_var) is None:
            print(f"\n----------------------------------------------------------------------------------\n")
            print(f"  The Run_Location is set to {templateVars['Run_Location']}.  The Script did not find ")
            print(f"  {sensitive_var} as an 'environment' variable.  To not be prompted for the value of ")
            print(f"  {templateVars['Variable']} each time add the following to your local environemnt:\n")
            print(f"   - export {sensitive_var}='{templateVars['Variable']}_value'")
            print(f"\n----------------------------------------------------------------------------------\n")

        if os.environ.get(sensitive_var) is None:
            while True:
                try:
                    secure_value = getpass.getpass(prompt=f'Enter the value for {templateVars["Variable"]}: ')
                    break
                except Exception as e:
                    print('Something went wrong. Error received: {}'.format(e))

            # Add the Variable to the Environment
            os.environ[sensitive_var] = '%s' % (secure_value)

            var_value = secure_value

        else:
            # Add the Variable to the Environment
            var_value = os.environ.get(sensitive_var)

        return var_value

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
                # print(r.text)

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

# Function to validate input for each method
def process_kwargs(required_args, optional_args, **kwargs):
    # Validate all required kwargs passed
    # if all(item in kwargs for item in required_args.keys()) is not True:
    #    error_ = '\n***ERROR***\nREQUIRED Argument Not Found in Input:\n "%s"\nInsufficient required arguments.' % (item)
    #    raise InsufficientArgs(error_)
    error_count = 0
    error_list = []
    for item in required_args:
        if item not in kwargs.keys():
            error_count =+ 1
            error_list += [item]
    if error_count > 0:
        error_ = '\n\n***Begin ERROR***\n\n - The Following REQUIRED Key(s) Were Not Found in kwargs: "%s"\n\n****End ERROR****\n' % (error_list)
        raise InsufficientArgs(error_)

    error_count = 0
    error_list = []
    for item in optional_args:
        if item not in kwargs.keys():
            error_count =+ 1
            error_list += [item]
    if error_count > 0:
        error_ = '\n\n***Begin ERROR***\n\n - The Following Optional Key(s) Were Not Found in kwargs: "%s"\n\n****End ERROR****\n' % (error_list)
        raise InsufficientArgs(error_)

    # Load all required args values from kwargs
    error_count = 0
    error_list = []
    for item in kwargs:
        if item in required_args.keys():
            required_args[item] = kwargs[item]
            if required_args[item] == None:
                error_count =+ 1
                error_list += [item]

    if error_count > 0:
        error_ = '\n\n***Begin ERROR***\n\n - The Following REQUIRED Key(s) Argument(s) are Blank:\nPlease Validate "%s"\n\n****End ERROR****\n' % (error_list)
        raise InsufficientArgs(error_)

    for item in kwargs:
        if item in optional_args.keys():
            optional_args[item] = kwargs[item]
    # Combine option and required dicts for Jinja template render
    templateVars = {**required_args, **optional_args}
    return(templateVars)

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
