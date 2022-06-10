# IAC - Easy ACI Python Wrapper

## Updates/News

* Note: First Initial Release.  Looking for Testers

## Pre-Requirements

* Deploy New/Existing ACI Fabrics using Terraform via a Python wrapper with an Excel spreadsheet.

1. Run the Intial Configuration wizard on the APICs.

2. If Integrating with TFCB (Terraform Cloud for Business), Sign up for an Account at [Terraform Cloud](https://app.terraform.io/). Log in and generate the User API Key. You will need this when you create the TF Cloud Target in Intersight.  If not a paid version, you will need to enable the trial account.

3. Clone this repository to your own VCS Repository for the VCS Integration with Terraform Cloud.

4. Integrate your VCS Repository into the TFCB Orgnization following these instructions: [VCS Integration](https://www.terraform.io/docs/cloud/vcs/index.html).  Be sure to copy the OAth Token which you will use later on for Workspace provisioning.

## Obtain tokens and keys

### Terraform Cloud Variables

* terraform_cloud_token

  instructions: [Terraform Cloud API Tokens](https://www.terraform.io/docs/cloud/users-teams-organizations/api-tokens.html)

* tfc_organization (TFCB Organization Name)
* agent_pool (The Name of the Agent Pool in the TFCB Account).  To Create: [Intersight Service for Terraform](https://community.cisco.com/t5/data-center-and-cloud-documents/intersight-service-for-terraform/ta-p/4301093)

### APIC Variables

* apicUser - If using SSH-KEY based Authetnication, it Must be a local user.

For Certificate based Authentication

* privateKey
* certName

For User Based Authentication

* apicPass

Note: for authentication with non-Local Credentials use the following format for the user: "apic:{login_domain}\\\\{user}"

### Nexus Dashboard Orchestrator Variables

* apicUser
* apicPass

### Import the Variables into your Environment before Running the Terraform Cloud Provider module(s) in this directory

Modify the terraform.tfvars file to the unique attributes of your environment for your domain and server profiles and policies.

Once finished with the modification commit the changes to your repository.

The Following examples are for a Linux based Operating System.  Note that the TF_VAR_ prefix is used as a notification to the terraform engine that the environment variable will be consumed by terraform.

* Terraform Cloud Variables - Linux

```bash
export TF_VAR_terraform_cloud_token="your_cloud_token"
```

* Terraform Cloud Variables - Windows

```powershell
$env:TF_VAR_terraform_cloud_token="your_cloud_token"
```

### APIC/Nexus Dashboard Orchestrator Credentials

* Certificate Based Authentication - Linux

```bash
export TF_VAR_apicUser="{apic_username}"
export TF_VAR_certName="{name_of_certificate_associated_to_the_user}"
export TF_VAR_privateKey=`~/Downloads/apic_private_key.txt`
```

* Certificate Based Authentication - Windows

```powershell
$env:TF_VAR_apicUser="{apic_username}"
$env:TF_VAR_certName="{name_of_certificate_associated_to_the_user}"
$env:TF_VAR_privateKey="$HOME\Downloads\apic_private_key.txt"
```

* User Based Authentication - Linux

```bash
export TF_VAR_apicUser="{apic_username}"
export TF_VAR_apicPass="{user_password}"
export TF_VAR_ndoUser="{ndo_username}"
export TF_VAR_ndoPass="{user_password}"
```

* User Based Authentication - Windows

```powershell
$env:TF_VAR_apicUser="{apic_username}"
$env:TF_VAR_apicPass="{user_password}"
$env:TF_VAR_ndoUser="{ndo_username}"
$env:TF_VAR_ndoPass="{user_password}"
```

### Terraform Cloud

When running in Terraform Cloud with VCS Integration, the first Plan will need to be run from the UI but subsiqent runs should trigger automatically, if auto-run is left on the workspace

### Running the Code:

* Execute the Script - Linux

```bash
./main.py {options}
```

* Execute the Script - Windows

```powershell
python main.py {options}
```

List of Options are below:

```bash
usage: main.py [-h] [-d DIR] [-wb WORKBOOK] [-ws WORKSHEET]

IaC Easy ACI Deployment Module

optional arguments:
  -h, --help            show this help message and exit
  -d DIR, --dir DIR     The Directory to use for the Creation of the Terraform Files.
  -wb WORKBOOK, --workbook WORKBOOK
                        The source Workbook.
  -ws WORKSHEET, --worksheet WORKSHEET
                        Only evaluate this single worksheet. Worksheet values are: 1. access - for Access 2. admin: for Admin 3. bridge_domains: for
                        Bridge Domains 4. contracts: for Contracts 5. epgs: for EPGs 6. fabric: for Fabric 7. l3out: for L3Out 8. port_convert: for Uplink
                        to Download Conversion 8. sites: for Sites 9. switches: for Switch Profiles 10. system_settings: for System Settings 11. tenants:
                        for Tenants 12. virtual_networking: for Virtual Networking
```

* -d - This should typically be utilized to speficy the output directory (Repo) for the Terraform Files.
* -wb - Name of the Workbook to be read.  If not specified the default is "ACI_Base_Workbookv2.xlsx
* -ws - Use this option to run the process on only a specific worksheet in the workbook
