#!/usr/bin/env python3

import jinja2
import os
import pkg_resources
import validating
from class_terraform import terraform_cloud
from easy_functions import process_kwargs
from easy_functions import sensitive_var_site_group
from easy_functions import write_to_site
from easy_functions import write_to_template
from openpyxl import load_workbook

aci_template_path = pkg_resources.resource_filename('class_access', 'templates/')

# Exception Classes
class InsufficientArgs(Exception):
    pass

class ErrException(Exception):
    pass

class InvalidArg(Exception):
    pass

class LoginFailed(Exception):
    pass

class system_settings(object):
    def __init__(self, type):
        self.templateLoader = jinja2.FileSystemLoader(
            searchpath=(aci_template_path + '%s/') % (type))
        self.templateEnv = jinja2.Environment(loader=self.templateLoader)
        self.type = type

