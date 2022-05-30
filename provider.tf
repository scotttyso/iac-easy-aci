terraform {
  required_version = ">= 1.1.0"
  required_providers {
    aci = {
      source  = "CiscoDevNet/aci"
      version = ">= 2.1.0"
    }
    mso = {
      source  = "CiscoDevNet/mso"
      version = ">= 0.6.0"
    }
  }
}

provider "aci" {
  cert_name   = var.certName
  password    = var.apicPass
  private_key = var.privateKey
  url         = "https://${var.apicHostname}"
  username    = var.apicUser
  insecure    = true
}

provider "mso" {
  domain   = var.ndoDomain
  insecure = true
  password = var.ndoPass
  platform = "nd"
  url      = "https://${var.ndoHostname}"
  username = var.ndoUser
}
