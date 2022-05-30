variable "annotation" {
  default     = ""
  description = "workspace tag value."
  type        = string
}

variable "apicHostname" {
  default     = "apic.example.com"
  description = "Cisco APIC Hostname"
  type        = string
}

variable "apicPass" {
  default     = ""
  description = "Password for User based Authentication."
  sensitive   = true
  type        = string
}

variable "apicUser" {
  default     = "admin"
  description = "Username for User based Authentication."
  type        = string
}

variable "apic_version" {
  default     = "5.2(1g)"
  description = "The Version of ACI Running in the Environment."
  type        = string
}

variable "certName" {
  default     = ""
  description = "Cisco ACI Certificate Name for SSL Based Authentication"
  sensitive   = true
  type        = string
}

variable "ndoDomain" {
  default     = "local"
  description = "Authentication Domain for Nexus Dashboard Orchestrator Authentication."
  sensitive   = true
  type        = string
}

variable "ndoHostname" {
  default     = "nxo.example.com"
  description = "Cisco Nexus Dashboard Orchestrator Hostname"
  type        = string
}

variable "ndoPass" {
  default     = ""
  description = "Password for Nexus Dashboard Orchestrator Authentication."
  sensitive   = true
  type        = string
}

variable "ndoUser" {
  default     = "admin"
  description = "Username for Nexus Dashboard Orchestrator Authentication."
  type        = string
}

variable "privateKey" {
  default     = ""
  description = "Cisco ACI Private Key for SSL Based Authentication."
  sensitive   = true
  type        = string
}
