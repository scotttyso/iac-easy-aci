#______________________________________________
#
# Date and time - Variables
#______________________________________________

date_and_time = {
  default = {
    annotation           = ""
    administrative_state = "enabled"
    authentication_keys  = [
    ]
    description          = ""
    display_format       = "local"
    master_mode          = "disabled"
    ntp_servers          = [
      {
        description              = "Richfield AD1"
        hostname                 = "10.101.128.15"
        key_id                   = "None"
        management_epg           = "default"
        management_epg_type      = "oob"
        maximum_polling_interval = 6
        minimum_polling_interval = 4
        preferred                = false
      },
      {
        description              = "Richfield AD2"
        hostname                 = "10.101.128.16"
        key_id                   = "None"
        management_epg           = "default"
        management_epg_type      = "oob"
        maximum_polling_interval = 6
        minimum_polling_interval = 4
        preferred                = true
      },
    ]
    offset_state         = "enabled"
    server_state         = "disabled"
    stratum_value        = 8
    time_zone            = "America/Detroit"
  }
}