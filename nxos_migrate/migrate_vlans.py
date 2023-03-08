#!/usr/bin/env python3
import json
import os
import pexpect
import re
import sys
import yaml
password = os.getenv('password')
username = os.getenv('username')
ydata = yaml.safe_load(open(sys.argv[1], 'r'))
sys_shell = os.environ['SHELL']
for switch in ydata['switches']:
    child = pexpect.spawn(sys_shell, encoding='utf-8')
    child.logfile_read = sys.stdout
    child.sendline(f'ssh {username}@{switch} | tee {switch}.txt')
    child.expect(f'tee {switch}.txt')
    logged_in = False
    while logged_in == False:
        i = child.expect(['Are you sure you want to continue', 'Password:', '[a-zA-Z0-9\-\_]+#'])
        if i == 0: child.sendline('yes')
        elif i == 1: child.sendline(password)
        elif i == 2: logged_in = True
    child.sendline('term length 0')
    child.expect('[a-zA-Z0-9\-\_]+#')
    child.sendline('show hsrp brief | incl Vlan')
    child.expect('[a-zA-Z0-9\-\_]+#')
    child.sendline("show run interface | sec Vlan | egrep 'interface|hsrp|ip ' | egrep -v 'dhcp|red|ospf'")
    child.expect('[a-zA-Z0-9\-\_]+#')
    child.sendline("show ip arp vrf all | incl Vlan")
    child.expect('[a-zA-Z0-9\-\_]+#')
    child.sendline('exit')
    child.close()

idict = {'vlans':{}}
re_arp = re.compile(r'^([\d\.]+)[ ]+[\d:]+[ ]+[a-z\d\.]+[ ]+(Vlan[\d]+) ')
re_hsrp_1 = re.compile(r'^  (Vlan[\d]+)[ ]+[ \d]+P (Active|Standby)[ ]+(local|[\d\.]+)[ ]+(local|[\d\.]+)[ ]+([\d\.]+) ')
re_hsrp_2 = re.compile(r'^    ip ([\d\.]+)$')
re_ip   = re.compile(r'  ip address ([\d\.]+)/[\d]+$')
re_vlan = re.compile(r'^interface (Vlan[\d]+)')
for switch in ydata['switches']:
    sw_file = open(f'{switch}.txt', 'r')
    for line in sw_file:
        if re.search(re_hsrp_1, line):
            regex = re.search(re_hsrp_1, line)
            idict['vlans'][regex.group(1)] = {
                'hsrp':regex.group(5),'primary':regex.group(3),'secondary':regex.group(4)
            }
        elif re.search(re_vlan, line):
            regex = re.search(re_vlan, line)
            vlan = regex.group(1)
            if not idict['vlans'].get(vlan):
                idict['vlans'][vlan] = {'hsrp':'','primary':'','secondary':''}
        elif re.search(re_ip, line):
            regex = re.search(re_ip, line)
            if re.search('\d', idict['vlans'][vlan]['primary']):
                primary = idict['vlans'][vlan]['primary']
                new = regex.group(1)
                last_1 = int(primary.split('.')[3])
                last_2 = int(new.split('.')[3])
                if last_1 < last_2:
                    idict['vlans'][vlan]['secondary'] = new
                else:
                    idict['vlans'][vlan]['primary'] = new
                    idict['vlans'][vlan]['secondary'] = primary
            else:
                idict['vlans'][vlan]['primary'] = regex.group(1)
        elif re.search(re_hsrp_2, line):
            regex = re.search(re_hsrp_2, line)
            idict['vlans'][vlan]['hsrp'] = regex.group(1)
        elif re.search(re_arp, line):
            regex = re.search(re_arp, line)
            if not idict['vlans'].get(regex.group(2)): idict['vlans'][regex.group(2)] = {}
            if not idict['vlans'][regex.group(2)].get('arp'):
                idict['vlans'][regex.group(2)]['arp'] = []
            idict['vlans'][regex.group(2)]['arp'].append(regex.group(1))
    sw_file.close()
    #os.remove(sw_file)

class MyDumper(yaml.Dumper):
    def increase_indent(self, flow=False, indentless=False):
        return super(MyDumper, self).increase_indent(flow, False)
yaml_file = 'full_table.yaml'
if not os.path.exists(yaml_file):
    create_file = f'type nul >> {yaml_file}'
    os.system(create_file)
wr_file = open('full_table.yaml', 'w')
wr_file.write(yaml.dump(idict, Dumper=MyDumper, default_flow_style=False))
wr_file.close()
vlanDict = {'remove_vlans':[]}
for k, v in idict['vlans'].items():
    if v.get('arp'):
        if len(v['arp']) < 4:
            vlanDict['remove_vlans'].append(k)
    else: vlanDict['remove_vlans'].append(k)
yaml_file = 'remove_list.yaml'
if not os.path.exists(yaml_file):
    create_file = f'type nul >> {yaml_file}'
    os.system(create_file)
wr_file = open('remove_list.yaml', 'w')
wr_file.write(yaml.dump(vlanDict, Dumper=MyDumper, default_flow_style=False))
wr_file.close()
exit()
