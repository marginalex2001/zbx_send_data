from zabbix_utils import ZabbixAPI
import os
import pandas as pd
import json
from termcolor import colored

def main():
    API_TOKEN = os.getenv("API_TOKEN")                                              #Получение токена из .env
    ZABBIX_IP = os.getenv("ZABBIX_URL")                                             #Получение IP адреса из .env
    TABLE = os.getenv("TABLE_NAME")                                                 #Получение имени таблицы из .env
    if API_TOKEN == '' or ZABBIX_IP == '':
        print(colored("ERROR. ZABBIX_API OR API_TOKEN CAN'T BE EMPTY!"))
        exit()
    
    table_exists = True

    if TABLE != '':
        try:
            table = pd.read_excel(TABLE)
        except Exception as e:
            print(f"ERROR WHILE OPEN TABLE")
        zbx = ZabbixAddHost(api_token=API_TOKEN,api_url=ZABBIX_IP,tablename=TABLE,table=table)
    else: 
        zbx = ZabbixAddHost(api_token=API_TOKEN,api_url=ZABBIX_IP)
                
    while True:
        msg = \
f"------------------------------------------\n\
Zabbix create host script:\n\
Choose option:\n\
    1 - Get hosts from zabbix server.\n\
    2 - Upload hosts from table to server.\n\
    3 - {colored("Exit", 'red')}\n\
------------------------------------------\n"
        try:
            i = int(input(msg))
        except Exception as e:
            print (colored("ERROR VALUE. TRY AGAIN.", 'red'))
            continue
        if i == 1:
            tablename = zbx.get_hosts()
            msg_good = 'All hosts recorded in ' + tablename
            print(colored(msg_good, 'green'))
        elif i == 2:
            zbx.create_hosts()
            msg_good = 'Hosts uploaded'
            print(colored(msg_good, 'green'))
        elif i == 3:
            print(colored("EXIT", 'red'))
            return False
        else:
            print(colored("ERROR VALUE. TRY AGAIN.", 'red'))
    
    # a.get_hosts()

class ZabbixAddHost():
    def __init__(self, api_token: str, api_url: str, tablename: str = "hosts.xlsx", table: pd.DataFrame | None = None):
        """
        Class for create zabbix hosts from `.xlsx` files and and vice versa\n
        `api_token` - Your ZABBIX API token, required for create\n
        `api_url` - IP-address of Your ZABBIX SERVER *(the address is specified without the protocol, but with the port and path indicated, for example: `127.0.0.1:8000/zabbix`)*, required for create\n
        `tablename` - Name of Your .xlsx file. **Must contain file extention.** *For example: `hosts.xlsx`*, optional for create\n
        `table` - Pandas dataframe with your hosts. Must contain the following columns: *'Hostname', 'Group', 'Status', 'IP address', 'DNS', 'Port', 'Type (Agent/SNMP)', 'Community', 'SNMP version', 'Template', 'Host type', 'Host model', 'OS', 'Inventory number', 'MAC', 'Rack', 'Serial number', 'hostid'*. Optional for create
        """
        self.__api_token = str(api_token)
        self.__api_url = str(api_url)
        self.__tablename = str(tablename)
        if table is None:
            self.__table = self.__create_table()
            print(f"Created table {self.__tablename}")
        elif isinstance(table, pd.DataFrame):
            self.__table = table.fillna('')
        else:
            msg = f'ERROR, TABLE "{tablename.upper()}" IS NOT PANDAS DATAFRAME'
            print(f"{'-'*len(msg)}\n{colored(msg, 'red')}\n{'-'*len(msg)}")
            exit()
        self.__hosts: list = self.__table.to_dict(orient='records')
        self.__zbx = ZabbixAPI(token=self.__api_token, url=self.__api_url)
        self.__groupids = self.__groups_names_to_IDs()
        self.__templateids = self.__templates_names_to_IDs()
        

    def __create_table(self) -> pd.DataFrame:
        """
        Creating an `.xlsx` file with the fields needed for the `ZabbixAddHost` class
        """
        cols = ['Hostname', 'Group', 'Status', 'IP address', 'DNS', 'Port', 'Type (Agent/SNMP)', 'Community', 'SNMP version', 'Template', 'Host type', 'Host model', 'OS', 'Inventory number', 'MAC', 'Rack', 'Serial number', 'hostid']
        df = pd.DataFrame(columns=cols)
        try:
            df.to_excel(self.__tablename, index=False)
        except Exception as e:
            self.__table_write_exception(e=e)
        return df
    
    def __table_write_exception(self, e: Exception):
            filename = 'error.txt'
            try:
                with open(filename, 'w') as file:
                    file.write(self.__table.to_string())
                msg = f"Error: {e}, dataframe written in {filename}"
            except Exception as f:
                msg = f"Error: {e}, dataframe not written: {f}"
            print(colored(msg, 'red'))
        
    def get_hosts(self) -> str:
        """
        Get list of hosts from zabbix server, and write it to `.xlsx`
        """
        src_data: list = self.__zbx.host.get ({                
            "output" : ["name", "hostid", "status"],            # Hostname,  hostid, status
            "selectHostGroups" : ["name"],                      # Hostgroup name            
            "selectParentTemplates" : ["name"],                 # Host Template
            "selectInterfaces" : ["ip", "dns", "port",
                                  "type", "details"],           # SNMP/Agent,   if SNMP -> community and version
            "selectInventory": ["os", "macaddress_a", "model", "tag", "type", "serialno_a", "location"]
        })
        to_table_data = self.__preparing_to_xlsx(src=src_data)
        df_new_data = pd.DataFrame(to_table_data)
        df_new_data['hostid'] = df_new_data['hostid'].astype(int)
        try:
            self.__table['hostid'] = self.__table['hostid'].astype(int)
            df_new_data_unique = df_new_data[~df_new_data['hostid'].isin(self.__table['hostid'])]       #if hostid in self.__table exist, record remove
        except:
            df_new_data_unique = df_new_data
        if df_new_data_unique.empty == False:
            self.__table = pd.concat([self.__table, df_new_data_unique], ignore_index=True)
            try:
                self.__table.to_excel(self.__tablename, index=False)
            except Exception as e:
                self.__table_write_exception(e)
        return str(self.__tablename)
        
    def __preparing_to_xlsx(self, src: list) -> list:
        """
        `src` conversion to new format for writing to `.xlsx` table
        """
        result = []
        for host in src:
            hostid, name, status = host['hostid'], host['name'], str()
            if host['status'] == '0':
                status = 'Enabled'
            elif host['status'] == '1':
                status = 'Disabled'
            hostgroups = {g['name'] for g in host['hostgroups']} 
            ips, int_type, ports, dnss = [], [], [], []
            snmp_community, snmp_version = [], []
            if host['interfaces']:
                for interface in host['interfaces']:
                    if interface['ip']:
                        ips.append(interface['ip'])
                    if interface['dns']:
                        dnss.append(interface['dns'])
                    if interface['port']:
                        ports.append(interface['port'])
                    if interface['type']:
                        if interface['type'] == '1':
                            int_type.append('Agent')
                        elif interface['type'] == '2':
                            int_type.append('SNMP')
                        elif interface['type'] == '3':
                            int_type.append('IPMI')
                        elif interface['type'] == '4':
                            int_type.append('JMX')
                    if interface['details']:
                        snmp_community.append(interface['details']['community'])
                        snmp_version.append(interface['details']['version'])
            if host['parentTemplates']:
                hosttemplates = {t['name'] for t in host['parentTemplates']}
            hosttemplates_str = ", ".join(hosttemplates) if hosttemplates else ''
            hostgroups_str = ", ".join(hostgroups) if hostgroups else ''
            ips_str = ", ".join(ips) if ips else ''
            int_type_str = ", ".join(int_type) if int_type else ''
            ports_str = ", ".join(ports) if ports else ''
            dns_str = ", ".join(dnss) if dnss else ''
            snmp_community_str = ", ".join(snmp_community) if snmp_community else ''
            snmp_version_str = ", ".join(snmp_version) if snmp_version else ''
            inventory_data = dict()
            if host['inventory']:
                inventory = host['inventory']
                inventory_data['type'] = inventory['type'] if inventory['type'] else ''
                inventory_data['model'] = inventory['model'] if inventory['model'] else ''
                inventory_data['tag'] = inventory['tag'] if inventory['tag'] else ''
                inventory_data['serialno_a'] = inventory['serialno_a'] if inventory['serialno_a'] else ''
                inventory_data['os'] = inventory['os'] if inventory['os'] else ''
                inventory_data['macaddress_a'] = inventory['macaddress_a'] if inventory['macaddress_a'] else ''
                inventory_data['location'] = inventory['location'] if inventory['location'] else ''

            result.append({
                "Hostname": name,
                "Group": hostgroups_str,
                "Status": status,
                "IP address": ips_str,
                "DNS": dns_str,
                "Port": ports_str,
                "Type (Agent/SNMP)": int_type_str,
                "Community": snmp_community_str,
                "SNMP version": snmp_version_str,
                "Template": hosttemplates_str,
                "Host type": inventory_data['type'] if host['inventory'] else '',
                "Host model": inventory_data['model'] if host['inventory'] else '',
                "OS": inventory_data['os'] if host['inventory'] else '',
                "Inventory number": inventory_data['tag'] if host['inventory'] else '',  # Используем 'tag' для поля Inventory number
                "MAC": inventory_data['macaddress_a'] if host['inventory'] else '',
                "Rack": inventory_data['location'] if host['inventory'] else '',
                "Serial number": inventory_data['serialno_a'] if host['inventory'] else '',
                "hostid": hostid
            })
        return result
    
    def create_hosts(self):
        """
        Creating host on zabbix server from `.xlsx` file
        """
        data = self.__data_to_json()
        for host in data:
            try:
                id = self.__zbx.host.create(host)
                id = id['hostids'][0]
                print(f"Host created. Host id is {id}")
                self.__table.loc[self.__table['Hostname'] == host['host'], 'hostid'] = int(id)
                
            except Exception as e:
                print(f"An error occurred when attempting to create items: {str(e)}")
        try:
            self.__table.to_excel(self.__tablename, index=False)
        except Exception as e:
            self.__table_write_exception(e)

    def __data_to_json(self) -> list:
        """
        Get values from `self.__hosts` and create valid JSON for Zabbix API
        """
        hosts_to_create = []
        for host in self.__hosts:
            description = str(
                    f"Model: {host['Host model']}\n"
                    f"MAC: {host['MAC']}\n"
                    f"Rack: {host['Rack']}\n"
                    f"S/N: {host['Serial number']}\n"
                    f"Inventory: {host['Inventory number']}\n"
                )
            hostname = host['Hostname']
            group = self.__name_to_id(self.__groupids,host['Group'],"groupid")  #host['Group'] # [{"groupid": "4"}, {"groupid": "29"}]
            templates = self.__name_to_id(self.__templateids,host['Template'], "templateid") #host['Template']# self.__templateids[f"{host['Template']}"]
            status = 0 if host['Status'] == 1 or host['Status'] == 'Enabled' else 1                             #Статус (0 - активирован/1 - деактивирован)
            host = {
                "host" : hostname,
                "groups" : group,
                "templates": templates,
                "status" : status,
                "description" : description,
                "inventory":{
                    "macaddress_a": str(host['MAC']),                                                               #MAC адрес
                    "type": str(host['Host type']),                                                                 #Тип хоста (сервер/свитч)
                    "name": str(host['Hostname']),                                                                  #Имя (для инвентаризации)
                    "os" : str(host['OS']),                                                                         #ОС
                    "tag" : str(host['Inventory number']),                                                          #Инв. номер (Тэг)
                    "location" : str(host['Rack']),                                                                 #Местоположение
                    "model" : str(host['Host model']),                                                              #Модель
                    "serialno_a" : str(host['Serial number'])                                                       #Серийник
                },
                "interfaces" : self.__create_interface(host['Type (Agent/SNMP)'],host['IP address'],host['DNS'],host['Port'],host['SNMP version'],host['Community'])
            }
            hosts_to_create.append(host)
        # print(json.dumps(hosts_to_create, sort_keys=True, indent=2))
        return hosts_to_create

    def __name_to_id(self, dict: dict, data: str, type: str) -> list:
        """
        Allows you to translate the ID of a value into its name.\n
        `dict` - dict key-value: `"name" : "ID"`\n
        `data` - Item names from `.xlsx`\n
        `type` - datatype (`templateid` or `hostid`)\n
        Returns JSON array of the form: `[{"type": "ID"},{"type": "ID"}]`
        """
        items_array = []
        groups_names = [item.strip() for item in data.split(',')]
        for item in groups_names:
            itemid = {
                type : dict[item]
            }
            items_array.append(itemid)
        return items_array

    def __templates_names_to_IDs(self) -> dict:
        """
        Get template names from `.xlsx` and convert them to Zabbix server template IDs 
        """
        table_templates = []
        for row in self.__hosts:
            template = [item.strip() for item in row['Template'].split(',')]
            table_templates = table_templates + template
        table_templates = list(dict.fromkeys(table_templates))
        try:
            templates = self.__zbx.template.get(
                {
                    "output": "extend",
                    "filter": {
                            "name": table_templates
                    }
                }
            )
        except Exception as e:
            print(e)
        templateids = {}
        for template in templates:
            templateids[template['name']] = template['templateid']
        return templateids

    def __groups_names_to_IDs(self) -> dict:
        """
        Get group names from `.xlsx` and convert them to Zabbix server group IDs 
        """
        table_groups = []
        for row in self.__hosts:
            group = [item.strip() for item in row['Group'].split(',')]      # Делает из строки <list> разделяя элементы через запятую
            table_groups += group
        table_groups = list(dict.fromkeys(table_groups))                    # удаляет повторы
        try:
            groups = self.__zbx.hostgroup.get(
                {
                    "output": "extend",
                    "filter": {
                            "name": table_groups
                    }
                }
            )
        except Exception as e:
            print(e)
        groupids = {}
        for group in groups:
            groupids[group['name']] = group['groupid']
        return groupids

    def __create_interface(self, type, ip, dns, port, community_ver, community) -> list:
        """
        Create list of interfaces
        """
        interfaces = []
        int_type = 2 if type == 2 or type == 'SNMP' else 1
        useip = 0 if dns != '' else 1
        interface = {
            "type" : int_type,
            "main" : 1,
            "useip" : useip,
            "ip" : ip,
            "dns" : dns,
            "port" : f"{port}"
        }
        if int_type == 2:
            details = {
                "version": int(community_ver) if community_ver else 2,
                "bulk": 0,
                "community" : f"{'{$SNMP_COMMUNITY}' if community == '' else community}"
            }
            interface["details"] = details
        interfaces.append(interface)
        return interfaces

if __name__ == "__main__":  
    main()
