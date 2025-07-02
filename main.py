from zabbix_utils import ZabbixAPI
import os
import pandas as pd
import json

def main():
    API_TOKEN = os.getenv("API_TOKEN")                                              #Получение токена из .env
    ZABBIX_IP = os.getenv("ZABBIX_URL")                                             #Получение IP адреса из .env
    TABLE = os.getenv("TABLE_NAME")                                                 #Получение имени таблицы из .env
    print(TABLE)
    try:
        table = pd.read_excel(TABLE)
    except Exception as e:
        print(f"Error: {e} \n---BREAK---")
        return 1
    a = ZabbixAddhost(API_TOKEN,ZABBIX_IP,table,TABLE)
    a.get_hosts()

class ZabbixAddhost():
    def __init__(self, api_token, api_url, table, tablename):
        self.__api_token = api_token
        self.__api_url = api_url
        self.__table = table.fillna('')
        self.__hosts = self.__table.to_dict(orient='records')
        self.__zbx = ZabbixAPI(token=self.__api_token, url=self.__api_url)
        self.__groupids = self.__groups_names_to_IDs()
        self.__templateids = self.__templates_names_to_IDs()
        self.__tablename = tablename

    def get_hosts(self):
        """
        Hostname,       Group,          Status
        IP Address,     DNS,            Port
        Inerface type,  (community),    (SNMP ver)
        Template,       Hostid
        Inventory [
        Type, Model, OS, Inv, MAC, Location, S/N
        ]
        """
        src_data = self.__zbx.host.get ({
            #           Hostname  hostid    
            "output" : ["name", "hostid", "status"],
            #                   groupid
            "selectHostGroups" : ["name"],
            #                       Template
            "selectParentTemplates" : ["name"],
            "selectInterfaces" : ["ip", "dns", "port",
                                # SNMP/Agent,   if SNMP, community and version
                                  "type", "details"],
            "selectInventory": ["os", "macaddress_a", "model", "tag", "type", "serialno_a", "location"]
        })
        to_table_data = self.__preparing_to_xlsx(src=src_data)
        df_new_data = pd.DataFrame(to_table_data)
        self.__table['hostid'] = self.__table['hostid'].astype(int)
        df_new_data['hostid'] = df_new_data['hostid'].astype(int)
        df_new_data_unique = df_new_data[~df_new_data['hostid'].isin(self.__table['hostid'])]
        if df_new_data_unique.empty ==False:
            self.__table = pd.concat([self.__table, df_new_data_unique], ignore_index=True)
            self.__table.to_excel(self.__tablename, index=False)
        


    def __preparing_to_xlsx(self, src):
        result = []
        # print(json.dumps(src,sort_keys=True, indent=4))
        for host in src:
            hostid, name, status = host['hostid'], host['name'], str()
            if host['status'] == '0':
                status = 'Enabled'
            elif host['status'] == '1':
                status = 'Disabled'
            hostgroups = {g['name'] for g in host['hostgroups']} # host['hostgroups']
            # print(f"{hostgroups_str}\n{hostid}\n{name}\n{status}\n\n") #{host['inventory'] if host['inventory'] else 'oups'}\n")
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
                        # for detail in interface['details']:
                        #     print(detail['community'])
            # print(f"{ips}, {dnss}, {ports}, {int_type}, {snmp_community}, {snmp_version}, {status} \n")
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
                # print(inventory_data)

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
                "Inventory number": inventory_data['tag'] if host['inventory'] else '',  # Используем 'tag' для поля Inv
                "MAC": inventory_data['macaddress_a'] if host['inventory'] else '',
                "Rack": inventory_data['location'] if host['inventory'] else '',
                "Serial number": inventory_data['serialno_a'] if host['inventory'] else '',
                "hostid": hostid
            })
        return result

    def tests(self):
        print(self.__hosts)
        self.__data_to_json()
    
    def create_host(self):
        data = self.__data_to_json()
        for host in data:
            try:
                id = self.__zbx.host.create(host)                                           #Печатает ID хоста, потом надо будет записывать их в таблицу О_о
                id = id['hostids'][0]
                print(f"Host created. Host id is {id}")
                self.__table.loc[self.__table['Hostname'] == host['host'], 'hostid'] = int(id)
                
            except Exception as e:
                print(f"An error occurred when attempting to create items: {str(e)}")
        try:
            self.__table.to_excel(self.__tablename, index=False)
        except Exception as e:
            print(f"An error while renew table: {str(e)}")
            try:
                self.__table.to_excel("Error.xlsx", index=False)
            except Exception as err:
                print(f"FATAL ERROR WHILE WRITTING TABLE. NO BACKUP")

    def __data_to_json(self):
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

    def __name_to_id(self, dict, data, type):
        items_array = []
        groups_names = [item.strip() for item in data.split(',')]
        for item in groups_names:
            itemid = {
                type : dict[item]
            }
            items_array.append(itemid)
        return items_array

    def __templates_names_to_IDs(self):
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

    def __groups_names_to_IDs(self):
        table_groups = []
        for row in self.__hosts:
            group = [item.strip() for item in row['Group'].split(',')]
            table_groups = table_groups + group
        table_groups = list(dict.fromkeys(table_groups))
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

    def __create_interface(self, type, ip, dns, port, community_ver, community):
        """
        "type": 1,                  1 - Agent,              2 - SNMP
        "main": 1,                  1 - default interface   0 - not default interface
        "useip": 1,                 1 - IP,                 0 - DNS
        "ip": "192.168.3.1",        IP address
        "dns": "",                  DNS name
        "port": "10050"             Port
        "details": {                for SNMP connections
            "version": 2,           SNMP version:           1 - SNMPv1;     2 - SNMPv2c;      3 - SNMPv3.
            "bulk": 0,              1 - (default) - use bulk requests.      0 - don't use bulk requests;
            "community" : "public"	SNMP community
        }
        """

        '''
"interfaces": [
    {
        "type": f"{2 if host['Type (Agent/SNMP)'] == 2 or host['Type (Agent/SNMP)'] == 'SNMP' else 1}",     #1 - agent, 2 - SNMP; добавить текстовое описание     host['Type (Agent/SNMP)']
        "main": 1,                          
        "useip": f"{0 if host['DNS'] != '' else 1}",                                                        #0 - vonnect via DNS; 1 - connect via IP 
        "ip": host['IP address'],                                                                           #IP адрес 
        "dns": host['DNS'],                                                                                 #DNS имя
        "port": host['Port']                                                                                #Порт мониторинга
    }
],
        '''
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
                "version": int(community_ver),
                "bulk": 0,
                "community" : f"{'{$SNMP_COMMUNITY}' if community == '' else community}"
            }
            interface["details"] = details
        interfaces.append(interface)
        return interfaces

if __name__ == "__main__":  
    main()
