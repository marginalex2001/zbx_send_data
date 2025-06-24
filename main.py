from zabbix_utils import ZabbixAPI
import os
from dotenv import load_dotenv
import pandas as pd

def main():
    load_dotenv()
    API_TOKEN = os.getenv("API_TOKEN")                                              #Получение токена из .env
    ZABBIX_IP = os.getenv("ZABBIX_URL")                                             #Получение IP адреса из .env
    TABLE = os.getenv("TABLE_NAME")                                                 #Получение имени таблицы из .env
    table = pd.read_excel(TABLE)
    a = ZabbixAddhost(API_TOKEN,table,ZABBIX_IP)
    a.create_host(TABLE)

class ZabbixAddhost():
    def __init__(self, api_token, table, api_url):
        self.__api_token = api_token
        self.__api_url = api_url
        self.__table = table.fillna('')
        self.__hosts = self.__table.to_dict(orient='records')
        self.__zbx = ZabbixAPI(url=self.__api_url)

        self.__zbx.login(token=self.__api_token)

    def prnt(self):
        pass
    
    def create_host(self, tablename):
        data = self.__create_json_to_add_host()
        for host in data:
            try:
                id = self.__zbx.host.create(host)                                           #Печатает ID хоста, потом надо будет записывать их в таблицу О_о
                id = id['hostids'][0]
                print(f"Host created. Host id is {id}")
                self.__table.loc[self.__table['Hostname'] == host['host'], 'hostid'] = int(id)
                
            except Exception as e:
                print(f"An error occurred when attempting to create items: {str(e)}")
        try:
            self.__table.to_excel(tablename, index=False)
        except Exception as e:
            print(f"An error while renew table: {str(e)}")
            try:
                self.__table.to_excel("Error.xlsx", index=False)
            except Exception as err:
                print(f"FATAL ERROR WHILE WRITTING TABLE. NO BACKUP")


    def __create_json_to_add_host(self):
        data = []
        for host in self.__hosts:                                                                                      #Добавить проверку на NaN; добавить проверку на SNMP и добавлять community
            if host['hostid'] == '' and (host['DNS'] != '' or host['IP address'] != ''):
                description = str(
                    f"Model: {host['Host model']}\n"
                    f"MAC: {host['MAC']}\n"
                    f"Rack: {host['Rack']}\n"
                    f"S/N: {host['Serial number']}\n"
                    f"Inventory: {host['Inventory number']}\n"
                )

                host_sum = {
                    "host" : host['Hostname'],                                                                                  #Имя хоста
                    "groups" : self.__groups_name_to_ID(host['Group']),
                    "status" : f"{0 if host['Status'] == 1 or host['Status'] == 'Enabled' else 1}",                             #Статус (0 - активирован/1 - деактивирован) добавить в таблицу enabled/disabled  host['Status']
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
                    "templates": self.__templates_name_to_ID(host['Template']),
                    "inventory_mode": 1,                                                                                    #Possible values: '-1' - (default) disabled; '0' - manual; '1' - automatic.
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
                    "description" : description
                }
                data.append(host_sum)
            if (host['DNS'] == '' and host['IP address'] == ''):
                print(f"Error while create host: no such IP-address or DNS for {host['Hostname']} [IP Address: {host['IP address']}; DNS: {host['DNS']}]")
        return data

    def __groups_name_to_ID(self, groups):
        parts = [part.strip() for part in groups.split(',')]
        b = []
        for part in parts:
            try:
                a = self.__zbx.hostgroup.get(
                    {
                        "output": ["groupid"],
                        "filter": 
                        {
                            "name": [part]
                        }
                    }
                )
            except Exception as e:
                print(f"An error occurred when get template items: {str(e)}")
            b = b + a
        return b

    def __templates_name_to_ID(self, templates):
        parts = [part.strip() for part in templates.split(',')]
        b = []
        for part in parts:
            try:
                a = self.__zbx.template.get(
                    {
                        "output": "templateid",
                        "filter": 
                        {
                            "host": [part]
                        }
                    }
                )
            except Exception as e:
                print(f"An error occurred when get template items: {str(e)}")
            b = b + a
        return b

    def __set_host_type(self):
        pass

if __name__ == "__main__":
    main()
