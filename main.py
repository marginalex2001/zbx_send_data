from zabbix_utils import ZabbixAPI
import os
from dotenv import load_dotenv
import pandas as pd

def main():
    table = pd.read_excel(TABLE)
    a = ZabbixAddhost(API_TOKEN,table,ZABBIX_IP)
    a.create_host()

class ZabbixAddhost():
    def __init__(self, api_token, table, api_url):
        self.__api_token = api_token
        self.__api_url = api_url
        # self.__table = table
        self.__hosts = table.to_dict(orient='records')
        self.__zbx = ZabbixAPI(url=self.__api_url)

        self.__zbx.login(token=self.__api_token)

    def prnt(self):
        self.__create_json_to_add_host()
        pass
    
    def create_host(self):
        data = self.__create_json_to_add_host()
        for host in data:
            try:
                print(self.__zbx.host.create(host))                                            #Печатает ID хоста, потом надо будет записывать их в таблицу О_о
            except Exception as e:
                print(f"An error occurred when attempting to create items: {str(e)}")

    def __create_json_to_add_host(self):
        data = []
        for host in self.__hosts:                                                                                      #Добавить проверку на NaN; добавить проверку на SNMP и добавлять community
            description = str(
                f"Model: {host['Host model']}\n"
                f"MAC: {host['MAC']}\n"
                f"Rack: {host['Rack']}\n"
                f"S/N: {host['Serial number']}\n"
                f"Inventory: {host['Inventory number']}\n"
            )

            host_sum = {
                "host" : host['Hostname'],                                                                      #Имя хоста
                "groups" : self.__groups_name_to_ID(host['Group']),
                "status" : host['Status'],                                                                      #Статус (0 - активирован/1 - деактивирован) добавить в таблицу enabled/disabled
                "interfaces": [
                    {
                        "type": host['Type (Agent/SNMP)'],                                                      #1 - agent, 2 - SNMP; добавить текстовое описание
                        "main": 1,                          
                        "useip": 1,
                        "ip": host["IP address"],                                                               #IP адрес 
                        "dns": "",                                                                              #DNS имя
                        "port": host['Port']                                                                    #Порт мониторинга
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
    load_dotenv()
    API_TOKEN = os.getenv("API_TOKEN")                                              #Получение токена из .env
    ZABBIX_IP = os.getenv("ZABBIX_URL")                                             #Получение IP адреса из .env
    TABLE = os.getenv("TABLE_NAME")                                                 #Получение имени таблицы из .env
    main()