# -*- coding: UTF-8 -*-
import pypyodbc
import requests
import json
import time
import os

'''
Python使用ODBC连接SqlServer，版本对应的DRIVER名称如下：
{SQL Server} - released with SQL Server 2000
{SQL Native Client} - released with SQL Server 2005 (also known as version 9.0)
{SQL Server Native Client 10.0} - released with SQL Server 2008
{SQL Server Native Client 11.0} - released with SQL Server 2012
'''

"""
华气厚普监控系统
数据库中缺少自增主键，因此结合触发器，生成记录到新的表。
CREATE TABLE [dbo].[_ExtPayOrder] (
    [orderid] bigint NOT NULL IDENTITY(1,1) PRIMARY KEY,
    [fpno] tinyint DEFAULT (0),
    [prname] varchar(20) DEFAULT '',
    [price] decimal(4,2) DEFAULT (0),
    [vol] decimal(8,2) DEFAULT (0),
    [amt] decimal(8,2) DEFAULT (0),
    [tvol] decimal(8,2) DEFAULT (0),
    [tamt] decimal(8,2) DEFAULT (0),
    [fueltime] datetime,
    [status] tinyint DEFAULT (0),
    [neworderid] varchar(50) DEFAULT '',
);
create trigger TempGas_Insert on [tempGas]
for insert
as
begin
    INSERT INTO [_ExtPayOrder]([fpno],[prname],[price],[vol],[amt],[tvol],[tamt],[fueltime])
    select [GunNo] as fpno, '' as prname,price,NowGas as vol,NowMoney as amt,0 as tvol,0 as tamt, TradeDate as fueltime from Inserted
end;
"""

import_api = 'http://{ip}:8858/ImportService/SaveOrder'


def query_save():
    try:
        conn = pypyodbc.connect(conn_str)
        cur = conn.cursor()
        sql = 'select top 1 [orderid],[fpno],[prname],[price],[vol],[amt],[tvol],[tamt],[fueltime] FROM [_ExtPayOrder]' \
              ' where [status]=0 order by orderid asc'
        cur.execute(sql)
        rows = cur.fetchall()
        if rows:
            for row in rows:
                data = {}
                orderId = row[0]
                data['OrderId'] = orderId
                data['NozNo'] = row[1]
                data['Price'] = float(row[3])
                data['Vol'] = float(row[4])
                data['Amount'] = float(row[5])
                data['FuelTime'] = row[8].strftime("%Y-%m-%d %H:%M:%S")
                data = json.dumps(data)
                print('..............................................')
                print("--->", data)
                rsp = requests.post(url=import_api, data=data, headers={'Content-Type': 'application/json'})
                print("<---", rsp.status_code, rsp.text)
                if rsp.status_code == 200:
                    result = json.loads(rsp.text)
                    if result['IsSaved']:
                        new_order_id = result['SavedOrderId']
                        cur.execute("update _ExtPayOrder set [status]=1,neworderid=%s where orderid='%s'" % (
                            new_order_id, orderId))
                        conn.commit()
                time.sleep(0.2)
        cur.close()
        conn.close()
    except Exception as ex:
        print('exception', ex)


if __name__ == '__main__':
    config_file = os.getcwd() + "\\data.agent.json"
    print(config_file)
    config_json = open(config_file, encoding='utf-8')
    agent_config = json.load(config_json)
    import_api = import_api.replace("{ip}", agent_config['import'])
    conn_str = r'DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s;charset="utf8"' % (
        agent_config['server'], agent_config['database'], agent_config['uid'], agent_config['pwd'])
    print('...work...', agent_config['import'])
    while True:
        query_save()
        time.sleep(1)
