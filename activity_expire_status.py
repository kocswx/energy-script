from config.config_db_server import conn

'''
服务端检查活动状态 修改已经过期的活动状态
'''

if __name__ == '__main__':
    cur = conn.cursor()
    cur.execute("UPDATE eng_erp.activity SET STATUS=-1 WHERE STATUS>-1 AND NOW()>END_TIME;")
    conn.commit()
    effectRow = cur.rowcount
    print('服务端检查活动状态 影响行数:', effectRow)
    cur.close()
    conn.close()
