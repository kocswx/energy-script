import pymysql as mdb

# _db = {'host': '127.0.0.1', "port": 3306, "user": 'root', "password": 'eng888@local', "database": "eng_client", "charset": "utf8"}
_db = {'host': '47.112.101.21',
       "port": 3306,
       "user": 'root',
       "password": 'sblv',
       "database": "eng_client",
       "charset": "utf8"}

conn = mdb.connect(
    host=_db.get('host'),
    port=_db.get('port'),
    user=_db.get('user'),
    password=_db.get('password'),
    database=_db.get("database"),
    charset="utf8")

if __name__ == '__main__':
    db_cur = conn.cursor()
    print("数据库连接正常")
    db_cur.close()
    conn.close()
