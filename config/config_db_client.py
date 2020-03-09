import pymysql as mdb

db = {'host': '127.0.0.1', "port": 3306, "user": 'root', "password": 'eng888@local', "database": "eng_client", "charset": "utf8"}
# db = {'host': '47.112.101.21',
#        "port": 3306,
#        "user": 'root',
#        "password": 'sblv',
#        "database": "eng_client",
#        "charset": "utf8"}

conn = mdb.connect(
    host=db.get('host'),
    port=db.get('port'),
    user=db.get('user'),
    password=db.get('password'),
    database=db.get("database"),
    charset="utf8")

if __name__ == '__main__':
    db_cur = conn.cursor()
    print("数据库连接正常")
    db_cur.close()
    conn.close()
