import pymysql as mdb

_db = {'host': '47.112.101.21', "port": 3306, "user": 'root', "password": 'sblv', "database": "eng_client"}

conn = mdb.connect(
    host=_db.get('host'),
    port=_db.get('port'),
    user=_db.get('user'),
    password=_db.get('password'),
    database=_db.get("database"),
    charset="utf8")

if __name__ == '__main__':
    db_cur = conn.cursor()
    db_cur.execute("select * from emp_info")
    results = db_cur.fetchall()
    if results:
        print(results)
    conn.close()
