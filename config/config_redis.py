# pip install redis
# import redis
#
# _pool = redis.ConnectionPool(host='47.112.101.21', port=6379, password='eng@redis', db=1)
# rd = redis.Redis(connection_pool=_pool)
# # redisCur.connection_pool.disconnect()
#
# if __name__ == '__main__':
#     key = "__x__"
#     rd.set(key, "1232132132")
#     print(rd.get(key))
#     rd.connection_pool.disconnect()
