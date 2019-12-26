from db_config_server import conn
from shift_report_comm import *
import sys

'''
服务端班次报表
站点号 班次号 班次日期
'''
# python shift_report_server.py 100031001 201911212217 2019-11-21 /opt/report/....

if __name__ == '__main__':
    if len(sys.argv) == 5:
        build_shift_report(conn, sys.argv)
    else:
        print('param error', sys.argv)(conn, sys.argv)
