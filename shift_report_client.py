from db_config_client import conn
from shift_report_comm import *
import sys

'''
客户端班次报表
站点号 班次号 班次日期
'''
# python shift_report_client.py 100031001 201911212217 2019-11-21 D:\\rpt.xlsx

if __name__ == '__main__':
    if len(sys.argv) == 5:
        build_shift_report(conn, sys.argv)
    else:
        print('param error', sys.argv)
