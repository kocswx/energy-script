from config.config_db_client import conn
from _old.shift_report_comm import *
from _util import prepare_path
import sys

'''
客户端班次报表
站点号 班次号 班次日期 报表路径
'''
# python shift_report_client.py 100031001 201911212217 2019-11-21 D:\\rpt.xlsx

report_db_name = conn.database

if __name__ == '__main__':
    if len(sys.argv) < 5:
        print('param error', sys.argv)
    else:
        try:
            file_name = sys.argv[4]
            prepare_path(file_name)
            wb = build_shift_report(conn, sys.argv, report_db_name)
            # wb.save("D:\\shift_report_%s.xlsx" % (shift_no))
            wb.save(file_name)
        except Exception as e:
            print("Unexpected error: {e}")
    conn.close()
