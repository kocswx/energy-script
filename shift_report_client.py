from db_config_client import conn
from shift_report_comm import *
import sys

'''
客户端班次报表
站点号 班次号 班次日期
'''
# python shift_report_client.py 100031001 201911212217 2019-11-21 D:\\rpt.xlsx

if __name__ == '__main__':
    if len(sys.argv) < 5:
        print('param error', sys.argv)
    else:
        try:
            file_name = sys.argv[4]
            prepare_path(file_name)
            wb = build_shift_report(conn, sys.argv)
            # wb.save("D:\\shift_report_%s.xlsx" % (shift_no))
            wb.save(file_name)
        except Exception as e:
            print(f"Unexpected error: {e}")
    conn.close()
