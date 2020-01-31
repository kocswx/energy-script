# coding:utf-8
from config.config_db_server import *
from stat_report_param import *
from stat_report_comm_data import *
from stat_report_comm_xls import *
from openpyxl import Workbook
import sys

db = DbConfig()

# 集团ID 站点ID "" "" 班次号 保存文件名
if __name__ == '__main__':
    # argv = ["stat_report_server_shift.py", 10003, 100031001, '', '', '201912291141',
    #         'D:\\eng\\client\\reports\\201912\\201912252036.xlsx']
    argv = sys.argv
    try:
        param = StatParam()
        param.init(argv, conn, db.DB_ERP, db.DB_ORDER)

        save_path = param.file_name
        print("save_path: ", save_path)
        prepare_path(save_path)
        cur = conn.cursor()

        shift_sql = "select SHIFT_NO,SHIFT_DATE,START_TIME,END_TIME,EMP_NAME FROM %s.shift_record " \
                    "where GROUP_ID='%s' AND STATION_ID='%s' AND SHIFT_NO='%s'" % (
                        db.DB_ORDER, param.group_id, param.station_id, param.shift_no)
        cur.execute(shift_sql)
        shift_row = cur.fetchone()
        param.shift_date = shift_row[1]
        param.shift_start_time = shift_row[2]
        param.shift_end_time = shift_row[3]
        # 生成汇总数据
        shift_emp_prod_pay(cur, db, param)
        shift_prod_noz_order(cur, db, param)
        shift_charge_order(cur, db, param)
        conn.commit()
        wb = Workbook()
        # 生成报表
        prod_noz_order_xls(wb, cur, db, param)
        emp_prod_pay_xls(wb, cur, db, param)
        charge_order_xls(wb, cur, db, param)
        wb.save(save_path)
    except Exception as e:
        print("Unexpected error: ", e)
    conn.close()
