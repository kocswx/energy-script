# coding:utf-8
from config.config_db_server import *
from stat_report_param import *
from stat_report_comm_xls import *
from openpyxl import Workbook

db = DbConfig()

# 集团ID [站点ID] 开始日期 结束日期 "" 保存文件名
if __name__ == '__main__':
    argv = ["stat_report_server_day.py", 10003, 100031001, '2019-01-01', '2020-01-28', '',
            '/10003/201912/uuid.xlsx']
    try:
        param = StatParam()
        param.init(argv, conn, db.DB_ERP, db.DB_ORDER)
        save_path = save_report_folder + param.file_name
        print("save_path: ", save_path)
        prepare_path(save_path)
        cur = conn.cursor()
        wb = Workbook()
        # 生成报表
        prod_noz_order_xls(wb, cur, db, param)
        emp_prod_pay_xls(wb, cur, db, param)
        charge_order_xls(wb, cur, db, param)
        wb.save(save_path)
    except Exception as e:
        print("Unexpected error:", e)
    conn.close()
