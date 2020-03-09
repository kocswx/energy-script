# coding:utf-8
from _openpyxl_constant import *
from _util import *


class Total:
    cnt = 0
    vol = 0
    receAmt = 0
    realAmt = 0
    discAmt = 0

    chargeAmt = 0
    giftAmt = 0
    totalAmt = 0

    rowCnt = 0


def _get_station_name(cur, db, station_id):
    sql = "SELECT STATION_ID,STATION_NAME FROM %s.station_info where STATION_ID='%s'" % (db.DB_ERP, station_id)
    cur.execute(sql)
    station_name = cur.fetchone()[1]
    return station_name


def _get_report_desc(param, cur, db):
    param_string = ""
    if param.station_id:
        sql = "select STATION_NAME from %s.station_info where STATION_ID='%s'" % (db.DB_ERP, param.station_id)
        # print(sql)
        cur.execute(sql)
        param.station_name = cur.fetchone()[0]
        print("站点：%s  " % param.station_name)
    if param.begin_date:
        param_string = param_string + "开始：%s  " % param.begin_date
    if param.end_date:
        param_string = param_string + "结束：%s  " % param.end_date
    if param.shift_no:
        cur.execute("select SHIFT_DATE,START_TIME,END_TIME,EMP_NAME from %s.shift_record "
                    "where STATION_ID='%s' AND SHIFT_NO='%s'" % (db.DB_ORDER, param.station_id, param.shift_no))
        print()
        shift = cur.fetchone()
        param.shift_date = shift[0]
        param.shift_start_time = shift[1]
        param.shift_end_time = shift[2]
        param.shift_emp_name = shift[3]
        param_string = param_string + "班次：%s 开始：%s 结束：%s 交班员工：%s" \
                       % (param.shift_no, param.shift_start_time, param.shift_end_time, param.shift_emp_name)
    print(param_string)
    return param_string


# 油品-油枪销售汇总
def prod_noz_order_xls(wb, cur, db, param_xml):
    print("prod_noz_order_xls")
    param = param_xml.copy()
    sheet = wb.active
    sheet.title = '油品-油枪销售汇总'
    cell_name = ['站点名称', '油品名称', '枪号', '笔数', '销售升数', '销售金额', '实收金额', '优惠金额']
    row_index = 1
    sheet.merge_cells("A1:%s1" % column_letter[len(cell_name) - 1])
    set_cell(sheet, row_index, 1, sheet.title, font_rpt_title, align_center, border)
    row_index += 1

    set_cell(sheet, row_index, 1, ('打印时间：%s' % get_now_full()), font_rpt_normal, align_right, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=len(cell_name))
    row_index += 1

    set_cell(sheet, row_index, 1, _get_report_desc(param, cur, db), font_rpt_normal, align_left, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=len(cell_name))
    row_index += 1

    for i in range(len(cell_name)):
        set_cell(sheet, row_index, i + 1, cell_name[i], font_rpt_cell_header, align_center, border)
        sheet.column_dimensions[column_letter[i]].width = 12
    row_index += 1

    total = Total()  # 总计

    if param.station_id:
        row_index = _prod_noz_order_xls_station(sheet, row_index, cur, db, param, total)
    else:
        cur.execute("SELECT STATION_ID,STATION_NAME FROM %s.station_info where GROUP_ID='%s'"
                    % (db.DB_ERP, param.group_id))
        rows = cur.fetchall()
        for row in rows:
            param.station_id = row[0]
            total_station = Total()
            row_index = _prod_noz_order_xls_station(sheet, row_index, cur, db, param, total_station)
            total.cnt += total_station.cnt
            total.vol += total_station.vol
            total.receAmt += total_station.receAmt
            total.realAmt += total_station.realAmt
            total.discAmt += total_station.discAmt
    # 总计
    set_cell(sheet, row_index, 1, '总计：', font_rpt_normal, align_center, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=3)
    set_cell(sheet, row_index, 4, total.cnt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 5, total.vol, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 6, total.receAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 7, total.realAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 8, total.discAmt, font_rpt_normal, align_center, border)
    row_index += 1
    # 统一行高
    for row_index in range(row_index):
        sheet.row_dimensions[row_index].height = 20
    # sheet.protection.sheet = True
    # sheet.protection.password = '123456'
    # sheet.protection.enable()


def _prod_noz_order_xls_station(sheet, row_index, cur, db, param, total_station):
    _merge_start = row_index
    station_name = _get_station_name(cur, db, param.station_id)
    set_cell(sheet, row_index, 1, station_name, font_rpt_normal, align_center, border)

    rpt_sql = "SELECT PROD_NAME,NOZ_NO,SUM(CNT),SUM(VOL),SUM(RECE_AMT),SUM(REAL_AMT),SUM(DISC_AMT),STATION_ID,PROD_ID " \
              "FROM %s.rpt_shift_prod_noz_order " \
              "WHERE GROUP_ID=%s AND STATION_ID='%s' " % (db.DB_REPORT, param.group_id, param.station_id)
    if param.shift_no:
        rpt_sql += (" AND SHIFT_NO='%s'" % param.shift_no)
    else:
        rpt_sql += (" AND SHIFT_DATE>='%s' AND SHIFT_DATE<='%s'" % (param.begin_date, param.end_date))
    rpt_sql += "GROUP BY STATION_ID,PROD_NAME,PROD_ID,NOZ_NO ORDER BY STATION_ID,PROD_ID,NOZ_NO"
    print(rpt_sql)

    cur.execute(rpt_sql)
    rows = cur.fetchall()
    temp_total = Total()
    if rows:
        for data_index in range(len(rows)):
            row = rows[data_index]
            prod = row[0]
            noz = row[1]
            cnt = row[2]
            vol = row[3]
            rece_amt = row[4]
            real_amt = row[5]
            disc_amt = row[6]
            set_cell(sheet, row_index, 1, station_name, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 2, prod, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 3, noz, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 4, cnt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 5, vol, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 6, rece_amt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 7, real_amt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 8, disc_amt, font_rpt_normal, align_center, border)
            row_index += 1
            # 行累计
            temp_total.rowCnt += 1
            temp_total.cnt += cnt
            temp_total.vol += vol
            temp_total.receAmt += rece_amt
            temp_total.realAmt += real_amt
            temp_total.discAmt += disc_amt
            total_station.cnt += cnt
            total_station.vol += vol
            total_station.receAmt += rece_amt
            total_station.realAmt += real_amt
            total_station.discAmt += disc_amt

            if data_index > 0 and (data_index + 1 == len(rows) or prod != rows[data_index + 1][0]):
                if temp_total.rowCnt > 1:
                    merge_start_row = row_index - temp_total.rowCnt
                    sheet.merge_cells(start_row=merge_start_row, start_column=2, end_row=row_index - 1, end_column=2)
                # 写油品小计
                set_cell(sheet, row_index, 2, '小计：', font_rpt_normal, align_center, border)
                sheet.merge_cells(start_row=row_index, start_column=2, end_row=row_index, end_column=3)
                set_cell(sheet, row_index, 4, temp_total.cnt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 5, temp_total.vol, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 6, temp_total.receAmt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 7, temp_total.realAmt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 8, temp_total.discAmt, font_rpt_normal, align_center, border)
                temp_total = Total()
                row_index += 1
    # 合计
    set_cell(sheet, row_index, 2, '合计：', font_rpt_normal, align_center, border)
    sheet.merge_cells(start_row=row_index, start_column=2, end_row=row_index, end_column=3)
    set_cell(sheet, row_index, 4, total_station.cnt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 5, total_station.vol, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 6, total_station.receAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 7, total_station.realAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 8, total_station.discAmt, font_rpt_normal, align_center, border)
    # 合并站点
    sheet.merge_cells(start_row=_merge_start, start_column=1, end_row=row_index, end_column=1)
    row_index += 1
    return row_index


# 员工-油品-支付方式汇总 非余额支付
def emp_prod_pay_xls(wb, cur, db, param_xml):
    print("emp_prod_pay_xls")
    param = param_xml.copy()
    sheet = wb.create_sheet()
    sheet.title = '员工加油-收款方式汇总'
    cell_name = ['站点名称', '加油员', '收款方式', '笔数', '销售升数', '销售金额', '实收金额', '优惠金额']
    row_index = 1

    sheet.merge_cells("A1:%s1" % column_letter[len(cell_name) - 1])
    set_cell(sheet, row_index, 1, sheet.title, font_rpt_title, align_center, border)
    row_index += 1

    set_cell(sheet, row_index, 1, ('打印时间：%s' % get_now_full()), font_rpt_normal, align_right, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=len(cell_name))
    row_index += 1

    set_cell(sheet, row_index, 1, _get_report_desc(param, cur, db), font_rpt_normal, align_left, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=len(cell_name))
    row_index += 1

    for i in range(len(cell_name)):
        set_cell(sheet, row_index, i + 1, cell_name[i], font_rpt_cell_header, align_center, border)
        sheet.column_dimensions[column_letter[i]].width = 12
    row_index += 1
    total = Total()  # 总计
    if param.station_id:
        row_index = _emp_prod_pay_xls_station(sheet, row_index, cur, db, param, total)
    else:
        cur.execute("SELECT STATION_ID,STATION_NAME FROM %s.station_info where GROUP_ID='%s'"
                    % (db.DB_ERP, param.group_id))
        rows = cur.fetchall()
        for row in rows:
            param.station_id = row[0]
            total_station = Total()
            row_index = _emp_prod_pay_xls_station(sheet, row_index, cur, db, param, total_station)
            total.cnt += total_station.cnt
            total.vol += total_station.vol
            total.receAmt += total_station.receAmt
            total.realAmt += total_station.realAmt
            total.discAmt += total_station.discAmt
    # 总计
    set_cell(sheet, row_index, 1, '总计：', font_rpt_normal, align_center, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=3)
    set_cell(sheet, row_index, 4, total.cnt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 5, total.vol, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 6, total.receAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 7, total.realAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 8, total.discAmt, font_rpt_normal, align_center, border)
    row_index += 1
    # 统一行高
    for row_index in range(row_index):
        sheet.row_dimensions[row_index].height = 20


def _emp_prod_pay_xls_station(sheet, row_index, cur, db, param, total_station):
    _merge_start = row_index
    station_name = _get_station_name(cur, db, param.station_id)
    set_cell(sheet, row_index, 1, station_name, font_rpt_normal, align_center, border)

    rpt_sql = "SELECT EMP_NAME,PAY_TYPE_NAME,SUM(CNT),SUM(VOL),SUM(RECE_AMT),SUM(REAL_AMT),SUM(DISC_AMT) " \
              "FROM %s.rpt_shift_emp_prod_pay where GROUP_ID='%s'and STATION_ID='%s' " \
              % (db.DB_REPORT, param.group_id, param.station_id)
    if param.shift_no:
        rpt_sql += (" AND SHIFT_NO='%s'" % param.shift_no)
    else:
        rpt_sql += (" AND SHIFT_DATE>='%s' AND SHIFT_DATE<='%s'" % (param.begin_date, param.end_date))
    rpt_sql += "GROUP BY EMP_NAME,PAY_TYPE_NAME"
    print(rpt_sql)

    cur.execute(rpt_sql)
    rows = cur.fetchall()
    temp_total = Total()
    if rows:
        for data_index in range(len(rows)):
            row = rows[data_index]
            emp_name = row[0]
            noz = row[1]
            cnt = row[2]
            vol = row[3]
            rece_amt = row[4]
            real_amt = row[5]
            disc_amt = row[6]
            set_cell(sheet, row_index, 1, station_name, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 2, emp_name, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 3, noz, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 4, cnt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 5, vol, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 6, rece_amt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 7, real_amt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 8, disc_amt, font_rpt_normal, align_center, border)
            row_index += 1
            # 行累计
            temp_total.rowCnt += 1
            temp_total.cnt += cnt
            temp_total.vol += vol
            temp_total.receAmt += rece_amt
            temp_total.realAmt += real_amt
            temp_total.discAmt += disc_amt
            total_station.cnt += cnt
            total_station.vol += vol
            total_station.receAmt += rece_amt
            total_station.realAmt += real_amt
            total_station.discAmt += disc_amt

            if data_index > 0 and (data_index + 1 == len(rows) or emp_name != rows[data_index + 1][0]):
                if temp_total.rowCnt > 1:
                    merge_start_row = row_index - temp_total.rowCnt
                    sheet.merge_cells(start_row=merge_start_row, start_column=2, end_row=row_index - 1, end_column=2)
                # 写小计
                set_cell(sheet, row_index, 2, '小计：', font_rpt_normal, align_center, border)
                sheet.merge_cells(start_row=row_index, start_column=2, end_row=row_index, end_column=3)
                set_cell(sheet, row_index, 4, temp_total.cnt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 5, temp_total.vol, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 6, temp_total.receAmt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 7, temp_total.realAmt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 8, temp_total.discAmt, font_rpt_normal, align_center, border)
                temp_total = Total()
                row_index += 1
    # 合计
    set_cell(sheet, row_index, 2, '合计：', font_rpt_normal, align_center, border)
    sheet.merge_cells(start_row=row_index, start_column=2, end_row=row_index, end_column=3)
    set_cell(sheet, row_index, 4, total_station.cnt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 5, total_station.vol, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 6, total_station.receAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 7, total_station.realAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 8, total_station.discAmt, font_rpt_normal, align_center, border)
    # 合并站点
    sheet.merge_cells(start_row=_merge_start, start_column=1, end_row=row_index, end_column=1)
    row_index += 1
    return row_index


# 充值汇总
def charge_order_xls(wb, cur, db, param_xml):
    print("charge_order_xls")
    param = param_xml.copy()
    sheet = wb.create_sheet()
    sheet.title = '充值汇总'
    cell_name = ['站点名称', '员工姓名', '收款方式', '笔数', '充值金额', '赠送金额', '实际充入']
    row_index = 1
    sheet.merge_cells("A1:%s1" % column_letter[len(cell_name) - 1])
    set_cell(sheet, row_index, 1, sheet.title, font_rpt_title, align_center, border)
    row_index += 1

    set_cell(sheet, row_index, 1, ('打印时间：%s' % get_now_full()), font_rpt_normal, align_right, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=len(cell_name))
    row_index += 1

    set_cell(sheet, row_index, 1, _get_report_desc(param, cur, db), font_rpt_normal, align_left, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=len(cell_name))
    row_index += 1

    for i in range(len(cell_name)):
        set_cell(sheet, row_index, i + 1, cell_name[i], font_rpt_cell_header, align_center, border)
        sheet.column_dimensions[column_letter[i]].width = 12
    row_index += 1
    total = Total()  # 总计
    if param.station_id:
        row_index = _charge_order_xls_station(sheet, row_index, cur, db, param, total)
    else:
        cur.execute("SELECT STATION_ID,STATION_NAME FROM %s.station_info where GROUP_ID='%s'"
                    % (db.DB_ERP, param.group_id))
        rows = cur.fetchall()
        for row in rows:
            param.station_id = row[0]
            total_station = Total()
            row_index = _charge_order_xls_station(sheet, row_index, cur, db, param, total_station)
            total.cnt += total_station.cnt
            total.chargeAmt += total_station.chargeAmt
            total.giftAmt += total_station.giftAmt
            total.totalAmt += total_station.totalAmt
    # 总计
    set_cell(sheet, row_index, 1, '总计：', font_rpt_normal, align_center, border)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=3)
    set_cell(sheet, row_index, 4, total.cnt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 5, total.chargeAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 6, total.giftAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 7, total.totalAmt, font_rpt_normal, align_center, border)
    row_index += 1
    # 统一行高
    for row_index in range(row_index):
        sheet.row_dimensions[row_index].height = 20


def _charge_order_xls_station(sheet, row_index, cur, db, param, total_station):
    _merge_start = row_index
    station_name = _get_station_name(cur, db, param.station_id)
    set_cell(sheet, row_index, 1, station_name, font_rpt_normal, align_center, border)

    rpt_sql = "SELECT EMP_NAME,PAY_TYPE_NAME,SUM(CNT),SUM(AMT),SUM(GIFT_AMT),SUM(TOTAL_AMT) " \
              "FROM %s.`rpt_shift_emp_charge` where GROUP_ID='%s' and STATION_ID='%s' " \
              % (db.DB_REPORT, param.group_id, param.station_id)
    if param.shift_no:
        rpt_sql += (" AND SHIFT_NO='%s'" % param.shift_no)
    else:
        rpt_sql += (" AND SHIFT_DATE>='%s' AND SHIFT_DATE<='%s'" % (param.begin_date, param.end_date))
    rpt_sql += "GROUP BY EMP_NAME,PAY_TYPE_NAME"
    print(rpt_sql)
    cur.execute(rpt_sql)
    rows = cur.fetchall()
    temp_total = Total()
    if rows:
        for data_index in range(len(rows)):
            row = rows[data_index]
            emp_name = row[0]
            pay_type_name = row[1]
            cnt = row[2]
            charge_amt = row[3]
            gift_amt = row[4]
            total_amt = row[5]
            set_cell(sheet, row_index, 1, station_name, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 2, emp_name, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 3, pay_type_name, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 4, cnt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 5, charge_amt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 6, gift_amt, font_rpt_normal, align_center, border)
            set_cell(sheet, row_index, 7, total_amt, font_rpt_normal, align_center, border)
            row_index += 1
            # 总累计
            temp_total.rowCnt += 1
            temp_total.cnt += cnt
            temp_total.chargeAmt += charge_amt
            temp_total.giftAmt += gift_amt
            temp_total.totalAmt += total_amt
            total_station.cnt += cnt
            total_station.chargeAmt += charge_amt
            total_station.giftAmt += gift_amt
            total_station.totalAmt += total_amt

            if data_index > 0 and (data_index + 1 == len(rows) or emp_name != rows[data_index + 1][0]):
                if temp_total.rowCnt > 1:
                    merge_start_row = row_index - temp_total.rowCnt
                    sheet.merge_cells(start_row=merge_start_row, start_column=2, end_row=row_index - 1, end_column=2)
                # 写小计
                set_cell(sheet, row_index, 2, '小计：', font_rpt_normal, align_center, border)
                sheet.merge_cells(start_row=row_index, start_column=2, end_row=row_index, end_column=3)
                set_cell(sheet, row_index, 4, temp_total.cnt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 5, temp_total.chargeAmt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 6, temp_total.giftAmt, font_rpt_normal, align_center, border)
                set_cell(sheet, row_index, 7, temp_total.totalAmt, font_rpt_normal, align_center, border)
                temp_total = Total()
                row_index += 1
    # 合计
    set_cell(sheet, row_index, 2, '合计：', font_rpt_normal, align_center, border)
    sheet.merge_cells(start_row=row_index, start_column=2, end_row=row_index, end_column=3)
    set_cell(sheet, row_index, 4, total_station.cnt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 5, total_station.chargeAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 6, total_station.giftAmt, font_rpt_normal, align_center, border)
    set_cell(sheet, row_index, 7, total_station.totalAmt, font_rpt_normal, align_center, border)
    # 合并站点
    sheet.merge_cells(start_row=_merge_start, start_column=1, end_row=row_index, end_column=1)
    row_index += 1
    return row_index
