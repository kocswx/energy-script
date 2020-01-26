from config.config_db_server import conn
from shift_report_comm import *
from _util import prepare_path
import sys

'''
服务端班次报表
站点号 班次号 班次日期 报表路径
'''


# python shift_report_server.py 100031001 201911212217 2019-11-21 /opt/report/....

# 充值汇总
def _shift_charge_order(argv, rpt_db_name, wb):
    station_id = argv[1]
    shift_no = argv[2]
    cur = conn.cursor()
    fetch_station = "SELECT GROUP_ID,STATION_ID,STATION_NAME FROM eng_erp.station_info where STATION_ID=%s"
    cur.execute(fetch_station % (station_id))
    ctx = cur.fetchone()
    print(ctx)
    shift_sql = "select SHIFT_NO,SHIFT_DATE,START_TIME,END_TIME,EMP_NAME FROM shift_record " \
                "where GROUP_ID='%s' AND STATION_ID='%s' AND SHIFT_NO='%s'" % (ctx[0], ctx[1], shift_no)
    print(shift_sql)
    cur.execute(shift_sql)
    shift_row = cur.fetchone()
    print(shift_row)
    if shift_row:
        print('_shift_charge_order:',
              cur.execute("DELETE FROM %s.`rpt_shift_emp_charge` WHERE STATION_ID='%s' AND SHIFT_NO='%s'" % (
                  rpt_db_name, ctx[1], shift_no)))
        sql = "INSERT INTO %s.`rpt_shift_emp_charge`(`GROUP_ID`,`STATION_ID`,`SHIFT_NO`,`SHIFT_DATE`,`EMP_ID`,`EMP_NAME`,`PAY_TYPE_ID`,`PAY_TYPE_NAME`,`CNT`,`AMT`,`GIFT_AMT`,`TOTAL_AMT`)" \
              "SELECT GROUP_ID,OPT_STATION_ID as STATION_ID,'%s' as SHIFT_NO,'%s' as SHIFT_DATE, OPT_EMP_ID as EMP_ID,OPT_EMP_NAME as EMP_NAME,PAY_TYPE_ID,PAY_TYPE_NAME,COUNT(1)AS CNT,SUM(AMT)AS AMT,SUM(GIFT_AMT)AS GIFT_AMT,SUM(AMT+GIFT_AMT)as TOTAL_AMT " \
              "from eng_crm.history_charge WHERE GROUP_ID='%s' AND OPT_STATION_ID='%s' AND CREATED_TIME>='%s' AND CREATED_TIME<'%s' " \
              "GROUP BY GROUP_ID,OPT_STATION_ID,OPT_EMP_ID,OPT_EMP_NAME,PAY_TYPE_ID,PAY_TYPE_NAME ORDER BY PAY_TYPE_ID" % (
                  rpt_db_name, shift_row[0], shift_row[1], ctx[0], ctx[1], shift_row[2], shift_row[3])
        # print(sql)
        cur.execute(sql)
        conn.commit()

        # 生成报表
        sheet = wb.create_sheet()
        sheet.title = '班次充值汇总'
        cell_name = ['员工姓名', '支付方式', '笔数', '充值金额', '赠送金额', '实际充入']
        row_index = 1
        sheet.merge_cells("A1:%s1" % column_letter[len(cell_name) - 1])
        sheet.cell(row_index, 1).value = sheet.title
        sheet.cell(row_index, 1).font = font_rpt_title
        sheet.cell(row_index, 1).alignment = align_center
        sheet.cell(row_index, 1).border = border
        row_index += 1

        sheet.cell(row_index, 1, '站点名称：%s' % ctx[2])
        sheet.cell(row_index, 1).alignment = align_left
        sheet.cell(row_index, 1).border = border
        sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=4)
        sheet.cell(row_index, 5, '交班员工：%s' % shift_row[4])
        sheet.cell(row_index, 5).alignment = align_right
        sheet.cell(row_index, 5).border = border
        sheet.merge_cells(start_row=row_index, start_column=5, end_row=row_index, end_column=len(cell_name))
        row_index += 1

        sheet.cell(row_index, 1, '班次号：%s' % shift_row[0])
        sheet.cell(row_index, 1).alignment = align_left
        sheet.cell(row_index, 1).border = border
        sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=2)
        sheet.cell(row_index, 3, '班次日期：%s' % shift_row[1])
        sheet.cell(row_index, 3).alignment = align_center
        sheet.cell(row_index, 3).border = border
        sheet.merge_cells(start_row=row_index, start_column=3, end_row=row_index, end_column=4)
        sheet.cell(row_index, 5, '打印时间：%s' % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        sheet.cell(row_index, 5).alignment = align_right
        sheet.cell(row_index, 5).border = border
        sheet.merge_cells(start_row=row_index, start_column=5, end_row=row_index, end_column=7)
        row_index += 1

        for i in range(len(cell_name)):
            sheet.cell(row_index, i + 1).font = font_rpt_cell_header
            sheet.cell(row_index, i + 1).alignment = align_center
            sheet.cell(row_index, i + 1).border = border
            sheet.cell(row_index, i + 1, cell_name[i])
            sheet.column_dimensions[column_letter[i]].width = 12
        row_index += 1

        total = Total()  # 总计
        emp_total = Total()  # 行汇总
        rpt_sql = "SELECT EMP_NAME,PAY_TYPE_NAME,SUM(CNT),SUM(AMT),SUM(GIFT_AMT),SUM(TOTAL_AMT) FROM %s.`rpt_shift_emp_charge` " \
                  "where station_id='%s' and shift_no='%s' GROUP BY EMP_NAME,PAY_TYPE_NAME" % (
                      rpt_db_name, station_id, shift_no)
        print(rpt_sql)
        cur.execute(rpt_sql)
        rows = cur.fetchall()
        if rows:
            for data_index in range(len(rows)):
                row = rows[data_index]
                prod = row[0]
                for i in range(len(cell_name)):
                    sheet.cell(row_index, i + 1).font = font_rpt_normal
                    sheet.cell(row_index, i + 1).alignment = align_center
                    sheet.cell(row_index, i + 1).border = border
                    sheet.cell(row_index, i + 1, row[i])
                row_index += 1
                # 总累计
                total.cnt += row[2]
                total.chargeAmt += row[3]
                total.giftAmt += row[4]
                total.totalAmt += row[5]
                # 行累计
                emp_total.cnt += row[2]
                emp_total.chargeAmt += row[3]
                emp_total.giftAmt += row[4]
                emp_total.totalAmt += row[5]
                emp_total.rowCnt += 1
                if data_index > 0 and (data_index + 1 == len(rows) or prod != rows[data_index + 1][0]):
                    if emp_total.rowCnt > 1:
                        merge_start_row = row_index - emp_total.rowCnt
                        sheet.merge_cells(start_row=merge_start_row, start_column=1, end_row=row_index - 1,
                                          end_column=1)
                    # 写油品小计
                    sheet.cell(row_index, 1).font = font_rpt_normal
                    sheet.cell(row_index, 1).alignment = align_center
                    sheet.cell(row_index, 1).border = border
                    sheet.cell(row_index, 1, '小计：')
                    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=2)
                    sheet.cell(row_index, 3).font = font_rpt_normal
                    sheet.cell(row_index, 3).alignment = align_center
                    sheet.cell(row_index, 3).border = border
                    sheet.cell(row_index, 3, emp_total.cnt)
                    sheet.cell(row_index, 4).font = font_rpt_normal
                    sheet.cell(row_index, 4).alignment = align_center
                    sheet.cell(row_index, 4).border = border
                    sheet.cell(row_index, 4, emp_total.chargeAmt)
                    sheet.cell(row_index, 5).font = font_rpt_normal
                    sheet.cell(row_index, 5).alignment = align_center
                    sheet.cell(row_index, 5).border = border
                    sheet.cell(row_index, 5, emp_total.giftAmt)
                    sheet.cell(row_index, 6).font = font_rpt_normal
                    sheet.cell(row_index, 6).alignment = align_center
                    sheet.cell(row_index, 6).border = border
                    sheet.cell(row_index, 6, emp_total.totalAmt)
                    emp_total = Total()
                    row_index += 1
        # 总计
        sheet.cell(row_index, 1).font = font_rpt_normal
        sheet.cell(row_index, 1).alignment = align_center
        sheet.cell(row_index, 1).border = border
        sheet.cell(row_index, 1, '总计：')
        sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=2)
        sheet.cell(row_index, 3).font = font_rpt_normal
        sheet.cell(row_index, 3).alignment = align_center
        sheet.cell(row_index, 3).border = border
        sheet.cell(row_index, 3, total.cnt)
        sheet.cell(row_index, 4).font = font_rpt_normal
        sheet.cell(row_index, 4).alignment = align_center
        sheet.cell(row_index, 4).border = border
        sheet.cell(row_index, 4, total.chargeAmt)
        sheet.cell(row_index, 5).font = font_rpt_normal
        sheet.cell(row_index, 5).alignment = align_center
        sheet.cell(row_index, 5).border = border
        sheet.cell(row_index, 5, total.giftAmt)
        sheet.cell(row_index, 6).font = font_rpt_normal
        sheet.cell(row_index, 6).alignment = align_center
        sheet.cell(row_index, 6).border = border
        sheet.cell(row_index, 6, total.totalAmt)
        row_index += 1
        # 统一行高
        for row_index in range(row_index):
            sheet.row_dimensions[row_index].height = 20
        sheet.protection.sheet = True
        sheet.protection.password = '123456'
        sheet.protection.enable()

        wb.security.workbookPassword = '123456'
        wb.security.lockStructure = True
        cur.close()
        return wb


report_db_name = 'eng_report'

if __name__ == '__main__':
    if len(sys.argv) < 5:
        print('param error', sys.argv)
    else:
        try:
            file_name = sys.argv[4]
            prepare_path(file_name)
            wb = build_shift_report(conn, sys.argv, report_db_name)
            _shift_charge_order(sys.argv, report_db_name, wb)
            # wb.save("D:\\shift_report_%s.xlsx" % (shift_no))
            wb.save(file_name)
        except Exception as e:
            print(f"Unexpected error: {e}")
    conn.close()
