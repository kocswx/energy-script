import sys
import datetime
from config.config_db_server import conn

cur = conn.cursor()


def daily_fuel_summary(day, row):
    """
    每日销量统计
    :param day:
    :return:
    """
    SQL_CLEAN_BY_DAY = "DELETE FROM eng_report.daily_fuel_summary WHERE GROUP_ID='%s' AND STATION_ID='%s' AND STAT_DATE='%s'"
    SQL_STAT_DAY = "INSERT INTO eng_report.daily_fuel_summary(GROUP_ID,STATION_ID,STAT_DATE,NOZ_NO,CNT,VOL,RECE_AMT,REAL_AMT,DISC_AMT,PAY_TYPE_ID,PAY_TYPE_NAME,EMP_ID,EMP_NAME,PROD_ID,PROD_NAME)" \
                   "SELECT GROUP_ID,STATION_ID,'%s' AS STAT_DATE,NOZ_NO,COUNT(1) AS CNT,SUM(VOL) AS VOL,SUM(RECE_AMT)AS RECE_AMT,SUM(REAL_AMT)AS REAL_AMT,SUM(DISC_AMT)AS DISC_AMT,PAY_TYPE_ID,PAY_TYPE_NAME,EMP_ID,EMP_NAME,PROD_ID,PROD_NAME " \
                   "FROM eng_order.fuel_order WHERE GROUP_ID='%s' AND STATION_ID='%s' and CREATED_TIME>='%s' AND CREATED_TIME<='%s'" \
                   "GROUP BY GROUP_ID,STATION_ID,NOZ_NO,PAY_TYPE_ID,PAY_TYPE_NAME,EMP_ID,EMP_NAME,PROD_ID,PROD_NAME"
    print(row[1], 'fuel_summary')
    day_start = day + " 00:00:00"
    day_end = day + " 23:59:59"
    cur.execute(SQL_CLEAN_BY_DAY % (row[0], row[1], day))
    sql = SQL_STAT_DAY % (day, row[0], row[1], day_start, day_end)
    cur.execute(sql)


def daily_charge_summary(day, row):
    """
    每日充值统计
    :param day:
    :return:
    """
    SQL_CLEAN_BY_DAY = "DELETE FROM eng_report.daily_charge_summary WHERE GROUP_ID='%s' AND STATION_ID='%s' AND STAT_DATE='%s'"
    SQL_STAT_DAY = "INSERT INTO eng_report.daily_charge_summary(`GROUP_ID`,`STATION_ID`,`STAT_DATE`,`CNT`,`AMT`,`GIFT_AMT`,`PAY_TYPE_ID`,`PAY_TYPE_NAME`,`EMP_ID`,`EMP_NAME`)" \
                   "SELECT GROUP_ID,OPT_STATION_ID as STATION_ID,'%s' AS STAT_DATE,COUNT(1) AS CNT,SUM(AMT) AS AMT,SUM(GIFT_AMT)AS GIFT_AMT," \
                   "PAY_TYPE_ID,PAY_TYPE_NAME,OPT_EMP_ID AS EMP_ID,OPT_EMP_NAME AS EMP_NAME FROM eng_crm.history_charge " \
                   "WHERE GROUP_ID='%s' AND OPT_STATION_ID='%s' and CREATED_TIME>='%s' AND CREATED_TIME<='%s' " \
                   "GROUP BY GROUP_ID,STATION_ID,PAY_TYPE_ID,PAY_TYPE_NAME,OPT_STATION_ID,OPT_EMP_ID,OPT_EMP_NAME"

    day_start = day + " 00:00:00"
    day_end = day + " 23:59:59"
    print(row[1], 'charge_summary')
    cur.execute(SQL_CLEAN_BY_DAY % (row[0], row[1], day))
    sql = SQL_STAT_DAY % (day, row[0], row[1], day_start, day_end)
    cur.execute(sql)


SQL_ALL_STATION = "SELECT GROUP_ID,STATION_ID FROM eng_erp.station_info"

if __name__ == '__main__':
    print("每日消费充值汇总 ", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    day = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")
    if len(sys.argv) == 2:
        day = sys.argv[1]
    print('统计日期:', day)
    cur.execute(SQL_ALL_STATION)
    rows = cur.fetchall()
    if rows:
        for row in rows:
            daily_fuel_summary(day, row)
            daily_charge_summary(day, row)
            conn.commit()
    conn.close()
    print('统计完成 %s.' % day)
