# coding:utf-8


# 油品-油枪销售汇总
def shift_prod_noz_order(cur, db, param):
    print('rpt_shift_prod_noz_order:',
          cur.execute("DELETE FROM %s.rpt_shift_prod_noz_order WHERE GROUP_ID='%s' and STATION_ID='%s' AND SHIFT_NO='%s'"
                      % (db.DB_REPORT, param.group_id, param.station_id, param.shift_no)))
    sql = "INSERT INTO %s.`rpt_shift_prod_noz_order`(`GROUP_ID`, `STATION_ID`, `SHIFT_NO`, `SHIFT_DATE`, `NOZ_NO`, `PROD_ID`,`PROD_NAME`, `CNT`, `VOL`, `RECE_AMT`, `REAL_AMT`, `DISC_AMT`)" \
          "SELECT GROUP_ID,STATION_ID,SHIFT_NO,'%s' AS SHIFT_DATE,NOZ_NO,PROD_ID,PROD_NAME,COUNT(1) AS CNT,ifnull(SUM(VOL), 0) AS VOL,ifnull(SUM(RECE_AMT), 0) AS RECE_AMT," \
          "ifnull(SUM(REAL_AMT), 0) AS REAL_AMT,ifnull(SUM(DISC_AMT), 0) AS DISC_AMT FROM %s.fuel_order WHERE GROUP_ID='%s' AND STATION_ID='%s' AND SHIFT_NO ='%s' GROUP BY GROUP_ID,STATION_ID,SHIFT_NO,NOZ_NO,PROD_ID,PROD_NAME" % (
              db.DB_REPORT, param.shift_date, db.DB_ORDER, param.group_id, param.station_id, param.shift_no)
    print(sql)
    cur.execute(sql)


# 员工-油品-支付方式汇总 非余额支付
def shift_emp_prod_pay(cur, db, param):
    print('shift_emp_prod_pay:',
          cur.execute("DELETE FROM %s.rpt_shift_emp_prod_pay WHERE GROUP_ID='%s' and STATION_ID='%s' AND SHIFT_NO='%s'"
                      % (db.DB_REPORT, param.group_id, param.station_id, param.shift_no)))
    sql = "INSERT INTO %s.`rpt_shift_emp_prod_pay` (`GROUP_ID`, `STATION_ID`, `SHIFT_NO`, `SHIFT_DATE`, `PROD_ID`, `PROD_NAME`, `EMP_ID`, `EMP_NAME`, `PAY_TYPE_ID`, `PAY_TYPE_NAME`, `CNT`, `VOL`, `RECE_AMT`, `REAL_AMT`, `DISC_AMT`) " \
          "SELECT GROUP_ID,STATION_ID,SHIFT_NO,'%s' as SHIFT_DATE,PROD_ID,PROD_NAME,EMP_ID,EMP_NAME,PAY_TYPE_ID,PAY_TYPE_NAME,COUNT(1) AS CNT,SUM(VOL) AS VOL,SUM(RECE_AMT) AS RECE_AMT, SUM(REAL_AMT) AS REAL_AMT,SUM(DISC_AMT) AS DISC_AMT " \
          "FROM eng_order.fuel_order WHERE GROUP_ID='%s' and STATION_ID='%s' AND PAY_TYPE_ID<>2 and SHIFT_NO='%s' GROUP BY EMP_ID,EMP_NAME,PROD_ID,PROD_NAME,PAY_TYPE_ID,PAY_TYPE_NAME, GROUP_ID,STATION_ID,SHIFT_NO" % (
              db.DB_REPORT, param.shift_date, param.group_id, param.station_id, param.shift_no)
    print(sql)
    cur.execute(sql)


# 充值汇总
def shift_charge_order(cur, db, param):
    print('_shift_charge_order:',
          cur.execute(
              "DELETE FROM %s.`rpt_shift_emp_charge` WHERE GROUP_ID='%s' and STATION_ID='%s' AND SHIFT_NO='%s'" % (
                  db.DB_REPORT, param.group_id, param.station_id, param.shift_no)))
    sql = "INSERT INTO %s.`rpt_shift_emp_charge`(`GROUP_ID`,`STATION_ID`,`SHIFT_NO`,`SHIFT_DATE`,`EMP_ID`,`EMP_NAME`,`PAY_TYPE_ID`,`PAY_TYPE_NAME`,`CNT`,`AMT`,`GIFT_AMT`,`TOTAL_AMT`)" \
          "SELECT GROUP_ID,OPT_STATION_ID as STATION_ID,'%s' as SHIFT_NO,'%s' as SHIFT_DATE, OPT_EMP_ID as EMP_ID,OPT_EMP_NAME as EMP_NAME,PAY_TYPE_ID,PAY_TYPE_NAME,COUNT(1)AS CNT,SUM(AMT)AS AMT,SUM(GIFT_AMT)AS GIFT_AMT,SUM(AMT+GIFT_AMT)as TOTAL_AMT " \
          "from eng_crm.history_charge WHERE GROUP_ID='%s' AND OPT_STATION_ID='%s' AND CREATED_TIME>='%s' AND CREATED_TIME<'%s' " \
          "GROUP BY GROUP_ID,OPT_STATION_ID,OPT_EMP_ID,OPT_EMP_NAME,PAY_TYPE_ID,PAY_TYPE_NAME ORDER BY PAY_TYPE_ID" % (
              db.DB_REPORT, param.shift_no, param.shift_date, param.group_id, param.station_id, param.shift_start_time,
              param.shift_end_time)
    print(sql)
    cur.execute(sql)
