# energy-script

python3 -m pip install --upgrade pip
pip install openpyxl
pip install pymysql
pip install pypyodbc
pip install pyinstaller==3.2.1

data.agent.py               华气厚普数据采集
stat_report_server_shift.py      服务端生成班次报表
stat_report_client_shift.py      客户端生成班次报表(exe)
daily_summary.py            每日销售充值汇总         每日凌晨12:05执行
activity_expire_status.py   修改已经过期的活动状态   每日凌晨12:00执行
stat_member_amt_history.py  会员会员余额汇总         每5分钟执行一次
stat_emp_pref.py            员工绩效报表             每5分钟执行一次
stat_report_server_day      日报/月汇总              JAVA程序调用



