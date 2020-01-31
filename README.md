# energy-script

python3 -m pip install --upgrade pip
pip3 install pyinstaller==3.2.1
pip3 install openpyxl
pip3 install pypyodbc
pip3 install pymysql
pip3 install requests
pip3 install redis

data.agent.py               华气厚普数据采集
shift_report_server.py      服务端生成班次报表
shift_report_client.py      客户端生成班次报表(exe)
daily_summary.py            日销售充值汇总           每日凌晨12:05执行
activity_expire_status.py   修改已经过期的活动状态   每日凌晨12:00执行
stat_member_amt_history.py  会员会员余额汇总         每5分钟执行一次
stat_emp_pref.py            员工绩效报表             每5分钟执行一次

stat_report_client
stat_report_server      
stat_report_param       参数配置
stat_report_comm_data   生成数据

参数 
站点，开始日期，结束日期，班次号
			
			
			
			
			
			






