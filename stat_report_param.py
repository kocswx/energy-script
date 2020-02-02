# coding:utf-8


class DbConfig:
    DB_ERP = 'eng_erp'
    DB_CRM = 'eng_crm'
    DB_ORDER = 'eng_order'
    DB_REPORT = 'eng_report'


class StatParam:
    group_id = None  # 外部输入
    station_id = None  # 外部输入
    begin_date = None  # 外部输入
    end_date = None  # 外部输入
    shift_no = None  # 外部输入
    file_name = None  # 外部输入

    station_name = None
    shift_date = None
    shift_start_time = None
    shift_end_time = None
    shift_emp_name = None

    param_string = None

    def init(self, param):
        if len(param) != 7:
            raise Exception('param error:len(argv) != 7')
        self.group_id = param[1]
        self.station_id = param[2]
        self.begin_date = param[3]
        self.end_date = param[4]
        self.shift_no = param[5]
        self.file_name = param[6]
        if self.station_id == 'NULL':
            self.station_id = None
        if self.begin_date == 'NULL':
            self.begin_date = None
        if self.end_date == 'NULL':
            self.end_date = None
        if self.shift_no == 'NULL':
            self.shift_no = None

    def copy(self):
        result = StatParam()
        result.group_id = self.group_id
        result.station_id = self.station_id
        result.begin_date = self.begin_date
        result.end_date = self.end_date
        result.shift_no = self.shift_no
        result.file_name = self.file_name
        return result
