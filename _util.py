import os
import datetime


def prepare_path(file_name):
    father_path = os.path.abspath(os.path.dirname(file_name) + os.path.sep + ".")
    if not os.path.exists(father_path):
        os.makedirs(father_path)


def get_now_full():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_now_year_month():
    return datetime.datetime.now().strftime("%Y-%m")


def get_data_start(date):
    return "%s 00:00:00" % date


def get_data_end(date):
    return "%s 23:59:59" % date


MEMBER_STATUS_DICT = {1: '正常', -1: '注销', 0: '冻结', -2: '冻结'}

MEMBER_TYPE_DICT = {1: '散户', 2: '总账户', 3: '子账户'}
