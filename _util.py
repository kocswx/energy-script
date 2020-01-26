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

