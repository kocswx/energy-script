import os


def prepare_path(file_name):
    father_path = os.path.abspath(os.path.dirname(file_name) + os.path.sep + ".")
    if not os.path.exists(father_path):
        os.makedirs(father_path)
