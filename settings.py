import configparser
from box_body import *
import pandas as pd


blacklist_dir = 'blacklist'
config_dir = 'config'
config_file = '%s%sconfig.ini' % (config_dir, os.sep)
results_dir = 'results'


blacklist_columns_set = {'SATAL ID', 'Last Name', 'First Name', 'Middle', 'DOB', 'Gender'}
compare_columns_set = {'姓名', '出生日期', '性别'}


def make_dir(dir_path):
    if not os.path.exists(dir_path):
        os.mkdir(dir_path)


for dir_ in [blacklist_dir, config_dir, results_dir]:
    make_dir(dir_)


def check_config_file():
    if not os.path.exists(config_file):
        with open(config_file, 'w') as f:
            pass
    config = configparser.ConfigParser()
    config.read(config_file)
    if not config.has_section('blacklist'):  # 检查是否存在section
        config.add_section('blacklist')
        with open(config_file, 'w') as f:
            config.write(f)
    if not config.has_option('blacklist', 'file_path'):  # 检查是否存在该option
        config.set('blacklist', 'file_path', '')
        with open(config_file, 'w') as f:
            config.write(f)


def reload_config():
    config = configparser.ConfigParser()
    config.read(config_file)
    blacklist_path = config.get('blacklist', 'file_path')
    return config, blacklist_path


check_config_file()
config, blacklist_path = reload_config()


def set_blacklist_path():
    blacklist_path = open_file_box(initialdir=blacklist_dir, filetypes=[('xls', '.xls'), ('xlsx', '.xlsx')])
    config.set('blacklist', 'file_path', blacklist_path)
    with open(config_file, 'w') as f:
        config.write(f)


def open_file(file_path, columns_set, converters):
    # try:
    file = pd.read_excel(file_path, converters=converters)
    file.fillna('', inplace=True)
    file = file[list(columns_set)]
    # except:
    #     return False, '文件格式错误，或读取文件失败'
    if set(list(file.columns)).intersection(columns_set) == columns_set:
        return file, ''
    else:
        return False, '文件未按格式要求设置'


if __name__ == '__main__':
    pass
