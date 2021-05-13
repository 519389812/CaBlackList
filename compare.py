from settings import *
import re
import os
import datetime
from dateutil.relativedelta import relativedelta


def change_date_format(date):  # 25AUG18
    try:
        month = {"JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
                 "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"}
        date = date[-2:] + month[date[2:5]] + date[:2]
        date = datetime.datetime.strptime(date, "%y%m%d")
        if date.date() > datetime.datetime.now().date():
            date = date - relativedelta(years=100)
        date = date.strftime("%Y-%m-%d")
        return date
    except:
        return ''


if __name__ == '__main__':
    if not blacklist_path:
        alert_box('欢迎使用本程序！首次使用您还未设置blacklist文件，请先设置。', '欢迎')
        set_blacklist_path()
        config, blacklist_path = reload_config()
    else:
        select = yes_no_box('欢迎使用本程序！\n当前使用的blacklist文件是：%s，是否重新设置？' % blacklist_path, '欢迎')
        if select == '是':
            set_blacklist_path()
            config, blacklist_path = reload_config()
    blacklist, error = open_file(blacklist_path, blacklist_columns_set, converters={'DOB': str})
    if blacklist is False:
        alert_box('设置出错，原因：%s' % error, '错误')
        os._exit(0)
    blacklist['DOB'] = blacklist['DOB'].str.split(' ').str[0]
    blacklist_dob_list = blacklist['DOB'].to_list()
    blacklist['Last Name'] = blacklist['Last Name'].str.replace(' ', '')
    blacklist['First Name'] = blacklist['First Name'].str.replace(' ', '')
    blacklist['Middle'] = blacklist['Middle'].str.replace(' ', '')
    blacklist['name_1'] = blacklist['Last Name'] + blacklist['First Name'] + blacklist['Middle']
    blacklist['name_2'] = blacklist['Last Name'] + blacklist['Middle'] + blacklist['First Name']
    blacklist['name_3'] = blacklist['First Name'] + blacklist['Last Name'] + blacklist['Middle']
    blacklist['name_4'] = blacklist['First Name'] + blacklist['Middle'] + blacklist['Last Name']
    blacklist['name_5'] = blacklist['Middle'] + blacklist['Last Name'] + blacklist['First Name']
    blacklist['name_6'] = blacklist['Middle'] + blacklist['First Name'] + blacklist['Last Name']
    compare_file_path = open_file_box('选择要匹配的文件', filetypes=[('xls', '.xls'), ('xlsx', '.xlsx')])
    compare_file, error = open_file(compare_file_path, compare_columns_set, converters={'出生日期': str})
    if compare_file is False:
        alert_box('设置出错，原因：%s' % error, '错误')
        os._exit(0)
    compare_file['姓名'] = compare_file['姓名'].str.replace('/', '').str.replace(' ', '').str.upper()
    compare_file['出生日期'] = compare_file['出生日期'].apply(change_date_format)
    compare_file['匹配'] = ''
    name_pattern = re.compile(r'(.*)MS|MR|MRS|SD|UM|CHD|MISS|MSTR|PROF|EXST|CBBG')
    match_blacklist = []
    for index_, row in compare_file.iterrows():
        if re.match(name_pattern, row['姓名']):
            name = re.findall(name_pattern, row['姓名'])[0].strip()
        else:
            name = row['姓名'].strip()
        if row['出生日期'] == '' or row['性别'] == '':
            for i, r in blacklist.iterrows():
                if name in [r['name_1'], r['name_2'], r['name_3'], r['name_4'], r['name_5'], r['name_6']]:
                    compare_file.loc[index_, '匹配'] = '信息不足，疑似'
                    break
        else:
            if row['出生日期'] in blacklist_dob_list:
                blacklist_match_dob = blacklist[blacklist['DOB'] == row['出生日期']]
                for i, r in blacklist_match_dob.iterrows():
                    if row['性别'] == r['Gender']:
                        if name in [r['name_1'], r['name_2'], r['name_3'], r['name_4'], r['name_5'], r['name_6']]:
                            compare_file.loc[index_, '匹配'] = '是'
                            break
    dt = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    compare_file.to_excel('%s%s%s.xlsx' % (results_dir, os.sep, dt))
    alert_box('匹配完成！共有 %s 人匹配或无法判断，请查看 %s.xlsx 文件' % (compare_file[compare_file['匹配'] != ''].shape[0], dt), '完成')
