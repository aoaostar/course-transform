# ! /usr/bin env python3
# -*- coding: utf-8 -*-
# author: Pluto

import json
import os
import re
import shutil
import sys

import numpy as np
import pandas as pd
from styleframe import StyleFrame

if __name__ != '__main__':
    exit()


# 转化课程具体时间为数组
def format_course_time(str):
    str = re.sub('\(.+?\)\[.+?\]', '', str)
    split = str.split(',')
    data = []
    for v in split:
        if '-' in v:
            v_split = v.split('-')
            rg = range(int(v_split[0]), int(v_split[1]) + 1)
            data.extend(rg)
        else:
            data.extend([v])
    return data


# 转化课程为字典
def format_course(file) -> dict:
    df = pd.read_excel(file, sheet_name=0, header=2, names=None, index_col=0)

    columns = df.columns
    data = {}
    for col in range(7):
        data[columns[col]] = {}
        for row in range(5):
            course = df.iloc[row, col]
            if not pd.isnull(course):
                split = str.split(course, "\n")
                data[columns[col]][df.index[row]] = {}
                for i in range(len(split)):
                    info = str.split(split[i], "◇")
                    info = [x.strip() for x in info if x.strip() != '']
                    course_time = format_course_time(info[2])
                    for course_time_index in course_time:
                        data[columns[col]][df.index[row]][int(course_time_index)] = "\n".join(info)
            else:
                data[columns[col]][df.index[row]] = None
    return data


# 获取结束周
def get_end_week(course_list):
    max_num = 0
    for a in course_list:
        if course_list[a]:
            for b in course_list[a]:
                if course_list[a][b]:
                    for c in course_list[a][b]:
                        max_num = max(int(c), max_num)
    return max_num


# 获取目标周的课程
def get_target_week_course(data: dict, week: int) -> dict:
    course_list = {}
    for a in data:
        if data[a]:
            course_list[a] = {}
            for b in data[a]:
                if data[a][b]:
                    if week in data[a][b]:
                        course_list[a][b] = data[a][b][week]
    return course_list


if len(sys.argv) == 2:
    filepath = sys.argv[1]
else:
    filepath = "course.xlsx"

print('读取文件名：%s' % filepath)
if os.path.exists(filepath):
    data = format_course("course.xlsx")
else:
    print('%s该文件不存在！' % filepath)
    exit()
print('全部课表：%s' % json.dumps(data))
course_list = {}
for week in range(1, get_end_week(data) + 1):
    course_list["第%s周" % week] = get_target_week_course(data, week)
print('分周课表：%s' % json.dumps(course_list))

# 清空courses
if os.path.isdir('courses'):
    shutil.rmtree('courses', ignore_errors=True)
if not os.path.isdir('courses'):
    os.mkdir('courses')

for course_index in course_list:

    df = pd.DataFrame(course_list[course_index], index=['0102', '0304', '0506', '0708', '091011'])
    sf = StyleFrame(df)
    sf.index.name = course_index
    #  计算每列表头的字符宽度
    column_widths = (
        sf.columns.to_series().apply(lambda x: len(str(x))).values
    )
    #  计算每列的最大字符宽度
    max_widths = (
        sf.applymap(lambda x: len(str(x))).agg(max).values
    )
    for row in range(2, 7):
        sf.set_column_width(row, column_widths[row - 1] + 2)
    for col in range(1, 8):
        sf.set_column_width(col, max_widths[col - 1] + 2)

    # 取前两者中每列的最大宽度
    widths = np.max(np.array([column_widths, max_widths]))
    writer = sf.ExcelWriter('courses/%s.xlsx' % course_index)
    sf.to_excel(writer, index=True)
    print('%s 写入成功！' % course_index)
