# ! /usr/bin env python3
# -*- coding: utf-8 -*-
# author: Pluto
import datetime
import json
import os
import re
import shutil
import sys
import pandas as pd
from styleframe import StyleFrame

if __name__ != '__main__':
    os._exit(0)


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


LOCAL_TIME = datetime.datetime.now()

FILE_PATH = "course1.xlsx"

START_TIME = "9.1"
if len(sys.argv) == 2:
    FILE_PATH = sys.argv[1]
elif len(sys.argv) == 3:
    FILE_PATH = sys.argv[1]
    START_TIME = sys.argv[2]

START_TIME = datetime.datetime.strptime(START_TIME, "%m.%d")
START_TIME = datetime.datetime(LOCAL_TIME.year, START_TIME.month, START_TIME.day)
# 当前日期
CURRENT_TIME = START_TIME

print('读取文件名：%s' % FILE_PATH)
print('当前时间：%s' % LOCAL_TIME)
print('起始时间：%s' % START_TIME)

if not FILE_PATH.endswith(".xlsx"):
    print('仅支持.xlsx文件！请更改文件格式后重试！')
    os._exit(0)

data = {}
if os.path.exists(FILE_PATH):
    data = format_course(FILE_PATH)
else:
    print('%s该文件不存在！' % FILE_PATH)
    os._exit(0)

course_list = {}
for week in range(1, get_end_week(data) + 1):
    course_list["第%s周" % week] = get_target_week_course(data, week)

# 创建文件
if not os.path.isdir('courses'):
    os.mkdir('courses')

writer = pd.ExcelWriter('courses/%s' % FILE_PATH)

for course_index in course_list:
    index = ['0102', '0304', '0506', '0708', '091011']
    df = pd.DataFrame(course_list[course_index], index=index)
    sf = StyleFrame(df)
    # 设置第一行第一列的值为第X周
    sf.index.name = course_index

    # 设置每周的显示日期
    colums = []
    for v in sf.columns:
        title = "%s\n%s-%s" % (v, CURRENT_TIME.month, CURRENT_TIME.day)
        colums.append(title)
        CURRENT_TIME += datetime.timedelta(days=1)

    sf.columns = colums


    #  计算每列的最大字符宽度

    def get_max_width(v):
        s = str(v).split("\n")
        s = map(lambda x: len(x), s)
        return max(s)


    max_width = (
        sf.applymap(get_max_width).agg(max).values
    )
    #  计算每列的最大字符宽度

    max_height = [1]
    for i in index:
        arr = map(lambda x: len(str(x).split("\n")), df.loc[i])
        max_height.append(max(arr))
    # 设置宽度
    for row in range(1, 8):
        sf.set_column_width(row, max(10, max_width[row - 1]) * 2)
    # 设置高度
    for row in range(1, 7):
        sf.set_row_height(row, max(3, max_height[row - 1] + 1.5) * 15)

    # 转化成excel
    sf.to_excel(writer, sheet_name=course_index, index=True)
    writer.save()

    print('%s 写入成功！' % course_index)

# 写出JSON
with open("courses/%s.json" % FILE_PATH, mode="w") as f:
    f.write(json.dumps(course_list))
    print("courses/%s.json 写入成功！" % FILE_PATH)
