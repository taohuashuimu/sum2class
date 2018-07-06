# coding=utf-8
# show the point of 统计成绩
# 2018.6.20   RHBZJ  created   V0.1

import pandas as pd

project = str(input("请输入需要统计的课程名称："))
number = str(input("请输入需要统计的班级（如电气1631）："))
score_level = int(input("请输入平时成绩比重（1：50%，2：60%）："))
df_late = pd.read_excel((project + " " + number + ".xlsx"), sheet_name='考勤')
df_good = pd.read_excel((project + " " + number + ".xlsx"), sheet_name='平时表现')
df_project = pd.read_excel((project + " " + number + ".xlsx"), sheet_name='项目')
df_end = pd.read_excel((project + " " + number + ".xlsx"), sheet_name='期末')
output_name = '输出总成绩' + ' ' + project + ' ' + number

# 根据“名单”统计分数
df_name = pd.read_excel((project + " " + number + ".xlsx"), sheet_name='名单')
parts = [df_name, df_late, df_good, df_project, df_end]
df_dest = pd.concat(parts, join_axes=[df_name.columns])
df_dest = df_dest.groupby('姓名').sum()
df_dest = df_dest.sort_values(by='学号', ascending=1)
df_dest = df_dest.reset_index()
if score_level == 1:
    level = 0.5
else:
    level = 0.6
df_dest["总分"] = (df_dest.加减分 * level) + (df_dest.期末成绩 * (1-level))

# 修改加减分为平时成绩
names = ['姓名', '学号', '平时成绩', '期末成绩', '总分']
df_dest.columns = names

# 输出结果
df_dest.to_excel(output_name + ".xlsx")
