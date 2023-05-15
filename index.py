import pandas as pd
import datetime
import os
from chinese_calendar import is_workday


path = os.getcwd()#获取当前路径

# 功能函数
#1、合并xlsx
def mergeXLSX(path1,path2,sheetname=""):
    # 如果sheetname=“”，说明xlsx里面只有一个sheet，把路径1文件夹里面的所有xlsx合并成1个xlsx
    # 如果sheetname不为空，说明xlsx里面可能不止一个sheet，把路径1文件夹里面的所有xlsx的指定sheet合并成1个xlsx
    # 保存到path2的文件夹

    # 获取当前目录下的文件列表
    file_list = os.listdir(path+'/data/path1')
    print(file_list)
    for filename in file_list:
        sheet = pd.read_excel(filename,sheet_name=None)
        print('sheet',sheet)
        for k,v in sheet.items():
            v = v.to_dict(orient='records')
            print(k,v)

def mergeCSV(path1,path2):#把路径1文件夹里面的所有csv合并成1个csv
    pass

def isworkday(date):#date格式是YYYY-MM-DD
    # 判断date是否为工作日，https://github.com/LKI/chinese-calendar
    # 输出True/False
    if date.find("-") == -1:
        print('err:date格式是YYYY-MM-DD')
    else:
        dateArr = date.split('-')
        return is_workday(datetime.date(int(dateArr[0]),int(dateArr[1]),int(dateArr[2])))

# 正文----------------------------------------------------------------------------------------------------
def mergeCSVtoXLSX(path):
    # 把考勤打卡记录表里面的所有csv文件合并，并保存成xlsx，保存位置还是考勤打卡记录 - 固定文件夹
    # 命名规则：考勤打卡记录合并.xlsx
     pass

def addTagInRecord(path,facelist):
    # 1、把“考勤打卡记录合并.xlsx”的每一列（除了打卡时间）的tab格式去掉
    # 2、新增“日期”、“时间”列，相当于把打卡时间拆开，但是打卡时间这一列不要动
    # 3、新增“是否刷脸”列，如果状态=“进”或者“出”，或者地点=facelist里面其中一个，那么值=“是”，否则值=“另行判断”
    # 4、新增id列（保存的时候放在最前面），id=员工编码+” “+日期，比如：21033887 2022-07-30
    # 5、输出保存在原文件夹，命名为：（原名）+预处理.xlsx  注：这个+是存在的
    pass

# 按人按天统计
def analyseRecord(path):
    # 1、读取该文件夹里面文件名带有“预处理”的xlsx，如果不止一个就合并
    # 2、按人按天统计当天最早打卡时间、当天最晚打卡时间、当天打卡次数、当天刷脸次数、是否存在迟到早退，
    # 也就是说id唯一，id就是这个时候辅助用的，一个id一行
    # 当天最早打卡时间、最晚打卡时间、当天打卡次数这三列不用说了
    # 当天刷脸次数=“是否刷脸”列的“是”有多少个
    # 是否存在迟到早退=（文本）
#       如果“当前班次”=“办公班”，
#           那么判断当天最早打卡时间是否早于9点，当天最晚打卡时间是否早于5点半
#           晚于9点“迟到”，早于5点半“早退”
#           如果2种都有，就是“迟到和早退”
#       如果“当前班次”不等于“办公班”，
#           那就pass

    # 输出保存还是在本文件夹
    # 命名：考勤打卡记录+统计.xlsx
    # 表格字段有：id，部门，员工编码，员工姓名，打卡时间，当前班次，当天最早打卡时间，当天最晚打卡时间，当天打卡次数，当天刷脸次数，是否存在迟到早退
    pass






