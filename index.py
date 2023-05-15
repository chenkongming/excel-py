import pandas as pd
import datetime
import os
from chinese_calendar import is_workday
from datetime import datetime, time

basePath = os.getcwd() # 获取当前路径

# 功能函数
def check_punch_time(punch_times):
    late_time = time(9, 0, 0)   # 9点
    early_time = time(5, 30, 0) # 5点半
    
    is_late = datetime.strptime(punch_times[0], "%H:%M:%S").time() > late_time
    is_early = datetime.strptime(punch_times[1], "%H:%M:%S").time() < early_time
    
    if is_late and is_early:
        return "迟到和早退"
    elif is_late:
        return "迟到"
    elif is_early:
        return "早退"
    else:
        return "正常"

def get_min_max_time(time_list):
    # 将时间字符串转换为datetime对象
    time_objects = [datetime.strptime(time, "%H:%M:%S") for time in time_list]
    
    # 获取最小时间和最大时间
    min_time = min(time_objects).strftime("%H:%M:%S")
    max_time = max(time_objects).strftime("%H:%M:%S")
    
    return min_time, max_time

def mergeArrayToObj(array):
    # 根据数组中的key，合并对应的key值
    # [{a:1,b:2},{a:11,b:22}] => {a:[1,11],b:[2,22]}
    merged_dict = {}
    for item in array:
        for key, value in item.items():
            if key in merged_dict:
                merged_dict[key].append(value)
            else:
                merged_dict[key] = [value]
    return merged_dict

def mergeDataById(input_array):
    output_dict = {}
    id_field = 'id'
    punch_time_field = '时间'
    current_shift_field = "当前班次"
    face_field = "是否刷脸"
    bumen_field = "部门"
    ygbm_field = "员工编码"
    ygxm_field = "员工姓名"
    dksj_field = "打卡时间"
    for item in input_array:
        item_id = item[id_field]
        punch_time = item[punch_time_field]
        current_shift = item[current_shift_field]
        face_value = item[face_field]
        bumen = item[bumen_field]
        ygbm = item[ygbm_field]
        ygxm = item[ygxm_field]
        dksj = item[dksj_field]

        if item_id not in output_dict:
            output_dict[item_id] = {
                punch_time_field: [],
                current_shift_field: [],
                face_field: [],
                bumen_field : [],
                ygbm_field : [],
                ygxm_field : [],
                dksj_field : [],
            }

        output_dict[item_id][punch_time_field].append(punch_time)
        output_dict[item_id][current_shift_field].append(current_shift)
        output_dict[item_id][face_field].append(face_value)

        output_dict[item_id][bumen_field].append(bumen)
        output_dict[item_id][ygbm_field].append(ygbm)
        output_dict[item_id][ygxm_field].append(ygxm)
        output_dict[item_id][dksj_field].append(dksj)

    return output_dict

def mergeXLSXData(dirPath,xlsx_files):
     # 创建一个空的数组用于存储合并后的数据
    mergeArray = []
    for filename in xlsx_files:
        sheet = pd.read_excel(dirPath + filename,sheet_name=None)
        for k,v in sheet.items():
            v = v.to_dict(orient='records')
            mergeArray +=v
    outputData = mergeDataById(mergeArray)
    return outputData

#1、合并xlsx
def mergeXLSX(path1="",path2="",sheetname=""):
    # 如果sheetname=“”，说明xlsx里面只有一个sheet，把路径1文件夹里面的所有xlsx合并成1个xlsx
    # 如果sheetname不为空，说明xlsx里面可能不止一个sheet，把路径1文件夹里面的所有xlsx的指定sheet合并成1个xlsx
    # 保存到path2的文件夹
    # 获取当前目录下的文件列表
    if path1 == '' or path2 == "":
       print('mergeXLSX函数参数错误')
       return
    dirPath = basePath + path1 +'/' # 相对当前路径
    # 获取目录下所有的文件名
    file_names = os.listdir(dirPath)
     # 过滤出 CSV 文件
    xlsx_files = [file for file in file_names if file.endswith('.xlsx')]
     # 创建一个空的数组用于存储合并后的数据
    mergeArray = []
    for filename in xlsx_files:
        sheet = pd.read_excel(dirPath + filename,sheet_name=None)
        for k,v in sheet.items():
            v = v.to_dict(orient='records')
            if sheetname!="":
                if k == sheetname:
                    mergeArray +=v
            else:
                mergeArray +=v
    outputData = mergeArrayToObj(mergeArray)
    output = pd.DataFrame(outputData)
    output.to_excel(basePath + path2,index = False)   #index默认是True，导致第一列是0,1,2,3,....,设置为False后可以去掉第一列。
    print('xlsx合并成功:' + basePath + path2)

def mergeCSV(path1,path2):#把路径1文件夹里面的所有csv合并成1个csv
    # 保存到path2的文件夹
    # 获取当前目录下的文件列表
    if path1 == '' or path2 == "":
       print('mergeCSV函数参数错误')
       return
    dirPath = basePath + path1 +'/' # 相对当前路径
    # 获取目录下所有的文件名
    file_names = os.listdir(dirPath)
    # 过滤出 CSV 文件
    csv_files = [file for file in file_names if file.endswith('.csv')]
    # 创建一个空的 DataFrame 用于存储合并后的数据
    merged_data = pd.DataFrame()
    # 遍历每个 CSV 文件并合并数据
    for file in csv_files:
        file_path = os.path.join(dirPath, file)
        df = pd.read_csv(file_path, encoding='utf-8-sig')  # 读取 CSV 文件
         # 去除列名（表头）的空格
        df.columns = df.columns.str.strip()
            # 去除每一行的前后空格
        df = df.apply(lambda x: x.str.strip() if x.dtype == 'object' else x)
        merged_data = pd.concat([merged_data, df])  # 合并数据
    # 输出合并后的数据
    # print(merged_data)

    # 可选：将合并后的数据保存为新的 CSV 文件
    merged_data.to_csv(basePath + path2, index=False,encoding='utf-8-sig')
    print('csv合并成功:' + basePath + path2)


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
    # 获取目录下所有的文件名
    dirPath = basePath + path +'/' # 相对当前路径
    file_names = os.listdir(dirPath)
    # 过滤出CSV文件
    csv_files = [file for file in file_names if file.endswith('.csv')]
    # 创建一个空的DataFrame用于存储合并后的数据
    merged_data = pd.DataFrame()
    # 遍历每个CSV文件并合并数据
    for file in csv_files:
        file_path = os.path.join(path, file)
        try:
            df = pd.read_csv(file_path)  # 读取CSV文件
            # 去除列名（表头）的空格
            df.columns = df.columns.str.strip()
            # 去除每一行的前后空格
            df = df.apply(lambda x: x.str.strip() if x.dtype == 'object' else x)
            merged_data = pd.concat([merged_data, df])  # 合并数据
        except pd.errors.ParserError as e:
            print(f"Error parsing file: {file} - {e}")
    # 导出为XLSX文件
    output_file = dirPath + '考勤打卡记录合并.xlsx'
    merged_data.to_excel(output_file, index=False)
    print('mergeCSVtoXLSX成功：',output_file)

def addTagInRecord(path="",facelist=[]):
    # 1、把“考勤打卡记录合并.xlsx”的每一列（除了打卡时间）的tab格式去掉
    # 2、新增“日期”、“时间”列，相当于把打卡时间拆开，但是打卡时间这一列不要动
    # 3、新增“是否刷脸”列，如果状态=“进”或者“出”，或者地点=facelist里面其中一个，那么值=“是”，否则值=“另行判断”
    # 4、新增id列（保存的时候放在最前面），id=员工编码+” “+日期，比如：21033887 2022-07-30
    # 5、输出保存在原文件夹，命名为：（原名）+预处理.xlsx  注：这个+是存在的
    # 读取Excel文件
    if path == "":
       print("path参数不能为空")
       return
    filename = basePath + path
    outputfilename = filename.split(".xlsx")[0] + "+预处理.xlsx"
    df = pd.read_excel(filename)
    # 添加"日期"列和"时间"列
    df['日期'] = df['打卡时间'].str[:10]
    df['时间'] = df['打卡时间'].str[11:]

    # 添加"id"列
    df.insert(0, 'id', df['员工编码'].astype(str) + ' ' + df['日期'].astype(str))

    # # # 添加"是否刷脸"列
    df['是否刷脸'] = df.apply(lambda row: '是' if row['状态'] in ['进', '出'] or row['地点'] in facelist else '另行判断', axis=1)

    # # # 保存修改后的Excel文件
    df.to_excel(outputfilename, index=False)

    print("预处理表格成功：" + outputfilename)
    # pass

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

    if path == "":
       print('analyseRecord函数参数不能为空')
       return
    dirPath = basePath + path +'/' # 相对当前路径
    # 获取目录下所有的文件名
    file_names = os.listdir(dirPath)
    xlsx_files = [file for file in file_names if file.endswith('.xlsx') and "预处理" in file]
    XLSXData = mergeXLSXData(dirPath,xlsx_files)
    
    # id列
    id_list = list(XLSXData.keys())
    # 最大时间和最小时间
    max_time_list = []
    min_time_list = []
    # 当天打卡次数
    punch_times = []
    #当天刷脸次数
    face_times = []
    #迟到or早退
    late_and_leave_early = []
    # 部门
    bm_list = []
    # 员工编码，员工姓名，打卡时间，当前班次
    ygbm_list = []
    ygxm_list = []
    dksj_list = []
    dqbc_list = []

    for key, value in XLSXData.items():
        for skey, svalue in value.items():
            if skey == '时间':
               min_time, max_time = get_min_max_time(svalue)
               max_time_list.append(max_time)
               min_time_list.append(min_time)
               punch_times.append(len(svalue))
               if value['当前班次'].count("办公班") > 0:
                  late_and_leave_early.append(check_punch_time([min_time,max_time]))
               else:
                  late_and_leave_early.append('-')
            if skey == '是否刷脸':
               face_times.append(svalue.count("是"))
            if skey == '部门':
               bm_list.append(svalue[0])   
            if skey == '员工编码':
               ygbm_list.append(svalue[0])
            if skey == '员工姓名':
               ygxm_list.append(svalue[0])
            if skey == '打卡时间':
               dksj_list.append(svalue)
            if skey == '当前班次':
               dqbc_list.append(svalue[0])    

    data = {
    'id': id_list,
    '部门':bm_list,
    '员工编码':ygbm_list,
    '员工姓名':ygxm_list,
    '打卡时间':dksj_list,
    '当天最早打卡时间': min_time_list,
    '当天最晚打卡时间':max_time_list,
    '当天打卡次数':punch_times,
    '当天刷脸次数':face_times,
    '是否存在迟到早退':late_and_leave_early,
    }
    # id，部门，员工编码，员工姓名，打卡时间，当前班次，当天最早打卡时间，当天最晚打卡时间，当天打卡次数，当天刷脸次数，是否存在迟到早退
    # 使用字典创建DataFrame
    df = pd.DataFrame(data)
    df.to_excel(dirPath+'考勤打卡记录+统计.xlsx', index=False)
    print('analyseRecord成功')
    pass


# isworkdayFlag = isworkday("2023-10-08")
# print(f'是否为工作日: {isworkdayFlag}')

# mergeXLSX('/data/path1','/data/path2/result.xlsx','员工表')
# mergeXLSX('/data/path1','/data/path2/result.xlsx')
# mergeCSV('/data/path1','/data/path2/result.csv');
# mergeCSVtoXLSX('./data/考勤打卡记录 - 固定')
# addTagInRecord('./data/考勤打卡记录 - 固定/考勤打卡记录合并.xlsx')
analyseRecord('./data/考勤打卡记录 - 固定')


