# 导入统计汇总相关包
import numpy as np
import pandas as pd
import matplotlib as mpl
# 导入地理编码模块
import compare_location as CL
# 导入OpenPyXL处理EXCEL的xlsx文件
import openpyxl

def add1(argument):
    if argument == None:
        return 1
    else:
        argument += 1
        return argument

# out_check(id,out_time,out_address)
# 功能：判断外出地址是否符合要求
# 参数：员工编号，外出时间，外出地址
# 返回：若同一天能查到2条同一位置的签卡记录则返回True，否则返回False
def out_check(oc_df, id, out_time, out_address):
    time_list = []
    for i in oc_df.index:
        oc_id = oc_df.loc[i, "员工编号"]
        oc_day = oc_df.loc[i, "签卡时间"][:10]
        oc_address = oc_df.loc[i, "地点"]
        if oc_id == id and oc_day == out_time and CL.compare_location(out_address, oc_address, 800) == 1:
            time_list.append(oc_df.loc[i, "签卡时间"][11:])

    if len(time_list) >= 2:
        return True
    else:
        return False


# 调整显示格式
pd.set_option('display.max_columns', 10)
pd.set_option('display.max_rows', 1000)
pd.set_option('display.width', 200)
# 读入外出记录单文件
df = pd.read_excel("DATA/IN外出记录单.xlsx", header=2, usecols=[1, 2, 3, 4, 5, 6, 9, 11])
# 调整各列的顺序
order = ["部门", "员工编号", "姓名", "人员类别", "外出时间", "外出类型", "外出地址", "相关项目编号"]
df = df[order]
# 增加外出校核列
df["项目类型"] = None
df["外出校核"] = None

# 扫描项目编号列，判断是否为pipeline项目、非pipeline项目、无项目编号
for i in df.index:
    project_type = ""

    if pd.isnull(df.loc[i, "相关项目编号"]):
        project_type = "无项目编号"
    else:
        if df.loc[i, "相关项目编号"][0:1] == 'P':
            project_type = "Pipeline项目"
        else:
            project_type = "非Pipeline项目"
    df.loc[i, "项目类型"] = project_type

# 读入外勤签卡记录
df_check = pd.read_excel("DATA/IN外勤签卡记录.xlsx", header=0, usecols=[1, 3, 4])
# print(df_check)

# 较验外出合规性
for i in df.index:
    if out_check(df_check, df.loc[i, "员工编号"], df.loc[i, "外出时间"], df.loc[i, "外出地址"]) == True:
        df.loc[i, "外出校核"] = True
    else:
        df.loc[i, "外出校核"] = False
# df中是完全的外勤记录信息
# print(df)
# df.to_excel("data/OUT外勤汇总.xlsx",columns=["部门","员工编号","姓名","人员类别","外出类型","项目类型","外出校核"])

# 使用openpyxl读入模板，填数后，输出到excel文件中
wb = openpyxl.load_workbook(filename="data/区域销售考勤统计及分析模板.xlsx")
ws_sale = wb['销售类']
ws_tech = wb['技术类']
ws_comm = wb['综合类']

# 逐行扫描外出记录，分别判断并填入
for row in df.index:
    if df.loc[row,'外出校核'] == True:
        # 根据人员类型先分销售类和技术类
        # =======处理销售类=======
        if df.loc[row,'人员类别'] == '销售':
            for r in range(2,ws_sale.max_row+1):
                if ws_sale.cell(r,2).value == df.loc[row,'姓名']:
                    # 外出次数累加
                    ws_sale.cell(r,9).value = add1(ws_sale.cell(r,9).value)
                    # 项目类型
                    if df.loc[row,'项目类型'] == 'Pipeline项目':
                        ws_sale.cell(r,10).value = add1(ws_sale.cell(r,10).value)
                    if df.loc[row,'项目类型'] == '非Pipeline项目':
                        ws_sale.cell(r,11).value = add1(ws_sale.cell(r,11).value)
                    if df.loc[row,'项目类型'] == '无项目编号':
                        ws_sale.cell(r,12).value = add1(ws_sale.cell(r,12).value)
                    # 根据外出类型和项目类型，判断填表列
                    # -------商务非正式交流-------
                    if df.loc[row,'外出类型'] == '商务非正式交流':
                        ws_sale.cell(r,13).value = add1(ws_sale.cell(r,13).value)
                        if  df.loc[row,'项目类型'] == 'Pipeline项目':
                            ws_sale.cell(r, 14).value = add1(ws_sale.cell(r, 14).value)
                        if  df.loc[row,'项目类型'] == '非Pipeline项目':
                            ws_sale.cell(r, 15).value = add1(ws_sale.cell(r, 15).value)
                        if  df.loc[row,'项目类型'] == '无项目编号':
                            ws_sale.cell(r, 16).value = add1(ws_sale.cell(r, 16).value)
                    # -------其他-------
                    if df.loc[row,'外出类型'] == '其他':
                        ws_sale.cell(r,17).value = add1(ws_sale.cell(r,17).value)
                        if  df.loc[row,'项目类型'] == '无项目编号':
                            ws_sale.cell(r, 18).value = add1(ws_sale.cell(r, 18).value)
        # =======处理技术类=======
        if df.loc[row,'人员类别'] == '技术售前' or df.loc[row,'人员类别'] == '技术售后':
            for r in range(2, ws_tech.max_row + 1):
                if ws_tech.cell(r, 2).value == df.loc[row, '姓名']:
                    # 外出次数累加
                    ws_tech.cell(r, 9).value = add1(ws_tech.cell(r, 9).value)
                    # 项目类型
                    if df.loc[row, '项目类型'] == 'Pipeline项目':
                        ws_tech.cell(r, 10).value = add1(ws_tech.cell(r, 10).value)
                    if df.loc[row, '项目类型'] == '非Pipeline项目':
                        ws_tech.cell(r, 11).value = add1(ws_tech.cell(r, 11).value)
                    if df.loc[row, '项目类型'] == '无项目编号':
                        ws_tech.cell(r, 12).value = add1(ws_tech.cell(r, 12).value)
                    # 根据外出类型和项目类型，判断填表列
                    # -------商务非正式交流-------
                    if df.loc[row, '外出类型'] == '商务非正式交流':
                        ws_tech.cell(r, 13).value = add1(ws_tech.cell(r, 13).value)
                        if  df.loc[row,'项目类型'] == 'Pipeline项目':
                            ws_tech.cell(r, 14).value = add1(ws_tech.cell(r, 14).value)
                        if  df.loc[row,'项目类型'] == '非Pipeline项目':
                            ws_tech.cell(r, 15).value = add1(ws_tech.cell(r, 15).value)
                        if  df.loc[row,'项目类型'] == '无项目编号':
                            ws_tech.cell(r, 16).value = add1(ws_tech.cell(r, 16).value)
                    # -------客户交流-------
                    if df.loc[row, '外出类型'] == '客户交流':
                        ws_tech.cell(r, 17).value = add1(ws_tech.cell(r, 17).value)
                        if  df.loc[row,'项目类型'] == 'Pipeline项目':
                            ws_tech.cell(r, 18).value = add1(ws_tech.cell(r, 18).value)
                        if  df.loc[row,'项目类型'] == '非Pipeline项目':
                            ws_tech.cell(r, 19).value = add1(ws_tech.cell(r, 19).value)
                        if  df.loc[row,'项目类型'] == '无项目编号':
                            ws_tech.cell(r, 20).value = add1(ws_tech.cell(r, 20).value)
                    # -------投标相关活动-------
                    if df.loc[row, '外出类型'] == '投标相关活动':
                        ws_tech.cell(r, 21).value = add1(ws_tech.cell(r, 21).value)
                        if df.loc[row, '项目类型'] == 'Pipeline项目':
                            ws_tech.cell(r, 22).value = add1(ws_tech.cell(r, 22).value)
                        if df.loc[row, '项目类型'] == '非Pipeline项目':
                            ws_tech.cell(r, 23).value = add1(ws_tech.cell(r, 23).value)
                        if df.loc[row, '项目类型'] == '无项目编号':
                            ws_tech.cell(r, 24).value = add1(ws_tech.cell(r, 24).value)
                    # -------售前客户培训和售后客户培训-------
                    if df.loc[row, '外出类型'] == '售前客户培训' or df.loc[row, '外出类型'] == '售前客户培训':
                        ws_tech.cell(r, 25).value = add1(ws_tech.cell(r, 25).value)
                        if df.loc[row, '项目类型'] == 'Pipeline项目':
                            ws_tech.cell(r, 26).value = add1(ws_tech.cell(r, 26).value)
                        if df.loc[row, '项目类型'] == '非Pipeline项目':
                            ws_tech.cell(r, 27).value = add1(ws_tech.cell(r, 27).value)
                        if df.loc[row, '项目类型'] == '无项目编号':
                            ws_tech.cell(r, 28).value = add1(ws_tech.cell(r, 28).value)
                    # -------安装实施次数/首次安装-------
                    if df.loc[row, '外出类型'] == '首次安装':
                        ws_tech.cell(r, 29).value = add1(ws_tech.cell(r, 29).value)
                    # -------故障排除次数/售后现场服务-------
                    if df.loc[row, '外出类型'] == '售后现场服务':
                        ws_tech.cell(r, 30).value = add1(ws_tech.cell(r, 30).value)
                    # -------巡检次数/巡检服务-------
                    if df.loc[row, '外出类型'] == '巡检服务':
                        ws_tech.cell(r, 31).value = add1(ws_tech.cell(r, 31).value)
                    # -------其他-------
                    if df.loc[row,'外出类型'] == '其他':
                        ws_tech.cell(r,32).value = add1(ws_tech.cell(r,32).value)
                        if  df.loc[row,'项目类型'] == '无项目编号':
                            ws_tech.cell(r, 33).value = add1(ws_tech.cell(r, 33).value)

    # 外出校验未通过，外出异常次数进行累计
    else:
        if df.loc[row,'人员类别'] == '销售':
            for r in range(2,ws_sale.max_row+1):
                if ws_sale.cell(r,2).value == df.loc[row,'姓名']:
                    ws_sale.cell(r,8).value = add1(ws_sale.cell(r,8).value)
        if df.loc[row,'人员类别'] == '技术售前' or df.loc[row,'人员类别'] == '技术售后':
            for r in range(2, ws_tech.max_row + 1):
                if ws_tech.cell(r, 2).value == df.loc[row, '姓名']:
                    ws_tech.cell(r, 8).value = add1(ws_tech.cell(r, 8).value)


# 读取OUT考勤报表，计算考勤异常次数
wb_check = openpyxl.load_workbook(filename="data/OUT考勤报表.xlsx")
ws_check = wb_check.active

for row_index in range(2, ws_check.max_row + 1):
    # 获取姓名和考勤异常次数
    name = ws_check.cell(row_index, 3).value
    count = 0
    for col_index in range(4, ws_check.max_column + 1):
        cell_str = str(ws_check.cell(row_index, col_index).value)
        if cell_str.find('<') != -1:
            count += 1
    # 扫描ws_sale加入考勤异常次数
    for r in range(2, ws_sale.max_row + 1):
        if ws_sale.cell(r, 2).value == name and count != 0:
            ws_sale.cell(r, 6).value = count
    # 扫描ws_tech加入考勤异常次数
    for r in range(2, ws_tech.max_row + 1):
        if ws_tech.cell(r, 2).value == name  and count != 0:
            ws_tech.cell(r, 6).value = count

# 输出文件
output_file = "data/OUT分析报表.xlsx"
wb.save(filename=output_file)