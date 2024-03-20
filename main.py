import xlwings as xw
from datetime import datetime as dt



# 获取数据表中的数据用于后续处理
# 打开excel,参数visible表示处理过程是否可视,add_book表示是否打开新的Excel程序
with xw.App(visible=False, add_book=False) as app:
    # 读取工作薄
    book = app.books.open(r'./test.xlsx')

    sht = book.sheets('sheet2')  # 指定名称获取sheet工作表

    #获取最右下角的单元格
    cell = sht.used_range.last_cell
    rows = cell.row  #获取该单元格的行数
    columns = cell.column  #获取该单元格的列数

    #获取点信息列表
    pid = sht.range((2,2),(2,columns)).value
    #print(pid)
    #获取点列表编号,转换为字典，用于后续写表排序
    id_unit = {key: index for index, key in enumerate(pid)}

    #获取说有参数
    data_raw = sht.range((4,1),(rows,columns)).value
    #data_raw[行][列]表示单元格数据，如[0][i]为第0行的第i个元素
    # 保存
    #book.save('./test.xlsx')
    book.close()

#机组负荷所在列为：data_raw[][id_unit['UNIT2_20MKA01CE005XQ21']+1]
load_1 = id_unit['UNIT1_10MKA01CE005XQ21'] + 1
load_2 = id_unit['UNIT2_20MKA01CE005XQ21'] + 1
list_1_500_isempty = True
list_1_600_isempty = True
list_1_700_isempty = True
list_1_800_isempty = True
list_1_900_isempty = True
list_1_950_isempty = True
list_1_500 = []
list_1_600 = []
list_1_700 = []
list_1_800 = []
list_1_900 = []
list_1_950 = []
list_1_500_top = []
list_1_600_top = []
list_1_700_top = []
list_1_800_top = []
list_1_900_top = []
list_1_950_top = []
list_2_500_isempty = True
list_2_600_isempty = True
list_2_700_isempty = True
list_2_800_isempty = True
list_2_900_isempty = True
list_2_950_isempty = True
list_2_500 = []
list_2_600 = []
list_2_700 = []
list_2_800 = []
list_2_900 = []
list_2_950 = []
list_2_500_top = []
list_2_600_top = []
list_2_700_top = []
list_2_800_top = []
list_2_900_top = []
list_2_950_top = []
#查找负荷
for i in range(0,len(data_raw)):
    #将当前日期转换为datetime对象
    # fmt_date = dt.strptime(data_raw[i][0], "%Y-%m-%d %H:%M:%S")
    #分离出日、月和日、时和分，单独的日(int)用于判断一天之内是否存在多条符合要求的记录
    # part_day = int(fmt_date.strftime("%d"))
    # part_day = data_raw[i][0].day
    part_date = data_raw[i][0].strftime("%m-%d")
    part_time = data_raw[i][0].strftime("%H:%M")
    #获取需要的数据
    my_list_1 = [part_date,
                 data_raw[i][id_unit['UNIT1_D6S002-FA'] + 1],
                 data_raw[i][id_unit['UNIT1_D2S801-CF'] + 1],
                 data_raw[i][id_unit['UNIT1_10HSA01CP001'] + 1],
                 data_raw[i][id_unit['UNIT1_10HSA02CP001'] + 1],
                 data_raw[i][id_unit['UNIT1_10HNA10CPSUM'] + 1],
                 data_raw[i][id_unit['UNIT1_10HNA20CPSUM'] + 1],
                 data_raw[i][id_unit['UNIT1_10HNA10CP006'] + 1],
                 data_raw[i][id_unit['UNIT1_10HNA10CP004'] + 1],
                 data_raw[i][id_unit['UNIT1_10HNA10CP002'] + 1] - data_raw[i][id_unit['UNIT1_10HNA10CP006'] + 1],
                 data_raw[i][id_unit['UNIT1_10HNA20CP002'] + 1] - data_raw[i][id_unit['UNIT1_10HNA20CP005'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1002.SH0017.AALM170603.I'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1002.SH0080.AALM080501.I'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1002.SH0080.AALM080502.I'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1007.SH0028.AALM023206.I'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1007.SH0028.AALM023207.I'] + 1],
                 data_raw[i][id_unit['As_DPU1037.SH0010.AALM011101.I'] + 1],
                 data_raw[i][id_unit['As_DPU1037.SH0010.AALM011102.I'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1002.SH0017.AALM170603.I'] + 1] - data_raw[i][id_unit['As_DPU1037.SH0010.AALM011101.I'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1002.SH0080.AALM080501.I'] + 1] - data_raw[i][id_unit['DeSu_DPU1002.SH0080.AALM080502.I'] + 1],
                 data_raw[i][id_unit['DeSu_DPU1007.SH0028.AALM023206.I'] + 1] - data_raw[i][id_unit['DeSu_DPU1007.SH0028.AALM023207.I'] + 1],
                 data_raw[i][id_unit['As_DPU1037.SH0010.AALM011101.I'] + 1] - data_raw[i][id_unit['As_DPU1037.SH0010.AALM011102.I'] + 1],
                 part_time]
    my_list_2 = [part_date,
                data_raw[i][id_unit['UNIT2_20HSA01CP001'] + 1],
                data_raw[i][id_unit['UNIT2_20HSA02CP001'] + 1],
                data_raw[i][id_unit['UNIT2_20HNA10CPSUM'] + 1],
                data_raw[i][id_unit['UNIT2_20HNA20CPSUM'] + 1],
                data_raw[i][id_unit['UNIT2_AIDFINLP'] + 1],
                data_raw[i][id_unit['UNIT2_20HNA10CP004'] + 1],
                data_raw[i][id_unit['UNIT2_20HNA10CP002'] + 1] - data_raw[i][id_unit['UNIT2_AIDFINLP'] + 1],
                data_raw[i][id_unit['UNIT2_20HNA20CP002'] + 1] - data_raw[i][id_unit['UNIT2_AIDFINLP'] + 1],
                data_raw[i][id_unit['DeSu_DPU1012.SH0017.AALM170603.I'] + 1],
                data_raw[i][id_unit['DeSu_DPU1012.SH0080.AALM080501.I'] + 1],
                data_raw[i][id_unit['DeSu_DPU1012.SH0080.AALM080502.I'] + 1],
                data_raw[i][id_unit['DeSu_DPU1016.SH0028.AALM023206.I'] + 1],
                data_raw[i][id_unit['DeSu_DPU1016.SH0028.AALM023207.I'] + 1],
                data_raw[i][id_unit['DeSu_DPU1012.SH0080.AALM080501.I'] + 1] - data_raw[i][id_unit['DeSu_DPU1012.SH0080.AALM080502.I'] + 1],
                data_raw[i][id_unit['DeSu_DPU1016.SH0028.AALM023206.I'] + 1] - data_raw[i][id_unit['DeSu_DPU1016.SH0028.AALM023207.I'] + 1],
                part_time]

    #采集二号机数据
    if 498 < data_raw[i][load_2] < 502:
        if list_2_500_isempty == True:
            list_2_500_top = my_list_2
            list_2_500.append(my_list_2)
            list_2_500_isempty = False
        elif my_list_2[0] != list_2_500_top[0]:
            list_2_500_top = my_list_2
            list_2_500.append(my_list_2)
    elif 598 < data_raw[i][load_2] < 602:
        if list_2_600_isempty == True:
            list_2_600_top = my_list_2
            list_2_600.append(my_list_2)
            list_2_600_isempty = False
        elif my_list_2[0] != list_2_600_top[0]:
            list_2_600_top = my_list_2
            list_2_600.append(my_list_2)
    elif 698 < data_raw[i][load_2] < 702:
        if list_2_700_isempty == True:
            list_2_700_top = my_list_2
            list_2_700.append(my_list_2)
            list_2_700_isempty = False
        elif my_list_2[0] != list_2_700_top[0]:
            list_2_700_top = my_list_2
            list_2_700.append(my_list_2)
    elif 798 < data_raw[i][load_2] < 802:
        if list_2_800_isempty == True:
            list_2_800_top = my_list_2
            list_2_800.append(my_list_2)
            list_2_800_isempty = False
        elif my_list_2[0] != list_2_800_top[0]:
            list_2_800_top = my_list_2
            list_2_800.append(my_list_2)
    elif 898 < data_raw[i][load_2] < 902:
        if list_2_900_isempty == True:
            list_2_900_top = my_list_2
            list_2_900.append(my_list_2)
            list_2_900_isempty = False
        elif my_list_2[0] != list_2_900_top[0]:
            list_2_900_top = my_list_2
            list_2_900.append(my_list_2)
    elif 948 < data_raw[i][load_2] < 952:
        if list_2_950_isempty == True:
            list_2_950_top = my_list_2
            list_2_950.append(my_list_2)
            list_2_950_isempty = False
        elif my_list_2[0] != list_2_950_top[0]:
            list_2_950_top = my_list_2
            list_2_950.append(my_list_2)

    #采集一号机数据
    if 498 < data_raw[i][load_1] < 502:
        if list_1_500_isempty == True:
            list_1_500_top = my_list_1
            list_1_500.append(my_list_1)
            list_1_500_isempty = False
        elif my_list_1[0] != list_1_500_top[0]:
            list_1_500_top = my_list_1
            list_1_500.append(my_list_1)
    elif 598 < data_raw[i][load_1] < 602:
        if list_1_600_isempty == True:
            list_1_600_top = my_list_1
            list_1_600.append(my_list_1)
            list_1_600_isempty = False
        elif my_list_1[0] != list_1_600_top[0]:
            list_1_600_top = my_list_1
            list_1_600.append(my_list_1)
    elif 698 < data_raw[i][load_1] < 702:
        if list_1_700_isempty == True:
            list_1_700_top = my_list_1
            list_1_700.append(my_list_1)
            list_1_700_isempty = False
        elif my_list_1[0] != list_1_700_top[0]:
            list_1_700_top = my_list_1
            list_1_700.append(my_list_1)
    elif 798 < data_raw[i][load_1] < 802:
        if list_1_800_isempty == True:
            list_1_800_top = my_list_1
            list_1_800.append(my_list_1)
            list_1_800_isempty = False
        elif my_list_1[0] != list_1_800_top[0]:
            list_1_800_top = my_list_1
            list_1_800.append(my_list_1)
    elif 898 < data_raw[i][load_1] < 902:
        if list_1_900_isempty == True:
            list_1_900_top = my_list_1
            list_1_900.append(my_list_1)
            list_1_900_isempty = False
        elif my_list_1[0] != list_1_900_top[0]:
            list_1_900_top = my_list_1
            list_1_900.append(my_list_1)
    elif 948 < data_raw[i][load_1] < 952:
        if list_1_950_isempty == True:
            list_1_950_top = my_list_1
            list_1_950.append(my_list_1)
            list_1_950_isempty = False
        elif my_list_1[0] != list_1_950_top[0]:
            list_1_950_top = my_list_1
            list_1_950.append(my_list_1)

#获取当前月份
now_month = data_raw[0][0].strftime("%Y.%m")
#生成表格名称
new_file_1 = '#1炉风烟系统差压分析（' + now_month + '）.xlsx'
new_file_2 = '#2炉风烟系统差压分析（' + now_month + '）.xlsx'

#将数据写入新表格
with xw.App(visible=False, add_book=False) as app:
    # 打开现有的 Excel 文件
    wb = xw.Book('#1炉模板.xlsx')
    # 获取要操作的sheet并写入数据
    sht = wb.sheets['950MW']
    sht.range('A2').value = list_1_950
    sht = wb.sheets['900MW']
    sht.range('A2').value = list_1_900
    sht = wb.sheets['800MW']
    sht.range('A2').value = list_1_800
    sht = wb.sheets['700MW']
    sht.range('A2').value = list_1_700
    sht = wb.sheets['600MW']
    sht.range('A2').value = list_1_600
    sht = wb.sheets['500MW']
    sht.range('A2').value = list_1_500

    #保存文件
    wb.save(new_file_1)

    # 关闭 Excel 应用
    wb.close()

    print(f'文件已保存为 {new_file_1}')

    wb = xw.Book('#2炉模板.xlsx')
    # 获取要操作的sheet并写入数据
    sht = wb.sheets['950MW']
    sht.range('A2').value = list_2_950
    sht = wb.sheets['900MW']
    sht.range('A2').value = list_2_900
    sht = wb.sheets['800MW']
    sht.range('A2').value = list_2_800
    sht = wb.sheets['700MW']
    sht.range('A2').value = list_2_700
    sht = wb.sheets['600MW']
    sht.range('A2').value = list_2_600
    sht = wb.sheets['500MW']
    sht.range('A2').value = list_2_500

    # 保存文件
    wb.save(new_file_2)

    # 关闭 Excel 应用
    wb.close()

    # 退出 Excel 应用
    xw.App().quit()

    print(f'文件已保存为 {new_file_2}')