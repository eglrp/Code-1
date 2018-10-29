# -*- coding: utf-8 -*-
import  xdrlib ,sys
import xlrd
import xlwt
#打开excel文件
def open_excel(file= 'test.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print (str(e))

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的索引  ，by_name：Sheet1名称
def excel_table_byname(file= '1234.xls', colnameindex=0, by_name=u'1234'):
    data = open_excel(file) #打开excel文件
    table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
    nrows = table.nrows #行数
    colnames = table.row_values(colnameindex) #某一行数据
    list =[] #装读取结果的序列
    for rownum in range(0, nrows): #遍历每一行的内容
         row = table.row_values(rownum) #根据行号获取行
         if row: #如果行存在
             app = [] #一行的内容
             for i in range(len(colnames)): #一列列地读取行的内容
                app.append(row[i])
             list.append(app) #装载数据
    return list


#将list中的内容写入一个新的file文件
def testXlwt(file = 'new.xls', list = []):
    book = xlwt.Workbook() #创建一个Excel
    sheet1 = book.add_sheet(file.split('.')[0]) #在其中创建一个名为hello的sheet
    i = 0 #行序号
    for app in list : #遍历list每一行
        j = 0 #列序号
        for x in app : #遍历该行中的每个内容（也就是每一列的）
            sheet1.write(i, j, x) #在新sheet中的第i行第j列写入读取到的x值
            j = j+1 #列号递增
        i = i+1 #行号递增
    # sheet1.write(0,0,'cloudox') #往sheet里第一行第一列写一个数据
    # sheet1.write(1,0,'ox') #往sheet里第二行第一列写一个数据
    book.save(file) #创建保存文件


def transformation_form(list_all,num_day_sort,name_list_unique):
    '''
    :param list_all:签到后的所有信息
    :param num_day_sort:日期按照从前到后排序 （作为列）
    :param name_list_unique:去重之后的名字  （作为行）
    :return:老师要求的形式
    '''

    # 按照名字来分组
    list_sorted_name=[]
    for name_list_unique_pre in name_list_unique:  # 名称
        temp = []
        for list_result_all_pre in list_all:
            if name_list_unique_pre in list_result_all_pre:
                temp.append(list_result_all_pre)
        list_sorted_name.append(temp)

    list_final_by_name=[]
    for list_sorted_name_pre in list_sorted_name:
        # 拿到学号
        number=list_sorted_name_pre[0][0]
        # 拿到名字
        name=list_sorted_name_pre[0][1]
        list_pre_final=[]
        list_pre_final.append(number)
        list_pre_final.append(name)
        for list_sorted_name_pre_pre in list_sorted_name_pre:
            # 早上7:00到10:00
            if int(list_sorted_name_pre_pre[2].split(' ')[-1].split(':')[0])  >= 6 and int(list_sorted_name_pre_pre[2].split(' ')[-1].split(':')[0]) <= 11:
                list_pre_final.append(list_sorted_name_pre_pre[2])
            # 下午13:00到16:00之间
            elif  int(list_sorted_name_pre_pre[2].split(' ')[-1].split(':')[0])  >= 12 and int(list_sorted_name_pre_pre[2].split(' ')[-1].split(':')[0])  <= 16 :
                list_pre_final.append(list_sorted_name_pre_pre[2])
            else:
                pass
        list_final_by_name.append(list_pre_final)

    # 转化为 列 为日期的形式
    # 遍历 所有日期

    result_final=[]
    # 写入头文件
    head=['学号','名字']
    for num_day_sort_pre in num_day_sort:
        head.append(num_day_sort_pre+'上午')
        head.append(num_day_sort_pre + '下午')
    result_final.append(head)




    # 读取一条记录
    for list_final_by_name_pre in list_final_by_name:
        result = []
        # 读取学号
        result.append(list_final_by_name_pre[0])
        # 读取名字
        result.append(list_final_by_name_pre[1])
        list_final_by_name_pre.pop(0)
        list_final_by_name_pre.pop(0)
        # 遍历所有时间
        for num_day_sort_pre in num_day_sort:
            # 先拿到当天所有记录
            record=[]
            for list_final_by_name_pre_pre in list_final_by_name_pre:
                if num_day_sort_pre in list_final_by_name_pre_pre:
                    record.append(list_final_by_name_pre_pre)
            #遍历当天所有记录
            if record:

                #拿到当天上午所有记录
                list_day_morning=[]
                for record_pre in record:
                    if int(record_pre.split(' ')[-1].split(':')[0]) < 12:
                        list_day_morning.append(record_pre)
                # 写入当天上午信息
                if list_day_morning:
                    # 当天上午若有多条数据，只拿第一条
                    # 判断 是否迟到  是否>8:40
                    if (int(list_day_morning[0].split(' ')[-1].split(':')[0])  == 8  and  int(list_day_morning[0].split(' ')[-1].split(':')[1])  > 40) or (int(list_day_morning[0].split(' ')[-1].split(':')[0])  > 8):
                         result.append(list_day_morning[0].split(' ')[-1]+'迟到')
                    else:
                        result.append(list_day_morning[0].split(' ')[-1] )
                else:
                    # 当天上午时间段内无记录
                    result.append('上午无记录')

                # 拿到当天下午所有记录
                list_day_afternoon=[]
                for record_pre in record:
                    if int(record_pre.split(' ')[-1].split(':')[0]) >= 12:
                        list_day_afternoon.append(record_pre)
                # 写入当天下午信息
                if list_day_afternoon:
                    # 当天下午若有多条数据，只拿第一条
                    # 判断 是否迟到  是否>14:40
                    if (int(list_day_afternoon[0].split(' ')[-1].split(':')[0]) == 14 and int( list_day_afternoon[0].split(' ')[-1].split(':')[1]) > 40) or (int(list_day_afternoon[0].split(' ')[-1].split(':')[0]) > 14):
                        result.append(list_day_afternoon[0].split(' ')[-1] + '迟到')
                    else:
                        result.append(list_day_afternoon[0].split(' ')[-1])
                else:
                    # 当天上午时间段内无记录
                    result.append('下午无记录')

            # 若当天记录为空，则 写入 迟到信息
            else:
                result.append('上午无记录')
                result.append('下午无记录')
        result_final.append(result)
    return result_final

def statistics(result):
    '''
    统计无记录、迟到名单
    '''
    # 去掉第一行的标题
    result.pop(0)
    result_final=[]
    # 写入头信息

    head=['名字','迟到次数','无记录次数']
    result_final.append(head)


    #遍历每一行，统计完结果放入list中
    for result_pre in result:
        temp = []
        temp.append(result_pre[0])  # 先写入名字
        late=0 #统计迟到次数
        no_record=0  #统计无记录次数
        #遍历第一行所有内容
        for people_pre in result_pre:
            if '迟到' in people_pre:
                late=late+1
            if '无记录' in people_pre:
                no_record=no_record+1
        # 过滤 正常数据
        if late==0 and no_record==0:
            pass
        else:
            temp.append(late)
            temp.append(no_record)
            result_final.append(temp)
    return result_final


def split_by_nummmmber(result):
    '''
    根据学号将结果分为几部分
    '''
    # 找出所有出现的不同年级的学号
    number_type=[]

    result_final=[]
    result_final.append(result[0])
    #去掉头信息
    result.pop(0)
    for result_pre in result:
        number_type.append(result_pre[0][0:3])
    number_type=list(set(number_type))
    number_type.sort()
    for number_type_pre in number_type:
        for result_pre in result:
            if number_type_pre in result_pre[0]:
                result_final.append(result_pre)
    return result_final




#主函数
def main():
   list_all=[]
   name_list=[]
   # XXX.xls ：原始表     XXX ：表中的第一页名字
   tables = excel_table_byname(file= '1234.xls', colnameindex=0, by_name=u'1234')
   # 去掉第一行的标题
   tables.pop(0)
   # 查询表格共有几天
   num_day=[]
   for row in tables:
       list_pre_info = []
       list_pre_info.append(row[1])
       list_pre_info.append(row[2])
       list_pre_info.append(row[4])
       num_day.append(row[4].split(' ')[0])
       list_all.append(list_pre_info)
       name_list.append(row[2])



   #日期按照从前到后排序（作为列）
   num_day_sort=sorted(list(set(num_day)))

   # 去重之后的名字 （作为行）
   name_list_unique=list(set(name_list))

   # 转化为 老师要求的形式
   result=transformation_form(list_all,num_day_sort,name_list_unique)

   # 根据学号将结果分为几部分
   result=split_by_nummmmber(result)

   # 将结果写入excel
   testXlwt('attendance.xls', result)

if __name__=="__main__":
    main()