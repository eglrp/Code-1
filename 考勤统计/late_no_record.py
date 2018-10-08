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
def excel_table_byname(file= '1234.xls', colnameindex=0):
    data = open_excel(file) #打开excel文件
    table = data.sheet_by_name(file.split('.')[0]) #根据sheet名字来获取excel中的sheet
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
    sheet1 = book.add_sheet(file.split('.')[0])
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



def statistics(result):
    '''
    统计无记录、迟到名单
    '''
    # 去掉第一行的标题
    result.pop(0)
    result_final=[]
    # 写入头信息

    head=['学号','名字','迟到次数','无记录次数']
    result_final.append(head)


    #遍历每一行，统计完结果放入list中
    for result_pre in result:
        temp = []
        temp.append(result_pre[0])  # 写入学号
        temp.append(result_pre[1])  # 写入名字
        late=0 #统计迟到次数
        no_record=0  #统计无记录次数
        #遍历第一行所有内容
        for people_pre in result_pre:
            if '迟到' in str(people_pre):
                late=late+1
            if '无记录' in str(people_pre):
                no_record=no_record+1
        # 过滤 正常数据
        if late==0 and no_record==0:
            pass
        else:
            temp.append(late)
            temp.append(no_record)
            result_final.append(temp)
    return result_final








#主函数
def main():
   list_all=[]
   name_list=[]
   # XXX.xls ：原始表     XXX ：表中的第一页名字
   tables = excel_table_byname(file= 'attendance.xls')

   # 统计无记录、迟到的名单
   result_late = statistics(tables)
   # 将结果写入excel
   testXlwt('late_no_record.xls', result_late)



if __name__=="__main__":
    main()