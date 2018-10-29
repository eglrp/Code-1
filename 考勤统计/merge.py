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
    # table = data.sheet_by_name(file.split('.')[0]) #根据sheet名字来获取excel中的sheet
    table = data.sheet_by_name('late_no_record')  # 根据sheet名字来获取excel中的sheet
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
    # sheet1 = book.add_sheet(file.split('.')[0])
    sheet1 = book.add_sheet('late_no_record')
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



def merge(tables_merge):
    '''
    统计无记录、迟到名单
    '''

    tables_merge.sort()

    # 获得所有人的学号
    list_number = []
    for tables_merge_pre in tables_merge:
        list_number.append(tables_merge_pre[0]) #学号
    list_number=list(set(list_number))
    list_number.sort()


    list_person=[]
    for list_number_pre in list_number:
        pre = []
        for tables_merge_pre in tables_merge:
            if list_number_pre in tables_merge_pre:
                pre.append(tables_merge_pre)
        list_person.append(pre)

    result_final=[]
    # 写入头信息
    head_0=['每月累积汇总表']
    head=['学号','名字','迟到次数','无记录次数']
    result_final.append(head_0)
    result_final.append(head)
    for  list_person_pre in list_person:
        #统计每个人的记录
        list_pre=[]
        list_pre.append(list_person_pre[0][0]) #写入学号
        list_pre.append(list_person_pre[0][1]) #写入姓名
        late=0.0
        norecord=0.0
        for list_person_pre_pre in list_person_pre:
            late=late+list_person_pre_pre[2]
            norecord = norecord + list_person_pre_pre[3]
        list_pre.append(int(late))
        list_pre.append(int(norecord))
        result_final.append(list_pre)

    return result_final


#主函数
def main():

   tables_A = excel_table_byname(file= 'late_no_record.xls')

   tables_B = excel_table_byname(file= 'all.xls')


   # 去掉头信息
   tables_A.pop(0)
   tables_B.pop(0)

   tables_merge = tables_A + tables_B

   # 合并统计无记录、迟到的名单
   result_late = merge(tables_merge)
   # 将结果写入excel
   testXlwt('merge_result.xls', result_late)

if __name__=="__main__":
    main()