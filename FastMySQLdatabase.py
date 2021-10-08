import json
import xlrd
import xlwt
import numpy as np
## 定义列表
# dict = {'table_name':'ss','table_column_name':'qq','table_column_type':'int'}

Table = np.dtype({'names':['table_name','table_column_name','table_column_type','table_column_CanBN','table_column_Default','table_column_constrain'],'formats':['S64','S64','S8','S4','S32','S32']})
MySql_Table=np.array([('NULL','NULL','NULL','NULL','NULL','NULL')]*100,dtype = Table)
## 获取列表
readbook = xlrd.open_workbook(r'C:\Users\YGS_Tu\PycharmProjects\pythonProject\数据库设计说明书.xls')      # 获取所有sheet
print('table numbers is',len(readbook.sheet_names()))# [u'sheet1', u'sheet2']

File = open("test.txt", "w")

for table_Names in readbook.sheet_names():
    Sheet_current = readbook.sheet_by_name(table_Names)
    table_Names_SQL = Sheet_current.cell_value(0,1)

    MySql_String_A = ""     #head  ID self add
    MySql_String_C = ""     #end

    MySql_String_B = {}     #columns
    MySql_String_B_1 = ""     #columns

    MySql_String_A = "CREATE TABLE IF NOT EXISTS `"+table_Names_SQL+"`( `ID` INT UNSIGNED AUTO_INCREMENT"
    MySql_String_C = ", PRIMARY KEY ( `ID` ) )ENGINE=InnoDB DEFAULT CHARSET=utf8; "
    ##
    # print("current table is",table_Names_SQL)

    Counter_MySql_Table = 0

    for Counter_Table_rows in range(3,Sheet_current.nrows):
        if((len(Sheet_current.cell_value(Counter_Table_rows,1))==0) | (Sheet_current.cell_value(Counter_Table_rows,1)=='ID')):
            continue
        B1="`"+Sheet_current.cell_value(Counter_Table_rows,1)+"`"
        B1.replace(" ", "")


        if(B1[1:7] == "S_Volu"):
            B1.replace(" ", "")
            B1.replace("\n", "")
            print(B1 ,"-",ord(B1[10]),"-")
        B1.replace(" ","")
        B1.replace("\n","")
        if len(Sheet_current.cell_value(Counter_Table_rows,3))==0 :
            B2="NVarChar(100)"
        else:
            B2=Sheet_current.cell_value(Counter_Table_rows,3)
        B3 = "NOT NULL"
        if(B2 == 'DATE'):
            MySql_String_B[Counter_MySql_Table]=B1+" "+B2
        else:
            MySql_String_B[Counter_MySql_Table]=B1+" "+B2+" "+B3
        Counter_MySql_Table = Counter_MySql_Table+1

    for M_String_B_counter in range(0,Counter_MySql_Table):
        MySql_String_B_1 = MySql_String_B_1 + " ,"+ MySql_String_B[M_String_B_counter]

    MySql_String=MySql_String_A+MySql_String_B_1+MySql_String_C


    # print(MySql_String)
    File.write(MySql_String)
    File.write("\n")