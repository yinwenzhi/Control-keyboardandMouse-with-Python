# introduction 
# https://pyautogui.readthedocs.io/en/latest/introduction.html
# https://hugit.app/posts/doc-pyautogui.html
# https://blog.csdn.net/waitforblack/article/details/80047253
import pyautogui
import pyperclip
import matplotlib
import pandas as pd


from engine.utils import typewirte 
from engine.utils import getimage
from engine.utils import getword
from engine.utils import ExcelObject
# import engine.utils


#指定原始excel文件路径
excelpath='data\\企业列表.xls'
#pcklfile='C:\\Users\\VI\\Desktop\\result.txt'
print('excel file get')
try:
  excelfile  =  ExcelObject(excelpath)
except:
  print('++++++++++++++++++++++++++++++')
  print('只支持“.xls”文件，请检查文件格式')
# 读取excel并根据第二列的值是否符合要求来更改第二列的值
# df = pd.read_excel(excelpath,headers = None ,usecols = [0,1]) # pandas 读取不包含第一行
originalSheet = excelfile.Sheet_data
rows = excelfile.Sheet_data.nrows

print('+++++++++++++++++++++++',rows)
for i in range(rows):
  print(excelfile.Sheet_data.row_values(i))
print('+++++++++++++++++++++++',rows)

for row in range(rows):

  # 使用pandas判断是否写入
  # if type(df.values[row,1]) != str or row == None :
  #   if type(df.values[row,0]) == str :
  #     companyname = str(df.values[row,0]) +'new'
  #     excelfile.write(0,row+1,1,companyname) # write函数和pd的readexce行数不匹配所以row+1

  # 使用xldt判断是否写入
  if originalSheet.cell(row,1).value == '' or originalSheet.cell(row,1) == None :
    if originalSheet.cell(row,0).value != '' :
      chiname = str(originalSheet.cell(row,0).value) +'new'
      engname = str(originalSheet.cell(row,0).value) +'eng'
      excelfile.write(0,row,1,chiname) 
      excelfile.write(0,row,2,engname)
excelfile  =  ExcelObject(excelpath)

# 打印修改结果
print('rows+++++++++++++++++++++++',rows)
for i in range(rows):
  print(excelfile.Sheet_data.row_values(i))
print('rows+++++++++++++++++++++++',rows)

  # print(row,df.values[row,0],df.values[row,1],excelfile.original_data.values[row,0],excelfile.original_data.values[row,1])
  # print(df.values[row,1],type(df.values[row,1]),)
  # print(type(df.values[row,1]))