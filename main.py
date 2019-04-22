# introduction 
# https://pyautogui.readthedocs.io/en/latest/introduction.html
# https://hugit.app/posts/doc-pyautogui.html
# https://blog.csdn.net/waitforblack/article/details/80047253
import pyautogui
import pyperclip
import matplotlib
import pandas as pd



from engine.utils import ExcelObject
from engine.gainname import GainName

# import engine.utils

# 初始化程序
print('starting...')
print('如果把鼠标光标在屏幕左上角，程序就会停止')#如果把鼠标光标在屏幕左上角，PyAutoGUI函数就
pyautogui.PAUSE = 2
pyautogui.FAILSAFE = True #如果把鼠标光标在屏幕左上角，PyAutoGUI函数就会产生pyautogui.FailSafeException异常。
screenWidth, screenHeight = pyautogui.size()
# currentMouseX, currentMouseY = pyautogui.position()
print('screenWidth, screenHeight = ',screenWidth, screenHeight )
gainfullname = GainName()

#指定原始excel文件路径
excelpath='data\\企业列表.xls'
#pcklfile='C:\\Users\\VI\\Desktop\\result.txt'
try:
  excelfile  =  ExcelObject(excelpath)
  print('excel file get')
except:
  print('++++++++++++++++++++++++++++++')
  print('只支持“.xls”文件，请检查文件格式')

# 读取excel并根据第二列的值是否符合要求来更改第二列的值
# df = pd.read_excel(excelpath,headers = None ,usecols = [0,1]) # pandas 读取不包含第一行
originalSheet = excelfile.Sheet_data
rows = excelfile.Sheet_data.nrows

# 查看原始数据
# print('+++++++++++++++++++++++',rows)
# for i in range(rows):
#   print(excelfile.Sheet_data.row_values(i))
# print('+++++++++++++++++++++++',rows)

for row in range(rows):

  compnum = 0
  # 使用pandas判断是否写入
  # if type(df.values[row,1]) != str or row == None :
  #   if type(df.values[row,0]) == str :
  #     companyname = str(df.values[row,0]) +'new'
  #     excelfile.write(0,row+1,1,companyname) # write函数和pd的readexce行数不匹配所以row+1

  # 使用xldt判断是否写入
  shortname = originalSheet.cell(row,0).value
  fullname = originalSheet.cell(row,1).value
  
  if fullname == '' or fullname == None :
    if shortname != '' :

      print('---------------------------------------------------------------------')
      print('请勿操作鼠标和键盘，否则将会造成程序错误,移动鼠标到窗口左上角以终止程序')
      print('第{}行，正在获取“{}”的公司全称：'.format(row,shortname.strip()))

      try:
        # fullname = str(shortname) +'new'
        # engname = str(shortname) +'eng'

        fullname ,engname= gainfullname.search( shortname )
      except:
        pass
      
      if fullname !='':
        print('获取成功，公司名称为：',fullname,engname)
        excelfile.write(0,row,1,fullname) 
        excelfile.write(0,row,2,engname)
        compnum +=1
      else:
        print('获取“{}”公司名称失败,将跳过并搜索下一个公司'.format(shortname.strip()))
    else:
      print('第{}行，没有公司简称，无法搜索'.format(row))

print('---------------------------------------------------------------------')
print('共{}行，查询完毕，共获得{}个公司全称'.format(rows,compnum))

# # 检查修改结果
# excelfile  =  ExcelObject(excelpath)
# print('rows+++++++++++++++++++++++',rows)
# for i in range(rows):
#   print(excelfile.Sheet_data.row_values(i))
# print('rows+++++++++++++++++++++++',rows)


  # print(row,df.values[row,0],df.values[row,1],excelfile.original_data.values[row,0],excelfile.original_data.values[row,1])
  # print(df.values[row,1],type(df.values[row,1]),)
  # print(type(df.values[row,1]))