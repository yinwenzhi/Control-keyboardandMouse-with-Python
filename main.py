# introduction 
# https://pyautogui.readthedocs.io/en/latest/introduction.html
# https://hugit.app/posts/doc-pyautogui.html
# https://blog.csdn.net/waitforblack/article/details/80047253
import pyautogui
import pyperclip
# import matplotlib
# import pandas as pd
import os


from engine.utils import ExcelObject
from engine.gainname import GainName

# import engine.utils

def getFileNames(path):
  names = [name for name in os.listdir(path)
        if os.path.isfile(os.path.join(path, name))]
  # print(names)
  return names

# 'data\\企业列表.xls'
def dealOneExcelFilebyName(excelFileName):
  # 定原始excel文件路径
  # excelpath=excelpath
  #pcklfile='C:\\Users\\VI\\Desktop\\result.txt'
  # 'data\\'
  
  print('excelFileName：',excelFileName)
  try:
    excelfile  =  ExcelObject(excelFileName)
    print('excel file get:',excelFileName)
  except :
    print()
    print('++++++++++++++++++++++++++++++')
    print('只支持“.xls”文件，请检查文件格式')

  # 读取excel并根据第二列的值是否符合要求来更改第二列的值
  # df = pd.read_excel(excelpath,headers = None ,usecols = [0,1]) # pandas 读取不包含第一行
  originalSheet = excelfile.Sheet_data
  rows = excelfile.Sheet_data.nrows
  gainfullname = GainName()
  # 查看原始数据
  # print('+++++++++++++++++++++++',rows)
  # for i in range(rows):
  #   print(excelfile.Sheet_data.row_values(i))
  # print('+++++++++++++++++++++++',rows)

  compnum = 0

  for row in range(rows):

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
        print('文件：{}'.format(excelFileName))
        print('第{}行，正在获取“{}”的公司全称：'.format(row,shortname.strip()))

        try:
          # fullname = str(shortname) +'new'
          # engname = str(shortname) +'eng'

          fullname ,engname = gainfullname.search( shortname )
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
  print('文件{}：共{}行，查询完毕，共获得{}个公司全称'.format(excelFileName,rows,compnum))

  
  # # 检查修改结果
  # excelfile  =  ExcelObject(excelpath)
  # print('rows+++++++++++++++++++++++',rows)
  # for i in range(rows):
  #   print(excelfile.Sheet_data.row_values(i))
  # print('rows+++++++++++++++++++++++',rows)


    # print(row,df.values[row,0],df.values[row,1],excelfile.original_data.values[row,0],excelfile.original_data.values[row,1])
    # print(df.values[row,1],type(df.values[row,1]),)
    # print(type(df.values[row,1]))

def dealOneExcelFilebyDate(path):
  excelFileName = path.split('\\')[-1]
  print('获得文件：',path)
  date = excelFileName.strip().split('.')[0][4:].split('_')
  print('日期：',date)
  try:
    excelfile  =  ExcelObject(path)
    # print('excel file get:',path)
  except :
    print()
    print('++++++++++++++++++++++++++++++')
    print('只支持“.xls”文件，请检查文件格式')

  # 读取excel并根据第二列的值是否符合要求来更改第二列的值
  originalSheet = excelfile.Sheet_data
  rows = excelfile.Sheet_data.nrows
  gainfullname = GainName()

  compnum = 0

  # 设置日期 
  gainfullname.setDate(date[0],date[1])

  for row in range(rows):
    # 使用xldt判断是否写入
    shortname = originalSheet.cell(row,0).value
    fullname = originalSheet.cell(row,1).value
    
    if fullname == '' or fullname == None :
      if shortname != '' :
        print('---------------------------------------------------------------------')
        print('请勿操作鼠标和键盘，否则将会造成程序错误')
        print('文件：{}'.format(excelFileName))
        print('第{}行，正在获取“{}”的公司全称：'.format(row,shortname.strip()))

        try:
          # fullname = str(shortname) +'new'
          # engname = str(shortname) +'eng'

          fullname ,engname = gainfullname.byPage(row,rows)
          # fullname ,engname = gainfullname.search(shortname)

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
  print('文件{}：共{}行，查询完毕，共获得{}个公司全称'.format(excelFileName,rows,compnum))


def main():

  path = 'C:\\Users\\dai\\Desktop\\simul\\投资机构正在处理文件'
  # path = 'C:\\Users\\dai\\Desktop\\simul\\测试'

  # 初始化程序
  print('starting...')
  print('如果把鼠标光标在屏幕左上角，程序就会停止')#如果把鼠标光标在屏幕左上角，PyAutoGUI函数就
  pyautogui.PAUSE = 0.7
  pyautogui.FAILSAFE = True #如果把鼠标光标在屏幕左上角，PyAutoGUI函数就会产生pyautogui.FailSafeException异常。
  screenWidth, screenHeight = pyautogui.size()
  # currentMouseX, currentMouseY = pyautogui.position()
  print('screenWidth, screenHeight = ',screenWidth, screenHeight )

  names = getFileNames(path)
  for idx,name in  enumerate(names):
    print('***********************************************************************')
    print('正在处理第{}个文件:   {}   '.format(idx+1,name))
    
    # get company name by search
    name =  os.path.join(path, name)
    # dealOneExcelFilebyName(name)

    # get company name by data
    dealOneExcelFilebyDate(name)

    print('第{}个文件   {}   处理完毕'.format(idx+1,name))

if __name__ == "__main__":
  pyautogui.FAILSAFE = True
  time = 0
  main()

  # while True:
  #   time +=1
  #   main()
  #   if time ==1:
  #     break