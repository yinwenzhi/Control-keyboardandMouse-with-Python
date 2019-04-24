import pyautogui
import pyperclip
# from aip import AipOcr

# from urllib.request import urlopen,Request
# from bs4 import BeautifulSoup
# import re
import xlrd
import xlwt
from xlutils.copy import copy
import pandas as pd 
import time
import datetime




class getimage():
  def __init__(self,height,wight):
    self.height = height
    self.wight = wight

  def getscreenshot(self):
  
    im = pyautogui.screenshot('my_screenshot.png',region=(0,0, self.height, self.wight))
    print('截图位置height: ',self.height,' height: ',self.wight)
    return im

class typewirte():
  def __init__(self,words):
    self.writing = words

  #  PyAutoGUI中文输入需要用粘贴实现
  #  Python 2版本的pyperclip提供中文复制
  def paste(self):
      # print(pyperclip.copy(self.writing))
      pyperclip.copy(self.writing)
      pyautogui.hotkey('ctrl', 'v')
      print('type words:',self.writing)

class getword():
  def __init__(self,image):
    
    """ 你的 APPID AK SK """
    self.APP_ID = '14662783'
    self.API_KEY = 'w5IwuG7Z4lBq0wQ9SH7TSuKh'
    self.SECRET_KEY = 'a9ov1phpjGRpOazWvS6DGPHiDPaBnBn0'
    self.image =image
  
  """ 读取图片 """
  def get_file_content(self,filePath):
      with open(filePath, 'rb') as fp:
          return fp.read()

  def getresultword(self):
    client = AipOcr(self.APP_ID, self.API_KEY, self.SECRET_KEY)
    #重新从文件读取图片 或者转换图片格式
    self.image = self.get_file_content('my_screenshot.png')
    
    """ 调用通用文字识别, 图片参数为本地图片 """
    words = client.basicGeneral(self.image)
    # TODO parse word
    print('get words: ' ,words )
    return words

    # return a 
    # """ 如果有可选参数 """
    # options = {}
    # options["language_type"] = "CHN_ENG"
    # options["detect_direction"] = "true"
    # options["detect_language"] = "true"
    # options["probability"] = "true"

    # """ 带参数调用通用文字识别, 图片参数为本地图片 """
    # client.basicGeneral(image, options)

    # url = "http//www.x.com/sample.jpg"

    # """ 调用通用文字识别, 图片参数为远程url图片 """
    # client.basicGeneralUrl(url);

    # """ 如果有可选参数 """
    # options = {}
    # options["language_type"] = "CHN_ENG"
    # options["detect_direction"] = "true"
    # options["detect_language"] = "true"
    # options["probability"] = "true"

    # """ 带参数调用通用文字识别, 图片参数为远程url图片 """
    # client.basicGeneralUrl(url, options)

class getcompanyfromExcel():
  def __init__():
    pass
  def getcompany(self):
    pass


class ExcelObject():
  #指定原始excel文件路径
  # excelfile='test.xls'
  #pcklfile='C:\\Users\\VI\\Desktop\\result.txt'

  def __init__(self,filepath):
    
    self.excelpath = filepath
    
    #读取excel
    #fiel_data = xlrd.open_workbook(excelfile)
    #formatting_info=True表示打开excel时并保存原有的格式
    self.fiel_data = xlrd.open_workbook(self.excelpath,formatting_info=True)

    #创建一个可以写入的副本
    #利用xlutils.copy函数，将xlrd.Book转为xlwt.Workbook，再用xlwt模块进行存储
    self.file_copy = copy(self.fiel_data)  


    #使用pandas库传入该excel的数值仅仅是为了后续判断插入数据时应插入行是哪行
    self.original_data = pd.read_excel(self.excelpath,encoding='utf-8')

    # Sheet_data = data.sheets()[0]          #通过索引顺序获取工作簿
    self.Sheet_data = self.fiel_data.sheet_by_index(0)    #通过索引顺序获取工作簿
    # self.Sheet_data = self.fiel_data.sheet_by_name(u'Sheet1')    #通过名称获取工作簿 通过xlrd的sheet_by_index()获取的sheet没有write()方法
  
    # self.Sheet_change = self.file_copy.get_sheet(sheetnum)#通过get_sheet()获取的sheet有write()方法
    # 按行列索引写入
    # Sheet_change.write(0,0,'Row 0, column 0 Value')

  #retain cell style
  def _getOutCell(self,writeSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    # self.setsheetchangebynum(sheetIndex)
    row = writeSheet._Worksheet__rows.get(rowIndex)
    # row = self.Sheet_change._Worksheet__rows.get(rowIndex)
    if not row: return None
    cell = row._Row__cells.get(colIndex)
    return cell

  #该函数中定义：对于没有任何修改的单元格，保持原有格式。
  def setOutCell(self,writeSheet, col, row, value):
    """ Change cell value without changing formatting. """

    # HACK to retain cell style.
    previousCell = self._getOutCell(writeSheet, col, row)
    # END HACK, PART I
    writeSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
      newCell = self._getOutCell(writeSheet, col, row)
      if newCell:
        newCell.xf_idx = previousCell.xf_idx

  def append(self,sheetIndex,str_write):
    """从末尾写入excel数据"""
    row=len(self.original_data)+1
    str_write="添加写入"+str_write
    writeSheet = self.file_copy.get_sheet(sheetIndex)
    self.setOutCell(writeSheet,0,row,str_write)
    self.file_copy.save(self.excelpath)
    print('write: ',str_write)
  
  def write(self,sheetIndex,row,col,value):
    """根据sheetIndex、row、col写入value"""
    writeSheet = self.file_copy.get_sheet(sheetIndex)
    self.setOutCell(writeSheet,col,row,value)
    # self.Sheet_change.write(row,col,value)
    self.file_copy.save(self.excelpath)

"""备忘录"""

# data = xlrd.open_workbook('excelfilepath') # 
# table = data.sheets()[0]
# table = data.sheet_by_index(0)
# table = data.sheet_by_name(u'Sheet1')
# table.row_values(i)
# table.col_values(i)
# nrows = table.nrows
# ncols = table.ncols
# sheetlist = table.sheets # A list of all sheets in the book.

# for i in range(nrows):
#   print(table.row_values(i))

# cell_A1 = table.cell(0,0).value
# cell_C4 = table.cell(2,3).value
"""


"""
# -*- coding: utf-8 -*-
# import xdrlib,sys
# import xlrd
# def open_excel(file='file.xls'):
# 	try:
# 		data = xlrd.open_workbook(file)
# 		return data
# 	except Exception,e:
		# print str(e)

"""根据索引获取Excel表格中的数据"""
#参数：file: Excel文件路径
#      colnameindex: 表头列名所在行的索引
#      by_index: 表的索引

# def excel_table_byindex(file='file.xls',colnameindex=0,by_index=0):
# 	data = open_excel(file)
# 	table = data.sheets()[by_index]
# 	nrows = table.nrows #行数
# 	ncols = table.ncols #列数
# 	colnames = table.row_values(colnameindex) #某一行数据
# 	list = []
# 	for rownum in range(1,nrows):
# 		row = table.row_values(rownum)#以列表格式输出
# 		if row:
# 			app = {}
# 			for i in range(len(colnames)):
# 				app[colnames[i]] = row[i]
# 			list.append(app)#向列表中插入字典类型的数据
# 	return list
 
# def main():
# 	tables = excel_table_byindex(file='test.xls')
# 	for row in tables:
# 		print row

# if __name__=="__main__":
# 	main()
""" """

"""通过名字索引"""
# # -*- coding: utf-8 -*-
# import xdrlib,sys
# import xlrd
# def open_excel(file='file.xls'):
# 	try:
# 		data = xlrd.open_workbook(file)
# 		return data
# 	except Exception,e:
# 		print str(e)

# def excel_table_byname(file='file.xls',colnameindex=0,by_name=u'Sheet1'):
# 	data = open_excel(file)
# 	table = data.sheet_by_name(by_name)
# 	nrows = table.nrows #行数
# 	colnames = table.row_values(colnameindex) #某一行数据
# 	list = []
# 	for rownum in range(1,nrows):
# 		row = table.row_values(rownum)
# 		if row:
# 			app = {}
# 			for i in range(len(colnames)):
# 				app[colnames[i]] = row[i]
# 			list.append(app)
# 	return list
 
# def main():
# 	tables = excel_table_byname(file='test.xls')
# 	for row in tables:
# 		print row

# if __name__=="__main__":
# 	main()

"""通过列表名索引"""
  # def getColumnIndex(table, columnName):
  #   columnIndex = None
  #     #print table
  #     for i in range(table.ncols):
  #       #print columnName
  #       #print table.cell_value(0, i)
  #       if(table.cell_value(0, i) == columnName):
  #         columnIndex = i
  #         break
  #   return columnIndex