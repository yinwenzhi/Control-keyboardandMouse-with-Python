import pyautogui
import pyperclip
import matplotlib
import pandas as pd

class GainName():
  def __init__(self):
    print('import gainName')
    pyautogui.moveTo(70, 235) #把光标移动到(100, 200)位置
    pyautogui.click()
  def search( self,comnam ):

      # 滚动按钮、
      print('scrolling...')
      pyautogui.scroll(1350)  # move mouse cursor to 100, 200, then scroll down 10 "clicks"
      pyautogui.scroll(-800)
      # pyautogui.click(670, 985) #重置搜索
      reset=pyautogui.locateOnScreen('engine\\anc\\重置.png')
      buttonx, buttony = pyautogui.center(reset)
      pyautogui.click(buttonx, buttony) #重置搜索
      print('重置搜索')

      pyautogui.scroll(1350)  # move mouse cursor to 100, 200, then scroll down 10 "clicks"
      # pyautogui.moveTo(75, 240) #把光标移动到(100, 200)位置
      # pyautogui.click()

      try:
          reset1 = pyautogui.locateOnScreen('engine\\anc\\输入.png')
      except:
          reset1 = 0 
      try:
          reset2 = pyautogui.locateOnScreen('engine\\anc\\输入.png')
      except:
          reset2 = 0
      try:
          reset3 = pyautogui.locateOnScreen('engine\\anc\\输入2.png')
      except:
          reset3 = 0
      
          # if reset:
          #     reset=pyautogui.locateOnScreen('engine\anc\输入.png')
      #     buttonx, buttony = pyautogui.center(reset)
      #     pyautogui.click(buttonx, buttony) #开始搜索
      #     print('查找搜索按钮失败，正在重试...')
      if not (reset1 and reset2):
          pyautogui.moveTo(70, 235) #把光标移动到(100, 200)位置
          pyautogui.click()
      #     reset=pyautogui.locateOnScreen('engine\\anc\\重置.png')
      #     buttonx, buttony = pyautogui.center(reset)
      #     pyautogui.click(buttonx-150, buttony) #开始搜索

      pyautogui.moveTo(320, 288) #把光标移动到(100, 200)位置
      pyautogui.click()

      pyperclip.copy(comnam)  #把text字符串中的字符复制到剪切板
      pyautogui.hotkey('ctrl', 'v')
      #把光标移动到(100, 200)位置 并点击
      pyautogui.moveTo(320, 325)
      pyautogui.click()

      pyautogui.scroll(-300)

      try:
          reset=pyautogui.locateOnScreen('engine\\anc\\搜索.png')
          buttonx, buttony = pyautogui.center(reset)
          pyautogui.click(buttonx, buttony) #开始搜索
          print('查找搜索按钮失败，正在重试...')
      except:
          reset=pyautogui.locateOnScreen('engine\\anc\\重置.png')
          buttonx, buttony = pyautogui.center(reset)
          pyautogui.click(buttonx-150, buttony) #开始搜索
      print('开始搜索')
      # 滚动按钮、
      print('scrolling...')
      pyautogui.scroll(-1350, x=20, y=945)  # move mouse cursor to 100, 200, then scroll down 10 "clicks"
    
      try:
          reset=pyautogui.locateOnScreen('engine\\anc\\机构简称列表.png')
          buttonx, buttony = pyautogui.center(reset)
          pyautogui.click(buttonx, buttony+55) #
          print('查找公司连接失败，正在重试...')
      except:
          reset=pyautogui.locateOnScreen('engine\\anc\\机构简称列表2.png')
          buttonx, buttony = pyautogui.center(reset)
          pyautogui.click(buttonx, buttony+55) #
      print('查找公司连接成功，正在获取公司信息...')   
      

      # 把光标移动到(100, 200)位置 并点击
      pyautogui.click(525, 990)

      # 复制公司信息
      pyautogui.moveTo(900, 200) #把光标移动到(100, 200)位置
      #  按住鼠标左键，把鼠标拖拽到(100, 200)位置
      pyautogui.dragTo(5,120, 0.5,button='left')
      pyautogui.hotkey('ctrl', 'c')
      totalmessage = pyperclip.paste()   #把剪切板上的字符串复制到text
      pyautogui.moveTo(530, 70) #关闭公司详情页面
      pyautogui.doubleClick()

      totalname =  totalmessage.split('\r\n',3)

      print(totalname)
      # idxth = 
      for idx,string in enumerate(totalname):
          if string.find('公司') != -1:
              idxth = idx 
      engname = totalname[2]
      chiname = totalname[1]
              
      return chiname , engname
