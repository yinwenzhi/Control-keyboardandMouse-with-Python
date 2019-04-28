import pyautogui
import pyperclip
# import matplotlib
import pandas as pd
import time

class GainName():
    def __init__(self):
        # print('import gainName')
        # pyautogui.moveTo(70, 235) #把光标移动到(100, 200)位置
        pyautogui.click(220, 450)
        self.pageNow = 1
    def findreset(self,imgpath):
        imaloca = None
        times = 0
        pyautogui.scroll(1350)
        while not imaloca :
            times +=1
            try:
                print('正在查找重置按钮...')
                imaloca = pyautogui.locateOnScreen(imgpath)
                buttonx, buttony = pyautogui.center(reset)
            except:
                print('重置搜索信息失败，正在重试...')
                pyautogui.scroll(-1000)
            if times == 3:
                # pass
                return False
        return imaloca

    def findshown (self):
        isshown = None
        findshowntimes = 0
        while not isshown:
            findshowntimes +=1
            pyautogui.doubleClick(500, 70)
            isshown = pyautogui.locateOnScreen('engine\\anc\\概览.png')
            # isshown = pyautogui.locateOnScreen('engine\\anc\\机构信息.png')

            if findshowntimes == 5:
                break
        if isshown:
            print('详情页已出现，准备开始复制公司信息')
            return True 
        else:
            print('详情页未出现，查询失败')
            return False

    def findGoTo (self):
        isshown = None
        findshowntimes = 0

        while not isshown:
            pyautogui.click(723, 959)
            pyautogui.scroll(-800)
            print('scrolliingdown')
            findshowntimes +=1
            pyautogui.doubleClick(500, 70)
            # isshown = pyautogui.locateOnScreen('engine\\anc\\goto.png')
            isshown = pyautogui.locateOnScreen('engine\\anc\\goto.png')
            # if 5 > self.pageNow >=1 :
            if isshown != None:
                break

            isshown = pyautogui.locateOnScreen('engine\\anc\\goto5.png')
            # if 9 > self.pageNow >=5:
            if isshown != None:
                break

            isshown = pyautogui.locateOnScreen('engine\\anc\\goto9.png')
            if isshown != None:
            # if 35 > self.pageNow >=9:
                break
            
            isshown = pyautogui.locateOnScreen('engine\\anc\\goto10.png')
            # if  self.pageNow >=35:
            if isshown != None:
                break
            
            isshown = pyautogui.locateOnScreen('engine\\anc\\goto38.png')
            # if  self.pageNow >=35:
            if isshown != None:
                break
            
            isshown = pyautogui.locateOnScreen('engine\\anc\\goto39.png')
            # if  self.pageNow >=35:
            if isshown != None:
                break

            if findshowntimes == 3:
                break
        if isshown:
            print('公司列表页已出现')
            return isshown 
        else:
            print('公司列表页未出现，查询失败')
            return False

    def setDate(self,beginDate,endDate):
        pyautogui.click(220, 450)
        
        reset = self.findreset('engine\\anc\\重置.png')
        if reset:
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.click(buttonx, buttony)
            print('重置搜索')
            pyautogui.scroll(1500)
        else :
            return False
        isfindInput = 0
       
        while  not isfindInput:
            try:
                reset1 = pyautogui.locateOnScreen('engine\\anc\\输入.png')
                inputbuttonx, inputbuttony = pyautogui.center(reset1)
                isfindInput = 1 
            except:
                print('第一次查找输入框失败')
                try:
                    reset2 = pyautogui.locateOnScreen('engine\\anc\\输入3.png')
                    inputbuttonx, inputbuttony = pyautogui.center(reset2)
                except:
                    print('第二次查找输入框失败')
                    
            if isfindInput != 1:
                pyautogui.click(70, 235)
                print('打开折叠栏')
    
        manulx = 505
        manuly = 778
        # 是否点击自定义
        try:
            dateTo = pyautogui.locateOnScreen('engine\\anc\\dateTo.png')
            dateTox, dateToy = pyautogui.center(dateTo)
        except:
            pyautogui.click(manulx, manuly)
            print('点击自定义')
            dateTo = pyautogui.locateOnScreen('engine\\anc\\dateTo.png')
            dateTox, dateToy = pyautogui.center(dateTo)
        # 填入日期
        # print('begin:',beginDate)
        pyperclip.copy(beginDate)  #把text字符串中的字符复制到剪切板
        pyautogui.click(manulx+80, manuly)
        pyautogui.hotkey('ctrl', 'v')
        # print('endDate:',endDate)
        pyperclip.copy(endDate)  #把text字符串中的字符复制到剪切板
        pyautogui.click(manulx+250, manuly)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        try:
            reset=pyautogui.locateOnScreen('engine\\anc\\搜索.png')
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.moveTo(buttonx, buttony)
            pyautogui.click() #开始搜索
            # print('查找搜索按钮失败，正在重试...')
        except:
            reset=pyautogui.locateOnScreen('engine\\anc\\重置.png') #寻找重置按钮，以相对位置寻找搜索按钮
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.moveTo(buttonx-150, buttony) #开始搜索
            pyautogui.click() #开始搜索
        # pyautogui.PAUSE = 0.5
        print('开始搜索')
        pyautogui.scroll(-1500)
        return True

    def search( self,comnam ):

        pyautogui.click(220, 450)
        reset = None # 清空 reset

        """重置搜索信息"""
        # try:
        #     print('正在查找重置按钮...')
        #     pyautogui.scroll(1350)
        #     reset = pyautogui.locateOnScreen('engine\\anc\\重置.png')
        # except:
            
        #     print('重置搜索信息失败，正在重试...')
        #     pyautogui.click(70, 235)
        #     print('关闭折叠栏')
        #     reset = pyautogui.locateOnScreen('engine\\anc\\重置.png')
        # if reset :
        #     buttonx, buttony = pyautogui.center(reset)
        #     pyautogui.click(buttonx, buttony) #点击重置按钮
        #     print('重置搜索')
        # else :
        #     print('重置搜索信息失败')
        #     return False 

        reset = self.findreset('engine\\anc\\重置.png') 
        if reset :
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.click(buttonx, buttony) #点击重置按钮
            print('重置搜索')
        else :
            print('重置搜索信息失败')
            return False 

            # # 滚动按钮、
            # print('scrolling...')
            # pyautogui.scroll(1350)  # move mouse cursor to 100, 200, then scroll down 10 "clicks"
            # pyautogui.scroll(-800)
            # # pyautogui.click(670, 985) #重置搜索
            # reset=pyautogui.locateOnScreen('engine\\anc\\重置.png')
            # buttonx, buttony = pyautogui.center(reset)
            # pyautogui.click(buttonx, buttony) #重置搜索
            # print('重置搜索')

        # pyautogui.scroll(1350)  # move mouse cursor to 100, 200, then scroll down 10 "clicks"
        # pyautogui.moveTo(75, 240) #把光标移动到(100, 200)位置
        # pyautogui.click()

        pyautogui.scroll(1500)
        try:
            reset1 = pyautogui.locateOnScreen('engine\\anc\\输入.png')
            inputbuttonx, inputbuttony = pyautogui.center(reset1)
        except:
            print('第一次查找输入框失败')
            reset1 = 0 
            try:
                reset2 = pyautogui.locateOnScreen('engine\\anc\\输入3.png')
                inputbuttonx, inputbuttony = pyautogui.center(reset2)
            except:
                print('第二次查找输入框失败')
                reset2 = 0
        # try:
        #     reset3 = pyautogui.locateOnScreen('engine\\anc\\输入2.png')
        # except:
        #     reset3 = 0
      
        # if not (reset1 and reset2 ):
        if  (reset1 or reset2):
            # 任何一个元素找到就不进行单击打开机构信息折叠栏操作
            # pyautogui.moveTo(70, 235) #把光标移动到(100, 200)位置
            pass    
        else:
            
            pyautogui.click(70, 235)
            print('打开折叠栏')
            
        
        
        # pyautogui.click(inputbuttonx, inputbuttony)
        pyautogui.click(320, 288) # 点击输入框
        print('点击输入框')
        pyperclip.copy(comnam)  #把text字符串中的字符复制到剪切板
        pyautogui.hotkey('ctrl', 'v')
        #把光标移动到(100, 200)位置 并点击
        #   pyautogui.moveTo(320, 325)
        pyautogui.moveTo(320, 325)
        pyautogui.click() # 确认输入选择
        pyautogui.click()  #再次点击减少程序错误机率
        pyautogui.scroll(-300)
        # pyautogui.PAUSE = 1
        try:
            reset=pyautogui.locateOnScreen('engine\\anc\\搜索.png')
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.moveTo(buttonx, buttony)
            pyautogui.click() #开始搜索

            # print('查找搜索按钮失败，正在重试...')
        except:
            reset=pyautogui.locateOnScreen('engine\\anc\\重置.png') #寻找重置按钮，以相对位置寻找搜索按钮
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.moveTo(buttonx-150, buttony) #开始搜索
            pyautogui.click() #开始搜索
        # pyautogui.PAUSE = 0.5
        print('开始搜索')
        # 滚动按钮、
        # print('scrolling...')
        # pyautogui.scroll(-800, x=20, y=945)  # move mouse cursor to 100, 200, then scroll down 10 "clicks"
        
        try:
            reset=pyautogui.locateOnScreen('engine\\anc\\简称.png')
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.moveTo(buttonx-20, buttony+45)
            pyautogui.click()
            pyautogui.click(buttonx-20, buttony+50) #
        except:
            # print('查找公司连接失败，正在重试...') 
            pyautogui.scroll(100)
            pyautogui.scroll(-300)
            reset=pyautogui.locateOnScreen('engine\\anc\\类型.png')
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.moveTo(buttonx-220, buttony+48)
            pyautogui.click() #
            pyautogui.click(buttonx-220, buttony+55)
        print('查找公司连接成功，正在获取公司信息...')   
        

        # 把光标移动到(100, 200)位置 并点击
        # pyautogui.click(525, 990)
        # pyautogui.PAUSE = 0.5
        # time.sleep(1)

        if self.findshown():
                
            pyautogui.moveTo(8,90)
            pyautogui.click() #把光标移动到(100, 200)位置
            # 如果查找到“公司概览”说明信息页已经出现

            # 以拖拽方式复制公司信息
            pyautogui.dragTo(800,170, 1,button='left')#  按住鼠标左键，把鼠标拖拽到(100, 200)位置
            time.sleep(0.5)
            # 以shift方式
            # pyautogui.keyDown('shift')  # hold down the shift key
            # pyautogui.moveTo(5,90)
            # pyautogui.click()
            # pyautogui.keyUp('shift')    # release the shift key
            # 以拖拽方式复制公司信息
            # pyautogui.dragTo(800,170, 1,button='left')#  按住鼠标左键，把鼠标拖拽到(100, 200)位置
            pyautogui.hotkey('ctrl', 'c')
            
            totalmessage = pyperclip.paste()   #把剪切板上的字符串复制到text
            #   pyautogui.moveTo(530, 70) #关闭公司详情页面
            # print('pyperclip.paste()  成功')
            pyautogui.doubleClick(530, 70)

        else :
            pyautogui.doubleClick(530, 70)
            return False
        totalname =  totalmessage.strip().split('\r\n')

        print(totalname)
        # idxth = 
        for idx,string in enumerate(totalname):
            if string.find('公司') != -1:
                idxth = idx 
        chiname = totalname[0]
        engname = totalname[1]
                
        return chiname , engname

    def getNamebyLine(self,line):
        # print('start ge company')
        pyautogui.moveTo(723, 959)
        pyautogui.click()
        pyautogui.scroll(-1500)
        firstx = 20
        firsty = 180
        nowx = firstx
        nowy = firsty + (line-1)*32
        # print('nowy',nowy)
        pyautogui.moveTo(nowx, nowy) 
        pyautogui.click()
        # print('dianji:line',line)
        if self.findshown():
                
            # 如果查找到“公司概览”说明信息页已经出现
            pyautogui.moveTo(8,90)
            pyautogui.click() 
            print('开始复制公司信息')
            # 以拖拽方式复制公司信息
            pyautogui.dragTo(800,170, 1,button='left')#  按住鼠标左键，把鼠标拖拽到(100, 200)位置
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'c')
            
            totalmessage = pyperclip.paste()   #把剪切板上的字符串复制到text
            pyautogui.doubleClick(530, 70)

        else :
            pyautogui.doubleClick(530, 70)
            return False
        totalname =  totalmessage.strip().split('\r\n')

        print('totalnema: ',totalname)
        return totalname

    def byPage(self,row,rows):
        # pyautogui.click(220, 450)
        # 是否是对应页
        # print('bypage')
        row -=3
        p = row%25
        if  p == 0:
            page = row//25
        else:
            page = row//25+1
        # print('page:',page)
        self.page = page
        if page == self.pageNow:
            print('page =pageNow')
            pass
        else:    
            # 设置页数
            print('set page :',page)

            pyautogui.scroll(-500)
            # goto=pyautogui.locateOnScreen('engine\\anc\\goto.png') 
            goto = self.findGoTo()
            print('set page :',page)

            buttonx, buttony = pyautogui.center(goto)
            pyautogui.click(buttonx-30, buttony)
            # 跳转到page
            pyperclip.copy(page)  #把text字符串中的字符复制到剪切板
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.click(buttonx, buttony)
            self.pageNow = page
        line = row %25
        if line ==0:
            line =25
        # print('line:',line)
        # if self.findGoTo():
        print('第{}个公司，共{}个公司，对应第{}页第{}行'.format(row,rows-4,page,line))
        totalname = self.getNamebyLine(line)
        
    
        chiname = totalname[0]
        engname = totalname[1]
                
        return chiname , engname

    def restart(self):
        pyautogui.moveTo(930, 7)
        pyautogui.click() # 关闭软件
        time.sleep(1)
        pyautogui.moveTo(573, 310)
        pyautogui.click()
        pyautogui.moveTo(495, 265)
        pyautogui.doubleClick() # 双击打开软件
        time.sleep(5)
        pyautogui.moveTo(1230, 650)
        pyautogui.click() # 登录
        time.sleep(5)
        pyautogui.hotkey('win', 'left')
        time.sleep(1)
        pyautogui.moveTo(30, 35)
        pyautogui.click()  # 打开机构页
        time.sleep(2)