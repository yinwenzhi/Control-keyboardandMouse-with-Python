import pyautogui
import pyperclip
# import matplotlib
import pandas as pd

class GainName():
    def __init__(self):
        print('import gainName')
        # pyautogui.moveTo(70, 235) #把光标移动到(100, 200)位置
        pyautogui.click(220, 450)

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
                return False
        return imaloca

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
        except:
            reset1 = 0 
        # try:
        #     reset2 = pyautogui.locateOnScreen('engine\\anc\\输入2.png')
        # except:
        #     reset2 = 0
        # try:
        #     reset3 = pyautogui.locateOnScreen('engine\\anc\\输入2.png')
        # except:
        #     reset3 = 0
      
        # if not (reset1 and reset2 ):
        if not (reset1):
            # 任何一个元素没有找到就单击打开机构信息折叠栏
            # pyautogui.moveTo(70, 235) #把光标移动到(100, 200)位置
            pyautogui.click(70, 235)
            print('打开折叠栏')
        
        pyautogui.click(320, 288) # 点击输入框

        pyperclip.copy(comnam)  #把text字符串中的字符复制到剪切板
        pyautogui.hotkey('ctrl', 'v')
        #把光标移动到(100, 200)位置 并点击
        #   pyautogui.moveTo(320, 325)
        pyautogui.moveTo(320, 325)
        pyautogui.click() # 确认输入选择

        pyautogui.scroll(-300)

        try:
            reset=pyautogui.locateOnScreen('engine\\anc\\搜索.png')
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.click(buttonx, buttony) #开始搜索

            print('查找搜索按钮失败，正在重试...')
        except:
            reset=pyautogui.locateOnScreen('engine\\anc\\重置.png') #寻找重置按钮，以相对位置寻找搜索按钮
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.click(buttonx-150, buttony) #开始搜索
            

        print('开始搜索')
        # 滚动按钮、
        print('scrolling...')
        pyautogui.scroll(-1350, x=20, y=945)  # move mouse cursor to 100, 200, then scroll down 10 "clicks"
        
        try:
            reset=pyautogui.locateOnScreen('engine\\anc\\简称.png')
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.click(buttonx-20, buttony+50) #
        except:
            print('查找公司连接失败，正在重试...')
            pyautogui.scroll(-100)
            reset=pyautogui.locateOnScreen('engine\\anc\\类型.png')
            buttonx, buttony = pyautogui.center(reset)
            pyautogui.click(buttonx-220, buttony+55) #
        print('查找公司连接成功，正在获取公司信息...')   
        

        # 把光标移动到(100, 200)位置 并点击
        # pyautogui.click(525, 990)
        # pyautogui.PAUSE = 0.5
        pyautogui.click(5, 90) #把光标移动到(100, 200)位置
        # 以拖拽方式复制公司信息
        pyautogui.dragTo(800,170, 1,button='left')#  按住鼠标左键，把鼠标拖拽到(100, 200)位置

        # 以shift方式
        # pyautogui.keyDown('shift')  # hold down the shift key
        # pyautogui.moveTo(5,90)
        # pyautogui.click()
        # pyautogui.keyUp('shift')    # release the shift key

        pyautogui.hotkey('ctrl', 'c')
        totalmessage = pyperclip.paste()   #把剪切板上的字符串复制到text
        #   pyautogui.moveTo(530, 70) #关闭公司详情页面
        pyautogui.doubleClick(530, 70)

        totalname =  totalmessage.strip().split('\r\n',3)

        print(totalname)
        # idxth = 
        for idx,string in enumerate(totalname):
            if string.find('公司') != -1:
                idxth = idx 
        chiname = totalname[0]
        engname = totalname[1]
                
        return chiname , engname
