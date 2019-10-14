#!/usr/bin/python
import os
import time
import datetime
                    
def saveData(dlg, tabLeft, tabTop, infoLeft, infoTop, fileName):
    dlg.ClickInput(button=u'left', coords=(tabLeft, tabTop))
    #time.sleep(.1)
    #涨幅板块先按“涨幅”排序
    #if(fileName.find('zhangfu') >= 0):
        #dlg.ClickInput(button=u'left', coords=(540, 125))#涨幅
    #财务板块先按“买入信号”排序
    if(fileName.find('caiwu') >= 0):
        dlg.ClickInput(button=u'left', coords=(920, 125))#买入信号
        #time.sleep(.1)
    dlg.ClickInput(button=u'right', coords=(infoLeft, infoTop))

    downOrder = 10
    if(fileName.find('bk') >= 0):
        downOrder = 9
    for down in range(downOrder):
        k.tap_key(k.down_key)    
    k.tap_key(k.right_key)
    k.tap_key(k.enter_key)

    dlgIO = app[u'导入导出对话框模板']

    k.type_string(fileName)

    dlgIO[u'下一步(N)'].ClickInput(button=u'left')
    #time.sleep(.1)
    dlgIO[u'下一步(N)'].CloseClick(button=u'left')
    #time.sleep(.1)
    #重复一次导出操作
    if(fileName.find('bk') < 0):
        dlgIO[u'取消'].CloseClick(button=u'left')
        dlg.ClickInput(button=u'right', coords=(infoLeft, infoTop))

        for down in range(downOrder):
            k.tap_key(k.down_key)    
        k.tap_key(k.right_key)
        k.tap_key(k.enter_key)

        dlgIO = app[u'导入导出对话框模板']

        k.type_string(fileName)

        dlgIO[u'下一步(N)'].ClickInput(button=u'left')
        #time.sleep(.1)
        dlgIO[u'下一步(N)'].CloseClick(button=u'left')
    
    
    dlgIO[u'完成'].Wait("enabled visible ready", 50, 3)
    dlgIO[u'完成'].CloseClick(button=u'left')

def getCurrentTimeInt():
    #获取当前时间
    now_time = datetime.datetime.now().strftime('%H%M%S')
    now_time_int = int(now_time)
    return now_time_int
    
def modifyCfgFile(date, index):
    f2 = open("D:/share/config.txt","w")
    tomorrow = str(int(date.strip('_'), 10) + 1) + '_\n'
    lines[0] = tomorrow
    lines[1] = '0\n'
    lines[12] = date + index + '\n'
    f2.writelines(lines)
    f2.close()
    
from pywinauto.application import Application
from pywinauto import clipboard
import win32api
import win32con
import os, sys, time
from pymouse import PyMouse
from pykeyboard import PyKeyboard
m = PyMouse()
k = PyKeyboard()

main = 'D:/share/shareAnalyze/shareAnalyze/x64/Release/shareAnalyze.exe'

buyShare_file = 'D:/share/buyShare.txt'
if os.path.exists(buyShare_file): # 如果文件存在
    os.remove(buyShare_file) # 则删除

jingJiaOverTime = 92500
marketStartTime = 93000
marketCloseTime = 150010
jingJiaTimeOffset = 5

app = Application().connect(path = r"C:\同花顺软件\同花顺\hexin.exe")

f1 = open("D:/share/config.txt","r")  
lines = f1.readlines()
num = len(lines)
f1.close()

#ShowMenuAndControls(app);
#pywinauto.application.findwindows.enum_windows()

#app = Application().start("C:\同花顺软件\同花顺\hexin.exe")
#time.sleep(.5)
#dlg = app['同花顺(v8.70.35)']

dlg = app.window_(title_re = ".*同花顺.*")

#点击功能键F6
dlg.ClickInput(button=u'left')
#time.sleep(.1)
k.tap_key(k.function_keys[6], 1)
#time.sleep(.1)

#点击“个股”按钮
dlg['Button7'].ClickInput(button=u'left')

dlgFrame = dlg.AfxFrameOrView42s

ths_rectangle = dlgFrame.Rectangle()

##############
date  = lines[0].strip()
#index = str(int(lines[1].strip(), 10)+1)
index = str(int(lines[1].strip(), 10))
if(getCurrentTimeInt() > jingJiaOverTime):
    index = str(int(index) + 1);
path  = lines[2].strip()
tail  = date + index + ".txt"
zhangfuFile = path + "zhangfu_" + tail
ddeFile     = path + "DDE_"     + tail
zhijinFile  = path + "zijin_"   + tail
zhuliFile   = path + "zhuli_"   + tail
bkrdFile    = path + "bkrd_"    + tail
bkzjFile    = path + "bkzj_"    + tail
bkzcFile    = path + "bkzc_"    + tail
caiwuFile   = path + "caiwu_"   + tail

##################
topOffset   = 105
leftOffset  = 60
ddeOffset   = 160
zijinOffset = 260
zhuliOffset = 360
caiwuOffset = 440
zhangfuOffset = 520
#left = ths_rectangle.left
#top  = ths_rectangle.top
left = 0
top  = 0

tabTop         = top  + topOffset;
zhangfuTabLeft = left + leftOffset;
ddeTabLeft     = left + ddeOffset;
zijinTabLeft   = left + zijinOffset;
zhuliTabLeft   = left + zhuliOffset;
caiwuTabLeft   = left + caiwuOffset;

infoleftOffset     = 200
infoTopOffset      = 250

curTime = getCurrentTimeInt()
cycleNum = 1
if(curTime > jingJiaOverTime) and (curTime < marketStartTime):
    cycleNum = 1

for i in range(cycleNum):
    saveData(dlg, zhangfuTabLeft, tabTop, infoleftOffset, infoTopOffset, zhangfuFile)
    #if(getCurrentTimeInt() < jingJiaOverTime):
    #    r_v = os.system(main)
    #    break
    #saveData(dlg, ddeTabLeft, tabTop, infoleftOffset, infoTopOffset, ddeFile)
    saveData(dlg, zijinTabLeft, tabTop, infoleftOffset, infoTopOffset, zhijinFile)

#if(getCurrentTimeInt() > jingJiaOverTime+jingJiaTimeOffset):
if (int(index) > 1):
    saveData(dlg, zhuliTabLeft, tabTop, infoleftOffset, infoTopOffset, zhuliFile)

if(getCurrentTimeInt() > marketCloseTime):
    saveData(dlg, caiwuTabLeft, tabTop, infoleftOffset, infoTopOffset, caiwuFile)

if (int(index) > 1):
    #点击“板块”按钮
    dlg['Button6'].ClickInput(button=u'left')
    #save bankuai data
    bkzjOffset     = 260
    bkzcOffset     = 350
    bkrdTabLeft    = leftOffset
    bkzjTabLeft    = bkzjOffset
    bkzcTabLeft    = bkzcOffset
    saveData(dlg, bkrdTabLeft, tabTop, infoleftOffset, infoTopOffset, bkrdFile)
    saveData(dlg, bkzjTabLeft, tabTop, infoleftOffset, infoTopOffset, bkzjFile)
    saveData(dlg, bkzcTabLeft, tabTop, infoleftOffset, infoTopOffset, bkzcFile)

if(getCurrentTimeInt() > jingJiaOverTime+jingJiaTimeOffset):
    #modify config.txt file, add index 1
    f2 = open("D:/share/config.txt","w")
    lines[1] = index + '\n'
    f2.writelines(lines)
    f2.close()

r_v = os.system(main)
print(r_v)

#程序下单
if(int(index) == 1):
    f1_haiTong = open(buyShare_file,"r")  
    lines_haiTong = f1_haiTong.readlines()
    num_haiTong = len(lines_haiTong)
    code_haiTong = lines_haiTong[0][:-1]
    f1_haiTong.close()

    if(num_haiTong == 1) and (code_haiTong != "9527"):
        #haiTong
        app_haiTong = Application().connect(path = r"C:\new_haitong\TdxW.exe")
        dlg_haiTong = app_haiTong.window_(title_re = ".*海通.*")
        dlg_haiTong.ClickInput(button=u'left')
        #k.tap_key(k.function_keys[6], 1)
        rectangle = dlg_haiTong.Rectangle()
        right = rectangle.right - rectangle.left - 119
        bottom  = rectangle.bottom - rectangle.top - 12
        dlg_haiTong.ClickInput(button=u'left', coords=(right, bottom))
        k.type_string(code_haiTong)
        k.tap_key(k.enter_key)
        dlg_haiTong.ClickInput(button=u'left', coords=((rectangle.right - rectangle.left - 400), 400))
        dlg_haiTong.ClickInput(button=u'left', coords=(right, bottom))
        k.type_string("221")
        k.tap_key(k.enter_key)
        '''
        dlg_haiTong[u'买入下单'].ClickInput(button=u'left')        
        k.tap_key(k.enter_key)
        #time.sleep(.1)
        k.tap_key(k.space_key)
        #time.sleep(.1)
        k.tap_key(k.space_key)
        '''
        #dlg_haiTongIO = app_haiTong[u'提示']
        #result = app_haiTong[u'提示'].Wait("exists",1 ,1)
        #if(result == 1):
        #    dlg_haiTongIO[u'确认'].CloseClick(button=u'left')

#if(getCurrentTimeInt() < jingJiaOverTime+jingJiaTimeOffset):
if (int(index) < 2):
    saveData(dlg, zhuliTabLeft, tabTop, infoleftOffset, infoTopOffset, zhuliFile)

if(getCurrentTimeInt() > marketCloseTime):
    modifyCfgFile(date, index)
    
'''
#点击“板块”按钮
dlg['Button6'].ClickInput(button=u'left')
#from scrapy.cmdline import execute
#execute()
'''

thsjl_right = ths_rectangle.right - ths_rectangle.left - 120
thsjl_bottom = bottom  = ths_rectangle.bottom - ths_rectangle.top - 10
dlg.ClickInput(button=u'left', coords=(thsjl_right, thsjl_bottom))
k.type_string(".001") 
k.tap_key(k.enter_key)


