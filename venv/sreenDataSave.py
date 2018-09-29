#!/usr/bin/python
import os
import time
import datetime
                    
def saveData(dlg, tabLeft, tabTop, infoLeft, infoTop, fileName):
    dlg.ClickInput(button=u'left', coords=(tabLeft, tabTop))
    time.sleep(.1)
    #财务板块先按“买入信号”排序
    if(fileName.find('caiwu') >= 0):
        dlg.ClickInput(button=u'left', coords=(920, 120))#买入信号
        time.sleep(.1)
    dlg.ClickInput(button=u'right', coords=(infoLeft, infoTop))

    downOrder = 10
    for down in range(downOrder):
        k.tap_key(k.down_key)    
    k.tap_key(k.right_key)
    k.tap_key(k.enter_key)

    dlgIO = app[u'导入导出对话框模板']

    k.type_string(fileName)

    dlgIO[u'下一步(N)'].ClickInput(button=u'left')
    time.sleep(.1)
    dlgIO[u'下一步(N)'].CloseClick(button=u'left')
    time.sleep(.1)
    dlgIO[u'完成'].Wait("enabled visible ready", 50, 3)
    dlgIO[u'完成'].CloseClick(button=u'left')

def getCurrentTimeInt():
    #获取当前时间
    now_time = datetime.datetime.now().strftime('%H%M%S')
    now_time_int = int(now_time)
    return now_time_int
    
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

jingJiaOverTime = 92500
marketCloseTime = 150010
jingJiaTimeOffset = 100

app = Application().connect(path = r"C:\同花顺软件\同花顺\hexin.exe")

f1 = open("D:/share/config.txt","r")  
lines = f1.readlines()
num = len(lines)
f1.close()

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


#ShowMenuAndControls(app);
#pywinauto.application.findwindows.enum_windows()

#app = Application().start("C:\同花顺软件\同花顺\hexin.exe")
#time.sleep(.5)
#dlg = app['同花顺(v8.70.35)']

dlg = app.window_(title_re = ".*同花顺.*")

#点击功能键F6
dlg.ClickInput(button=u'left')
time.sleep(.1)
k.tap_key(k.function_keys[6], 3)
time.sleep(.1)

#点击“个股”按钮
dlg['Button6'].ClickInput(button=u'left')

dlgFrame = dlg.AfxFrameOrView42s

rectangle = dlgFrame.Rectangle()

topOffset   = 105
leftOffset  = 60
ddeOffset   = 260
zijinOffset = 360
zhuliOffset = 460
caiwuOffset = 540
#left = rectangle.left
#top  = rectangle.top
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

saveData(dlg, zhangfuTabLeft, tabTop, infoleftOffset, infoTopOffset, zhangfuFile)
#if(getCurrentTimeInt() < jingJiaOverTime):
#    r_v = os.system(main)
#    break
#saveData(dlg, ddeTabLeft, tabTop, infoleftOffset, infoTopOffset, ddeFile)
saveData(dlg, zijinTabLeft, tabTop, infoleftOffset, infoTopOffset, zhijinFile)

if(getCurrentTimeInt() > jingJiaOverTime+jingJiaTimeOffset):
    saveData(dlg, zhuliTabLeft, tabTop, infoleftOffset, infoTopOffset, zhuliFile)

if(getCurrentTimeInt() > marketCloseTime):
    saveData(dlg, caiwuTabLeft, tabTop, infoleftOffset, infoTopOffset, caiwuFile)

if (int(index) > 1):
    #点击“板块”按钮
    dlg['Button8'].ClickInput(button=u'left')
    #save bankuai data
    bkzjOffset     = 160
    bkzcOffset     = 250
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
#from scrapy.cmdline import execute
#execute()

