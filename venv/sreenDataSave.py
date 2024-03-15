#!/usr/bin/python
import os
import time
import datetime

def mark_change(fileName, lists):
    #fht = open(fileNameHt, 'r', encoding='UTF-8')
    fp = open(fileName, 'r')
    lines = fp.readlines()
    num = len(lines)
    fp.close()
    
    group = (int)(num / 4)
    mark = lines[1:group]
    tip = lines[(group+1):2*group]
    tipword = lines[(2*group+1):3*group]
    time = lines[(3*group+1):4*group] 
    
    list = []
    for i in range(len(lists)):
        code = lists[i][0]
        if('6' == code[0]):
            list.append('01' + code + '=')
        else:
            list.append('00' + code + '=')
    
    for i in range(len(list)):
        for j in range(len(mark)):        
            code = mark[j][0:9]
            if(list[i] == code):
                mark.pop(j)
                break
    for i in range(len(list)):
        for j in range(len(tip)):        
            code = tip[j][0:9]
            if(list[i] == code):
                tip.pop(j)
                break            
    for i in range(len(list)):
        for j in range(len(tipword)):        
            code = tipword[j][0:9]
            if(list[i] == code):
                tipword.pop(j)
                break
    for i in range(len(list)):
        for j in range(len(time)):        
            code = time[j][0:9]
            if(list[i] == code):
                time.pop(j)
                break
    
    fp = open(fileName, "w")            
    fp.write('[MARK]\n')
    for i in range(len(list)):
        fp.write(list[i] + '7\n')
    for i in range(len(mark)):
        fp.write(mark[i])
    fp.write('[TIP]\n')
    for i in range(len(list)):
        continueDay = lists[i][5]
        fp.write(list[i] + continueDay + '\n')
    for i in range(len(tip)):
        fp.write(tip[i])
    fp.write('[TIPWORD]\n')
    for i in range(len(lists)):
        limitReason = lists[i][12]
        fp.write(list[i] + limitReason + '\n')
    for i in range(len(tipword)):
        fp.write(tipword[i])
    fp.write('[TIME]\n')
    for i in range(len(list)):
        firstLimitTime = lists[i][6]
        fTIme = str(int(firstLimitTime.replace(':', ''), 10))
        fp.write(list[i] + fTIme + '\n')
    for i in range(len(time)):
        fp.write(time[i])
    fp.close()
    
    return list
                    
def saveData(dlg, tabLeft, tabTop, infoLeft, infoTop, fileName):    
    downOrder = 6
    dlg.ClickInput(button=u'left', coords=(tabLeft, tabTop))
    #time.sleep(.1)
    #涨幅板块先按“涨幅”排序
    #if(fileName.find('zhangfu') >= 0):
        #dlg.ClickInput(button=u'left', coords=(540, 125))#涨幅
    #财务板块先按“买入信号”排序
    if(fileName.find('caiwu') >= 0):
        #dlg.ClickInput(button=u'left', coords=(920, 125))#买入信号
        dlg.ClickInput(button=u'left', coords=(740, 100))#买入信号
        downOrder = 6
        #time.sleep(.1)
    if(fileName.find('zhangting') >= 0):
        #dlg.ClickInput(button=u'left', coords=(1080, 125))#封成比
        downOrder = 6
        #time.sleep(.1)
    if(fileName.find('zijin') >= 0):
        downOrder = 6
    if(fileName.find('zhuli') >= 0):
        downOrder = 6
        #time.sleep(.1)
    dlg.ClickInput(button=u'right', coords=(infoLeft, infoTop))

    if(fileName.find('bk') >= 0):
        time.sleep(.5)
        downOrder = 6
    if(fileName.find('bkzj') >= 0):
        downOrder = 6
    if(fileName.find('bkzc') >= 0):
        downOrder = 6
    time.sleep(.1)
    k.tap_key(k.right_key)
    for down in range(downOrder):
        k.tap_key(k.up_key)    
    k.tap_key(k.right_key)
    k.tap_key(k.enter_key)

    dlgIO = app[u'导入导出对话框模板']

    k.type_string(fileName)

    dlgIO[u'下一步(N)'].ClickInput(button=u'left')
    #time.sleep(.1)
    dlgIO[u'下一步(N)'].CloseClick(button=u'left')
    #time.sleep(.1)
    #重复一次导出操作
    '''
    if(fileName.find('bk') < 0):
        dlgIO[u'取消'].CloseClick(button=u'left')
        dlg.ClickInput(button=u'right', coords=(infoLeft, infoTop))
        time.sleep(.1)
        k.tap_key(k.right_key)
        for down in range(downOrder):
            k.tap_key(k.down_key)    
        k.tap_key(k.right_key)
        k.tap_key(k.enter_key)

        dlgIO = app[u'导入导出对话框模板']

        k.type_string(fileName)

        dlgIO[u'下一步(N)'].ClickInput(button=u'left')
        #time.sleep(.1)
        dlgIO[u'下一步(N)'].CloseClick(button=u'left')
    '''
    
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


jingJiaReadyTime = 92630
currenTime = getCurrentTimeInt()
while(currenTime < jingJiaReadyTime):
    time.sleep(1)
    print('currenTime: ', currenTime)
    currenTime = getCurrentTimeInt()

main = 'D:/share/shareAnalyze/shareAnalyze/x64/Release/shareAnalyze.exe'
excelFile = 'C:/Users/Administrator/Documents/ZTFP.xlsx'

buyShare_file = 'D:/share/buyShare.txt'
if os.path.exists(buyShare_file): # 如果文件存在
    os.remove(buyShare_file) # 则删除

jingJiaOverTime = 92690
marketStartTime = 92690
marketCloseTime = 153100
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
dlg['Button32'].ClickInput(button=u'left')

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
zhangtingFile   = path + "zhangting_" + "THS_"   + tail

##################
topOffset   = 80#105
leftOffset  = 60
ddeOffset   = 160
zijinOffset = 260#340#260
zhuliOffset = 360#430#360
caiwuOffset = 440#520#440
zhangfuOffset = 520
zhangtingOffset = 130#160
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
zhangtingTabLeft   = left + zhangtingOffset;

infoleftOffset     = 400
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
    saveData(dlg, zhangtingTabLeft, tabTop, infoleftOffset, infoTopOffset, zhangtingFile)

if (int(index) > 1):
    #点击“板块”按钮
    dlg['Button31'].ClickInput(button=u'left')
    #save bankuai data
    bkzjOffset     = 200#260
    bkzcOffset     = 270#350
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



thsjl_right = ths_rectangle.right - ths_rectangle.left - 120
thsjl_bottom = bottom  = ths_rectangle.bottom - ths_rectangle.top - 10
dlg.ClickInput(button=u'left', coords=(thsjl_right, thsjl_bottom))
k.type_string(".001") 
k.tap_key(k.enter_key)

#调取大单净量分时
dlg.ClickInput(button=u'right', coords=(int(ths_rectangle.right/2), int(ths_rectangle.bottom/2+89)))
downOrder = 7
for down in range(downOrder):
    k.tap_key(k.up_key)
k.tap_key(k.enter_key)
dlg_XZZB = app[u'请选择指标']
for down in range(25):
    k.tap_key(k.down_key)
dlg_XZZB[u'确定'].ClickInput(button=u'left')



r_v = os.system(main)
print(r_v)

#程序下单
if(int(index) != 0):
    '''
    f1_TDX = open(buyShare_file,"r")  
    lines_TDX = f1_TDX.readlines()
    num_TDX = len(lines_TDX)
    code_TDX = lines_TDX[0][:-1]
    f1_TDX.close()
    '''
    print('GOOD LUCK!')
    '''
    #TDX
    app_TDX = Application().connect(path = r"D:\new_jyplug\TdxW.exe")
    dlg_TDX = app_TDX.window_(title_re = ".*通达信.*")
    dlg_TDX.ClickInput(button=u'left')
    rectangle = dlg_TDX.Rectangle()
    #选项菜单坐标
    xuanXiangLeft = rectangle.right - rectangle.left - 314 #150长江证券
    xuanXiangTop = 26
    
    #断开行情主站
    dlg_TDX.ClickInput(button=u'left', coords=(xuanXiangLeft, xuanXiangTop))    
    for down in range(4):
        k.tap_key(k.down_key)
    k.tap_key(k.enter_key)
    #连接行情主站
    dlg_TDX.ClickInput(button=u'left', coords=(xuanXiangLeft, xuanXiangTop))
    k.tap_key(k.down_key)
    k.tap_key(k.enter_key)
    k.tap_key(k.enter_key)
    '''
    
    
    
    
    #TDX
    #app_TDX = Application().connect(path = r"F:\new_tdx\Tdxw.exe")
    app_TDX = Application().connect(class_name = "TdxW_MainFrame_Class")
    dlg_TDX = app_TDX.window_(title_re = ".*通达信.*") #通达信 #金长江
    dlg_TDX.click_input(button=u'left')
    rectangle = dlg_TDX.Rectangle()
    
    
    '''
    if(getCurrentTimeInt() > marketCloseTime):
        #净买比
        jingMaiBiLeft = rectangle.right - rectangle.left - 1268
        jingMaiBiTop = 52
        dlg_TDX.click_input(button=u'left', coords=(jingMaiBiLeft, jingMaiBiTop))
        time.sleep(.2)
        dlg_TDX.click_input(button=u'left', coords=(jingMaiBiLeft, 200))
        time.sleep(.5)
        #shuju_TDX = app_TDX.window_(title_re = ".*数据导出.*")
        right = rectangle.right - rectangle.left - 100
        bottom  = rectangle.bottom - rectangle.top - 16
        dlg_TDX.click_input(button=u'left', coords=(right, bottom))
        time.sleep(.5)
        k.type_string("34")
        k.tap_key(k.enter_key)
        time.sleep(.5)
        
        #d:\share\zhangting_20240315.txt
        dlg_SJDC = dlg_TDX[u'数据导出']
        #dlg_SJDC.click_input(button=u'left', coords=(1, 1))
        time.sleep(.5)
        dlg_SJDC[u'所有数据(显示列开始所有栏目)'].click_input(button=u'left')
        time.sleep(.5)
        dlg_SJDC[u'导出'].click_input(button=u'left')
        #k.tap_key(k.enter_key)
        
        time.sleep(60)
        k.tap_key(k.enter_key)
    '''
    
    #选项菜单坐标
    xuanXiangLeft = rectangle.right - rectangle.left - 415 #470 #1980 #150长江证券 314金融终端 2250
    xuanXiangTop = 17
    
    #断开行情主站
    dlg_TDX.click_input(button=u'left', coords=(xuanXiangLeft, xuanXiangTop))   
    k.tap_key(k.down_key)
    #k.tap_key(k.enter_key)
    for down in range(3):
        k.tap_key(k.down_key)
    k.tap_key(k.enter_key)
    #连接行情主站
    dlg_TDX.click_input(button=u'left', coords=(xuanXiangLeft, xuanXiangTop))
    k.tap_key(k.down_key)
    #k.tap_key(k.enter_key)
    #for down in range(5):
    #    k.tap_key(k.down_key)
    k.tap_key(k.enter_key)
    k.tap_key(k.enter_key)
    
    '''
    #选项菜单坐标
    xuanXiangLeft = rectangle.right - rectangle.left - 2200 #1980 #150长江证券 314金融终端
    xuanXiangTop = 26
    
    #断开行情主站
    dlg_TDX.ClickInput(button=u'left', coords=(xuanXiangLeft, xuanXiangTop))   
    k.tap_key(k.down_key)
    #k.tap_key(k.enter_key)
    for down in range(3):
        k.tap_key(k.down_key)
    k.tap_key(k.enter_key)
    #连接行情主站
    dlg_TDX.ClickInput(button=u'left', coords=(xuanXiangLeft, xuanXiangTop))
    k.tap_key(k.down_key)
    #k.tap_key(k.enter_key)
    for down in range(0):
        k.tap_key(k.down_key)
    k.tap_key(k.enter_key)
    k.tap_key(k.enter_key)
    '''
    
    """
    time.sleep(1)
    
    #同步自选股
    tongBuLeft = 20
    tongBuTop  = rectangle.bottom - rectangle.top - 1260
    dlg_TDX.ClickInput(button=u'left', coords=(tongBuLeft, tongBuTop))
    tongBuLeft = rectangle.right - rectangle.left - 1210
    tongBuTop  = 1200
    dlg_TDX.ClickInput(button=u'left', coords=(tongBuLeft, tongBuTop))
    k.tap_key(k.enter_key)
    k.tap_key(k.enter_key)
    """
    
    
    '''
    if(num_TDX == 1) and (code_TDX != "9527"):        
        dlg_TDX.ClickInput(button=u'left')
        #k.tap_key(k.function_keys[6], 1)
        rectangle = dlg_TDX.Rectangle()
        right = rectangle.right - rectangle.left - 119
        bottom  = rectangle.bottom - rectangle.top - 16
        dlg_TDX.ClickInput(button=u'left', coords=(right, bottom))
        k.type_string(code_TDX)
        k.tap_key(k.enter_key)
        dlg_TDX.ClickInput(button=u'left', coords=((rectangle.right - rectangle.left - 400), 400))
        dlg_TDX.ClickInput(button=u'left', coords=(right, bottom))
        k.type_string("221")
        k.tap_key(k.enter_key)
        
        
        
        dlg_TDX[u'买入下单'].ClickInput(button=u'left')        
        k.tap_key(k.enter_key)
        #time.sleep(.1)
        k.tap_key(k.space_key)
        #time.sleep(.1)
        k.tap_key(k.space_key)
        
        
        
        
        #dlg_TDXIO = app_TDX[u'提示']
        #result = app_TDX[u'提示'].Wait("exists",1 ,1)
        #if(result == 1):
        #    dlg_TDXIO[u'确认'].CloseClick(button=u'left')

    #通达信添加 缠通套利 指标
    dlg_TDX.ClickInput(button=u'left')    
    dlg_TDX.ClickInput(button=u'right', coords=((rectangle.right - rectangle.left - 700), (rectangle.bottom - rectangle.top - 700)))
    k.tap_key(k.down_key)
    k.tap_key(k.down_key)         
    k.tap_key(k.right_key)
    k.tap_key(k.down_key)   
    k.tap_key(k.down_key)   
    k.tap_key(k.enter_key)
    #dlg_CTTL = app[u'请选择指标']
    for down in range(38):
        k.tap_key(k.down_key)
    k.tap_key(k.right_key)
    k.tap_key(k.down_key)   
    k.tap_key(k.down_key) 
    k.tap_key(k.enter_key)
    '''
#if(getCurrentTimeInt() < jingJiaOverTime+jingJiaTimeOffset):
#if (int(index) < 2):
#    saveData(dlg, zhuliTabLeft, tabTop, infoleftOffset, infoTopOffset, zhuliFile)

if(getCurrentTimeInt() > marketCloseTime):
    now = datetime.datetime.now()
    string = now.strftime('%Y%m%d')
    fileName = 'D:/share/zhangting_' + string + '.txt'
    fzhangting = open(fileName, "r")
    
    lines_zhangting = fzhangting.readlines()
    num = len(lines_zhangting)
    fzhangting.close()

    fzhangting = open(fileName, "w")  
    lineNum = (num - 1) / 2 - 1          
    fzhangting.write(str(int(lineNum)) + '\n')
    for i in range(int(num)):
        if(lines_zhangting[i] != '\n'):
            fzhangting.write(lines_zhangting[i][0:-1] + '\n')
    fzhangting.close()

    fzhangting = open(fileName, "r")
    lines_zhangting = fzhangting.readlines()
    num = len(lines_zhangting)
    fzhangting.close()
    
    lists = []
    i = 2
    num = num - 1
    while(i < num):
        list = lines_zhangting[i]
        list = list.split('\t')
        while '' in list:
            list.remove('')
        list[0] = list[0][2:9]
        lists.append(list)
        i = i + 1    
    mark_change('D:/Doc/Stock/TDX_KXG/T0002/mark.dat', lists)
    mark_change('F:/new_tdx/T0002/mark.dat', lists)
    
    modifyCfgFile(date, index)
    r_v = os.system(main)
    print(r_v)
    
'''
#点击“板块”按钮
dlg['Button6'].ClickInput(button=u'left')
#from scrapy.cmdline import execute
#execute()
'''




if(int(index) == 1):
    r_v = os.system(excelFile)
    print(r_v)
'''
time.sleep(10)
if(getCurrentTimeInt() < marketStartTime):
    r_v = os.system('C:\Windows\System32\shutdown /s /f /t 1559')
    print(r_v)
'''