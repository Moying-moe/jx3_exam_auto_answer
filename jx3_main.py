import time
import datetime
import hashlib
import random
import base64
import os.path
import urllib.request
import urllib.parse
import sys
import json
import traceback
import math

from PIL import ImageGrab
from PIL import Image
from io import BytesIO

import pickle as pkl

from win32api import GetSystemMetrics

import win32con
import ctypes
import ctypes.wintypes
import threading

import difflib

import tkinter

#############################################################
# set appid and appkey here
# the value below is not real. please edit it if you want to run this app
APPID = 0
APPKEY = ''
#############################################################


def nonce_str():
    randlist = ['0','1','2','3','4','5','6','7','8','9','a','b','c','d','e','f']
    t = ''
    for i in range(16):
        t += randlist[random.randint(0,15)]
    return t

'''shutlist = ['，','。','《','》','、','？','！','【','】','“','”','…',
            '（','）','；','：',',','.','<','>','\\','/','?','!','[',']',
            '"','(',')',';',':','*',' ']'''
def fuzzyfinder(user_input, collection):
    ranklist = []
    for each in collection:
        apd = each,difflib.SequenceMatcher(a=user_input,b=each).quick_ratio()
        ranklist.append(apd)
    ranklist.sort(key = lambda x:x[1],reverse = True)
    return ranklist[0]

user32 = ctypes.windll.user32  #加载user32.dll
id1 = 105 #注册热键的唯一id，用来区分热键
id2 = 106

class Hotkey(threading.Thread):  #创建一个Thread.threading的扩展类  

    def run(self):
        user32.UnregisterHotKey(None, id1)
        user32.UnregisterHotKey(None, id2)
        if not user32.RegisterHotKey(None, id1, 0, win32con.VK_F9):   # 注册快捷键F9并判断是否成功，该热键用于执行一次需要执行的内容。
            showerror("热键创建失败!", "F9热键创建失败!请确定没有其他程序占用这个热键")
            print("Unable to register id", id1) # 返回一个错误信息
            with open('error%s.log'%str(datetime.datetime.fromtimestamp(time.time())).split('.')[0].replace(':','-'),'w') as f:
                f.write('F9热键创建错误\n' + traceback.format_exc())
            window.destroy()
            if IS_SHOWN:
                root.destroy()
            sys.exit()
        if not user32.RegisterHotKey(None, id2, 0, win32con.VK_F10):   # 注册快捷键F9并判断是否成功，该热键用于执行一次需要执行的内容。
            showerror("热键创建失败!", "F10热键创建失败!请确定没有其他程序占用这个热键")
            print("Unable to register id", id2) # 返回一个错误信息
            with open('error%s.log'%str(datetime.datetime.fromtimestamp(time.time())).split('.')[0].replace(':','-'),'w') as f:
                f.write('F10热键创建错误\n' + traceback.format_exc())
            window.destroy()
            if IS_SHOWN:
                root.destroy()
            sys.exit()

        #以下为检测热键是否被按下，并在最后释放快捷键  
        try:  
            msg = ctypes.wintypes.MSG()  

            while True:
                if user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:

                    if msg.message == win32con.WM_HOTKEY:  
                        if msg.wParam == id1:
                            global SCANNING,nowpic,lb
                            lb['text'] = '正在识别 请稍候...'
                            nowpic = ImageGrab.grab(bbox=(posx1, posy1, posx2, posy2))
                            SCANNING = True
                            is_change()
                            hotkeyF9()
                        elif msg.wParam == id2:
                            btnexit()

                    user32.TranslateMessage(ctypes.byref(msg))  
                    user32.DispatchMessageA(ctypes.byref(msg))

        finally:
            user32.UnregisterHotKey(None, id1)
            user32.UnregisterHotKey(None, id2)

def f1():
    global lb
    lb['text'] = '识别中 请稍候...'
    hotkeyF9()


def is_change():
    global root,nowpic
    img = ImageGrab.grab(bbox=(posx1, posy1, posx2, posy2))
    if img != nowpic:
        global lb
        nowpic = img
        t = threading.Thread(target=f1)
        t.setDaemon(True)
        t.start()
    if SCANNING:
        root.after(50,is_change)
    

def hotkeyF9():
    global lb,SCANNING
    #lb['text'] = '正在识别 请稍候...'
    img = ImageGrab.grab(bbox=(posx1, posy1, posx2, posy2))
    output_buffer = BytesIO()
    img.save(output_buffer, format = 'JPEG')
    byte_data = output_buffer.getvalue()
    grab_pic = base64.b64encode(byte_data)

    temp = [('app_id',APPID),
            ('time_stamp',int(time.time())),
            ('nonce_str',nonce_str()),
            ('image',grab_pic)]
    temp.sort()
    temp.append(('app_key',APPKEY))
    temp = dict(temp)
    calc = urllib.parse.urlencode(temp)
    apikey = hashlib.md5(calc.encode('utf-8')).hexdigest().upper()
    del temp['app_key']
    temp['sign'] = apikey

    data = urllib.parse.urlencode(temp).encode('utf-8')
    url = r'https://api.ai.qq.com/fcgi-bin/ocr/ocr_generalocr'
    req = urllib.request.Request(url,data)
    response = urllib.request.urlopen(req)
    result = response.read().decode('utf-8')
    temp = json.loads(result)

    if temp['ret'] != 0:
        print('错误！可能是截图出错或者响应时间超过五秒。')
        print(str(temp))
        lb['text'] = '出错！可能是同一时间使用者太多。按F9重新开始，按F10退出。'
        SCANNING = False
    else:
        text = ''
        for each in temp['data']['item_list']:
            text += each['itemstring']
        text = text.replace('单选题:','').replace('单选题：','')
        question = fuzzyfinder(text, timu)
        quest = question[0]
        print('答案：',tiku[quest])
        if question[1] < 0.4:
            lb['text'] = tiku[quest]+'(可信度很低 可能出现了错误 已停止自动查题 按F9继续 按F10退出)'
            SCANNING = False
        else:
            lb['text'] = tiku[quest]+'(按F10退出)'



def btn():
    global window
    global root,lb
    global IS_SHOWN
    IS_SHOWN = True
    window.state('iconic')
    wbtn['state'] = 'disabled'
    
    hotkey = Hotkey()  
    hotkey.start()
    root = tkinter.Tk()
    root.wm_attributes('-topmost',1)
    if BG:
        root.wm_attributes("-transparentcolor","red")
    else:
        root.wm_attributes("-transparentcolor","white")
    root.overrideredirect(True)
    root["background"] = "white"
    root.geometry("+%s+%s"%(showx,showy))
    lb = tkinter.Label(root, text = '欢迎使用剑网三智能答题器 按F9开始查题 按F10退出',bg='white')
    lb.pack()
    root.mainloop()

def btnexit():
    global window
    if IS_SHOWN:
        global root
        root.destroy()
    user32.UnregisterHotKey(None, id1)
    user32.UnregisterHotKey(None, id2)
    window.destroy()
    sys.exit()

def answerbg(_):
    global BG
    BG = not BG
    with open('set.cfg','wb') as f:
        pkl.dump(BG,f)
        pkl.dump(float(tk_ratio.get()),f)
    if IS_SHOWN:
        global root
        if BG:
            root.wm_attributes("-transparentcolor","red")
        else:
            root.wm_attributes("-transparentcolor","white")

def change_ratio(_):
    global posx1, posy1, posx2, posy2, showx, showy, root
    ratio = float(tk_ratio.get())
    posx1 = int(GetSystemMetrics(0) * 0.040625 * ratio)
    posy1 = int(GetSystemMetrics(1) * 0.25 * ratio)
    posx2 = int(GetSystemMetrics(0) * 0.410625 * ratio)
    posy2 = int(GetSystemMetrics(1) * 0.32 * ratio)
    showx = int(GetSystemMetrics(0) * 0.0625 * ratio)
    showy = int(GetSystemMetrics(1) * (290/900) * ratio)
    root.geometry("+%s+%s"%(showx,showy))
    with open('set.cfg','wb') as f:
        pkl.dump(BG,f)
        pkl.dump(float(tk_ratio.get()),f)

window = tkinter.Tk()
try:
    with open('set.cfg','rb') as f:
        BG = pkl.load(f)
        tk_ratio = tkinter.StringVar()
        tk_ratio.set(pkl.load(f))
except:
    print('配置文件错误')
    with open('error%s.log'%str(datetime.datetime.fromtimestamp(time.time())).split('.')[0].replace(':','-'),'w') as f:
        f.write('配置文件错误\n' + traceback.format_exc())
    sys.exit()
IS_SHOWN = False
SCANNING = False
window.protocol("WM_DELETE_WINDOW", btnexit)
window.title('剑网三智能答题器')
wlb = tkinter.Label(window, text = '\n\n欢迎使用剑网三智能答题器')
wbtn = tkinter.Button(window, text = '开始查题',command = btn,state = 'disabled')
wbtn2 = tkinter.Button(window, text = '退出', command = btnexit)
wcbtn = tkinter.Checkbutton(window, text = '答案显示背景(如果看不清答案请选择)')
if BG:
    wcbtn.select()
wscl = tkinter.Scale(window, from_ = 0, to = 1, resolution = 0.01,
                     orient = tkinter.HORIZONTAL,variable = tk_ratio)
wscl.set(float(tk_ratio.get()))
wlb2 = tkinter.Label(window, text = '窗口比例:')
wscl.bind('<ButtonRelease-1>',change_ratio)
wcbtn.bind('<Button-1>',answerbg)
wlb.grid(row = 0, column = 0, columnspan = 2)
wbtn.grid(row = 1, column = 0, columnspan = 2)
wbtn2.grid(row = 2, column = 0, columnspan = 2)
wcbtn.grid(row = 3, column = 0, columnspan = 2)
wlb2.grid(row = 4, column = 0)
wscl.grid(row = 4, sticky = tkinter.W+tkinter.E, column = 1)


haoshi = time.time()
print('欢迎使用剑网三智能答题器')
try:
    with open('data.pkl','rb') as f:
        tiku = pkl.load(f)
    tiku = dict(tiku)
    timu = [] # 用于模糊搜索
    for each in tiku:
        timu.append(each)
except:
    wlb['text'] = '''\n欢迎使用剑网三智能答题器
题库加载失败 请检查目录下"data.pkl"文件是否损坏 或重新打开/下载插件'''
    print('题库加载失败 请检查目录下"data.pkl"文件是否损坏 或重新打开/下载插件')
    with open('error%s.log'%str(datetime.datetime.fromtimestamp(time.time())).split('.')[0].replace(':','-'),'w') as f:
        f.write('题库加载错误\n' + traceback.format_exc())
    window.destroy()
    if IS_SHOWN:
        root.destroy()
    sys.exit()
haoshi = time.time() - haoshi
wlb['text'] = '''题库加载成功欢迎使用剑网三智能答题器
题库加载成功 耗时%.3f秒
请打开科举界面 点击"开始查题"按钮'''%haoshi
print('题库加载成功')
print('请打开科举界面 点击"开始查题"按钮')
wbtn['state'] = 'normal'

ratio = float(tk_ratio.get())
posx1 = int(GetSystemMetrics(0) * 0.040625 * ratio)
posy1 = int(GetSystemMetrics(1) * 0.25 * ratio)
posx2 = int(GetSystemMetrics(0) * 0.410625 * ratio)
posy2 = int(GetSystemMetrics(1) * 0.32 * ratio)
showx = int(GetSystemMetrics(0) * 0.05625 * ratio)
showy = int(GetSystemMetrics(1) * 0.5 * ratio)

window.mainloop()
