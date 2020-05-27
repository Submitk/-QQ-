# 原理是先将需要发送的文本放到剪贴板中，然后将剪贴板内容发送到qq窗口
# 之后模拟按键发送enter键发送消息
import win32api
import win32con
import requests
import json
import os
import random
import win32gui
import win32clipboard as w
import time
from PIL import Image
from ctypes import *

def get_msg():
    url = 'http://open.iciba.com/dsapi/'
    response = requests.post(url).text
    data = json.loads(response)
    img = data['fenxiang_img']
    name = str(data['dateline'])+'.jpg'
    with open(name,'wb') as f:
        f.write(requests.get(img, timeout=30).content)
    f.close()
    return name
def setImage(imgpath):
    im = Image.open(imgpath)
    im.save('1.bmp')
    aString = windll.user32.LoadImageW(0, r"1.bmp", win32con.IMAGE_BITMAP, 0, 0, win32con.LR_LOADFROMFILE)
 
    if aString != 0:  ## 由于图片编码问题  图片载入失败的话  aString 就等于0
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_BITMAP, aString)
        w.CloseClipboard()  
 
def getText():
    """获取剪贴板文本"""
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    return d

def setText(aString):
    """设置剪贴板文本"""
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
    w.CloseClipboard()

def send_qq(to_who,msg):
    """发送qq消息
    to_who：qq消息接收人
    msg：需要发送的消息
    """
    # 将消息写到剪贴板
    setText(msg)
#     sendByUser('poplar')
    
    # 获取qq窗口句柄
    qq = win32gui.FindWindow(None, to_who)
    # 投递剪贴板消息到QQ窗体
    win32gui.SendMessage(qq, 258, 22, 2080193)
    win32gui.SendMessage(qq, 770, 0, 0)
    # 模拟按下回车键
    win32gui.SendMessage(qq, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.SendMessage(qq, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    win32gui.SendMessage(qq, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
    win32api.keybd_event(13, 0, 0, 0)
    win32gui.SendMessage(qq, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
    win32gui.SetForegroundWindow(qq)
    
def send_qq_img(to_who,name):
    """发送qq消息
    to_who：qq消息接收人
    msg：需要发送的消息
    """
    
    # 将消息写到剪贴板
    setImage(name)
    # 获取qq窗口句柄
    qq = win32gui.FindWindow(None, to_who)
    # 投递剪贴板消息到QQ窗体
    win32gui.SendMessage(qq, 258, 22, 2080193)
    win32gui.SendMessage(qq, 770, 0, 0)
    # 模拟按下回车键
    win32gui.SendMessage(qq, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.SendMessage(qq, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    win32gui.SendMessage(qq, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
    win32api.keybd_event(13, 0, 0, 0)
    win32gui.SendMessage(qq, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
    win32gui.SetForegroundWindow(qq)

# 测试(๑•ᴗ•๑)♡
os.system('dd.bat')
time.sleep(3)
name = get_msg()
msg_list = {'1':'给公主大人请安｡:.ﾟヽ(｡◕‿◕｡)ﾉﾟ.:｡+ﾟ','2':'今天也要元气满满哦！( •̀∀•́ )','3':'撒浪嘿!o‿≖✧','4':'仙气十足小公主！٩(๑´0`๑)۶','5':'想念你(,,•́ . •̀,,)'}
msg = msg_list[str(random.randint(1,5))]
to_who='宇宙第一超级美少女'
send_qq(to_who,msg)
time.sleep(3)
send_qq_img(to_who,name)