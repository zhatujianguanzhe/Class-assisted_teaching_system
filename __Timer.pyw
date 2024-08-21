# -*- coding: utf-8 -*-
from tkinter import *
import tkinter.filedialog
from tkinter.ttk import *
import tkinter
import win32api
import ctypes,os,sys,configparser,re,time

if getattr(sys, 'frozen', None):
    res_icon_folder = str(sys._MEIPASS)+'\\icons'
else:
    res_icon_folder = str(os.path.dirname(__file__))+'\\icons'

def remove_BOM(config_path): #去掉配置文件开头的BOM字节
    content = open(config_path,encoding='utf-8').read()
    try:
        content = re.sub(r"\xfe\xff","", content)
        content = re.sub(r"\xff\xfe","", content)
        content = re.sub(r"\ufeff","", content)
        content = re.sub(r"\xef\xbb\xbf","", content)
    except:
        pass
    open(config_path, 'r+',encoding='utf-8').write(content)

try:
    remove_BOM('settings.ini')
    settings = configparser.ConfigParser()
    settings.read('settings.ini',encoding='utf-8')
    ctypes.windll.shcore.SetProcessDpiAwareness(int(settings.get('DPI','DPI_mode')))
except:
    if win32api.GetSystemMetrics(0)>1920:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(0)
        except:
            pass
    else:#<1920
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass







def close_root_Button_Timer_window():
    root_Button_Timer_window.destroy()
    sys.exit()
def ok_(x):
    ok()
def clock():
    #time_clock分钟,要先进行转换
    global time_clock,time_text
    try:
        def convert_minutes_to_hours_and_minutes(total_minutes):
            hours = total_minutes // 60
            minutes = total_minutes % 60
            return [hours, minutes]
        time_clock_temp=time_clock

        for i in range(time_clock):
            if time_clock_temp<=30 and time_clock_temp>10:
                time_text['fg']='blue'
                time_text['bg']='white'
            elif time_clock_temp<=10 and time_clock_temp>5:
                time_text['fg']='yellow'
                time_text['bg']='black'
            elif time_clock_temp<=5 and time_clock_temp>3:
                time_text['fg']='blue'
                time_text['bg']='red'
            elif time_clock_temp<=3 and time_clock_temp>0:
                time_text['fg']='yellow'
                time_text['bg']='red'
            StandardTime=convert_minutes_to_hours_and_minutes(time_clock_temp)
            time_text['text']='%s:%s'%(str(StandardTime[0]),str(StandardTime[1]).zfill(2))
            root_Button_Timer_window.update()
            for i in range(60):#60
                root_Button_Timer_window.update()
                time.sleep(1)
            time_clock_temp=time_clock_temp-1

        time_text['fg']='black'
        time_text['bg']='green'
        time_text['text']='考试结束'
        root_Button_Timer_window.update()
        win32api.Beep(2000,500);win32api.Beep(2000,500);win32api.Beep(2000,500);win32api.Beep(2000,500);win32api.Beep(2000,500)
    except:
        try:
            time_text['text']='错误'
        except:
            pass
        return
def ok():
    global time_clock,time_text
    try:
        time_clock=int(time_entry.get())
    except:
        win32api.MessageBeep()
        return  0
    if time_entry.get().replace(' ','')=='':
        win32api.MessageBeep()
        return  0
    lb.destroy()
    time_entry.destroy()
    ok_button.destroy()
    time_text=tkinter.Label(root_Button_Timer_window,font=('黑体',128),justify='center',cursor='arrow')
    time_text.pack(fill='both',expand=True)

    clock()
root_Button_Timer_window=Tk()
root_Button_Timer_window.title("计时器")
# 设置窗口大小、居中
width = 1000
height = 650
screenwidth = root_Button_Timer_window.winfo_screenwidth()
screenheight = root_Button_Timer_window.winfo_screenheight()
geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root_Button_Timer_window.geometry(geometry)
root_Button_Timer_window.focus_set()
root_Button_Timer_window.protocol("WM_DELETE_WINDOW", close_root_Button_Timer_window)
root_Button_Timer_window.bind('<Return>',ok_)
root_Button_Timer_window.attributes('-topmost', 'true')
root_Button_Timer_window.iconbitmap(str(res_icon_folder)+'/icon.ico')


lb=tkinter.Label(root_Button_Timer_window,text='时间(分钟):',anchor='w')
lb.place(x=20,y=20,width=130,height=40)

time_entry=Entry(root_Button_Timer_window,)
time_entry.place(x=170,y=20,width=150,height=40,)
time_entry.focus()

ok_button=Button(root_Button_Timer_window,text='确定',command=ok)
ok_button.place(x=340,y=20,width=120,height=40)

root_Button_Timer_window.mainloop()