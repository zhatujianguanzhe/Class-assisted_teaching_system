# -*- coding: utf-8 -*-
from tkinter import *
from tkinter.ttk import *
import tkinter.filedialog
from PIL import Image,ImageTk
from datetime import datetime
import tkinter,tkcalendar#tkcalendar是日期控件库,推荐在清华镜像源里下载
import win32api,win32con,pywintypes,win32gui
import ctypes,os,sys,configparser,re,json,time,threading,xlwings,random,subprocess,keyboard,webbrowser

Agreement='''"班级辅助授课系统"软件使用协议:
注意: 在下文中"本软件"及"软件"两词指代"班级辅助授课系统"软件.
注意: 点击"确定"按钮代表您从此往后都同意以下全部协议,如果不同意,请立即点击"取消"按钮关闭本软件.
1.使用本软件所造成的任何形式的任何损失均与本软件开发者无关
2.本软件遵守AGPL-3.0协议,准确内容请查看其官网: www.fsf.org
3.使用了本软件的用户将不能以任何原因起诉本软件开发者.
4.未经授权的用户不得商用此软件(如需商用,请与开发者以E-mail形式联系,获得授权后方可商用).
5.本协议的效力及解释受中华人民共和国法律的管辖.
6.本协议的最终解释权归本软件开发者所有.
7.用户必须同意以上协议才能使用本软件.'''

if getattr(sys, 'frozen', None):
    res_icon_folder = str(sys._MEIPASS)+'\\icons'
else:
    res_icon_folder = str(os.path.dirname(__file__))+'\\icons'

running_threading=True

Permissions='STUDENT'

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


if os.path.exists('DATA/Agreement/Agree')!=True:
    #要求你同意协议
    if win32api.MessageBox(0,Agreement,'协议',win32con.MB_ICONWARNING|win32con.MB_TOPMOST|win32con.MB_DEFBUTTON2|win32con.MB_OKCANCEL)!=1:
        sys.exit()
    else:
        with open('DATA/Agreement/Agree','w',encoding='utf-8') as f:
            f.write('')




def InputBox(title='', text='', parent=None, default='', canspace=False, canempty=False, canspecialtext=False,show=''):
    rt=None
    def close_handler():
        if parent != None:
            parent.attributes('-disabled', 'false')
        Input_Box_window.destroy()
        if parent != None:
            parent.focus_set()

    def close_handler_cancel_(nothing):
        close_handler_cancel()

    def close_handler_cancel():
        nonlocal rt
        rt = None
        close_handler()

    def save():
        nonlocal rt
        rt = str(entry_filename.get())
        if not canspace and ' ' in rt:
            win32api.MessageBeep()
            entry_filename.focus()
            return 
        elif canspecialtext==False and ( '"' in entry_filename.get() or ']' in entry_filename.get() or '[' in entry_filename.get() or '#' in entry_filename.get()):
            win32api.MessageBeep()
            entry_filename.focus()
            return 
        elif canempty==True:
            if canspecialtext==False and ( '"' in entry_filename.get() or ']' in entry_filename.get() or '[' in entry_filename.get() or '#' in entry_filename.get()):
                win32api.MessageBeep()
                entry_filename.focus()
                return 
            else:
                close_handler()
                return
        elif not canempty and rt.replace(' ', '') == '' :
            win32api.MessageBeep()
            entry_filename.focus()
            return 
        else:
            close_handler()
            return

    def save_(nothing):
        nonlocal rt
        try:
            if Input_Box_window.focus_get() == save_btn:
                save()
            elif Input_Box_window.focus_get() == close:
                close_handler_cancel()
            else:
                save()
        except:
            save()

    def focus_see_(nothing):
        if Input_Box_window.focus_get() != None:
            if Input_Box_window.focus_get() == save_btn:
                close['default'] = 'normal'
                save_btn['default'] = 'active'
            elif Input_Box_window.focus_get() == close:
                save_btn['default'] = 'normal'
                close['default'] = 'active'
            else:
                close['default'] = 'normal'
                save_btn['default'] = 'active'

    if parent != None:
        parent.attributes('-disabled', 'true')
        Input_Box_window = Toplevel(parent)
        Input_Box_window.wm_transient(parent)
    else:
        Input_Box_window = Tk()
    
    Input_Box_window.title(str(title))
    width = 420
    height = 120
    screenwidth = Input_Box_window.winfo_screenwidth()
    screenheight = Input_Box_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    Input_Box_window.geometry(geometry)
    Input_Box_window.resizable(width=False, height=False)
    Input_Box_window.bind('<Return>', save_)
    Input_Box_window.focus_set()
    Input_Box_window.protocol("WM_DELETE_WINDOW", close_handler_cancel)
    Input_Box_window.bind('<Escape>', close_handler_cancel_)

    label_filename = Label(Input_Box_window, text=str(text), anchor="w")
    label_filename.place(x=20, y=20, width=80, height=30)

    entry_filename = Entry(Input_Box_window, takefocus=True,show=show)
    entry_filename.place(x=100, y=20, width=300, height=30)
    entry_filename.focus()
    entry_filename.insert(0, str(default))

    save_btn = Button(Input_Box_window, text="确定", takefocus=True, command=save, default='active')
    save_btn.place(x=220, y=70, width=80, height=30)

    close = Button(Input_Box_window, text="取消", takefocus=True, command=close_handler_cancel)
    close.place(x=320, y=70, width=80, height=30)

    save_btn.bind("<FocusIn>", focus_see_)
    save_btn.bind("<FocusOut>", focus_see_)
    close.bind("<FocusIn>", focus_see_)
    close.bind("<FocusOut>", focus_see_)

    Input_Box_window.wm_iconbitmap(str(res_icon_folder) + '/icon.ico')
    Input_Box_window.wait_window(Input_Box_window)

    return rt

def InputComboBox(title='', text='', parent=None,default='',value=(''),state='readonly'):
    rt=''
    def close_handler():
        if parent!=None:
            parent.attributes('-disabled', 'false')
        Input_ComboboxBox_window.destroy()
        if parent!=None:
            parent.focus_set()
    def close_handler_cancel_(nothing):
        close_handler_cancel()
    def close_handler_cancel():
        nonlocal rt
        rt=''
        close_handler()
    def save():
        nonlocal rt
        rt=str(entry_filename.get())
        if rt.replace(' ','')!='':
            close_handler()
            return 0
        else:
            win32api.MessageBeep()
            entry_filename.focus()
            return 0
    def save_(nothing):
        nonlocal rt
        try:
            if Input_ComboboxBox_window.focus_get()==save_btn:
                save()
            elif Input_ComboboxBox_window.focus_get()==close:
                close_handler_cancel()

            else:
                save()
        except:
            save()
    def focus_see_(nothing):
        if Input_ComboboxBox_window.focus_get()!=None:
            if Input_ComboboxBox_window.focus_get()==save_btn:
                close['default']='normal'
                save_btn['default']='active'
            elif Input_ComboboxBox_window.focus_get()==close:
                save_btn['default']='normal'
                close['default']='active'
            else:
                close['default']='normal'
                save_btn['default']='active'


    if parent!=None:
        parent.attributes('-disabled', 'true')
        Input_ComboboxBox_window=Toplevel(parent)
        Input_ComboboxBox_window.wm_transient(parent)
    else:
        Input_ComboboxBox_window=Tk()
    Input_ComboboxBox_window.title(str(title))
    width=420
    height=120

    screenwidth = Input_ComboboxBox_window.winfo_screenwidth()
    screenheight = Input_ComboboxBox_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    Input_ComboboxBox_window.geometry(geometry)
    Input_ComboboxBox_window.resizable(width=False, height=False)
    Input_ComboboxBox_window.bind('<Return>',save_)
    Input_ComboboxBox_window.focus_set()
    Input_ComboboxBox_window.protocol("WM_DELETE_WINDOW", close_handler_cancel)
    Input_ComboboxBox_window.bind('<Escape>',close_handler_cancel_)
    

    label_filename = tkinter.Label(Input_ComboboxBox_window,text=str(text),anchor="w")
    label_filename.place(x=20, y=20, width=80, height=30)

    Var_value=StringVar()
    Var_value.set(str(default))
    entry_filename = Combobox(Input_ComboboxBox_window,values=value,textvariable=Var_value,state=state)
    entry_filename.place(x=100, y=20, width=300, height=30)
    entry_filename.focus()
 

    save_btn= Button(Input_ComboboxBox_window, text="确定",command=save,default='active')
    save_btn.place(x=220, y=70, width=80, height=30)

    close =Button(Input_ComboboxBox_window, text="取消", command=close_handler_cancel)
    close.place(x=320, y=70, width=80, height=30)
    
    save_btn.bind("<FocusIn>", focus_see_)
    save_btn.bind("<FocusOut>", focus_see_)
    close.bind("<FocusIn>", focus_see_)
    close.bind("<FocusOut>", focus_see_)
    
    
    Input_ComboboxBox_window.wm_iconbitmap(str(res_icon_folder)+'/icon.ico')
    Input_ComboboxBox_window.wait_window(Input_ComboboxBox_window)

    return str(rt)

def InputTimeBox(title='', text_begin='',text_end='', parent=None, default=[[0,0],[0,0]]):
    rt=None
    def close_handler():
        if parent!=None:
            parent.attributes('-disabled', 'false')
        InputTimeBox_Window.destroy()
        if parent!=None:
            parent.focus_set()
    def close_handler_cancel_(nothing):
        close_handler_cancel()
    def close_handler_cancel():
        nonlocal rt
        rt=None
        close_handler()
    def save():
        nonlocal rt
        try:
            temp=int(Entry_begin_hour.get())
            if temp<0 or temp>23:
                raise ValueError
        except:
            Entry_begin_hour.focus()
            win32api.MessageBeep()
            return 0
        
        try:
            temp=int(Entry_begin_minute.get())
            if temp<0 or temp>59:
                raise ValueError
        except:
            Entry_begin_minute.focus()
            win32api.MessageBeep()
            return 0
        
        try:
            temp=int(Entry_end_hour.get())
            if temp<0 or temp>23:
                raise ValueError
        except:
            Entry_end_hour.focus()
            win32api.MessageBeep()
            return 0
        
        try:
            temp=int(Entry_end_minute.get())
            if temp<0 or temp>59:
                raise ValueError
        except:
            Entry_end_minute.focus()
            win32api.MessageBeep()
            return 0
        
        rt=[[int(Entry_begin_hour.get()),int(Entry_begin_minute.get())],[int(Entry_end_hour.get()),int(Entry_end_minute.get())]]
        close_handler()

        
    def save_(nothing):
        try:
            if InputTimeBox_Window.focus_get()==save_btn:
                save()
            elif InputTimeBox_Window.focus_get()==close:
                close_handler_cancel()
            else:
                save()
        except:
            save()
    def focus_see_(nothing):
        if InputTimeBox_Window.focus_get()!=None:
            if InputTimeBox_Window.focus_get()==save_btn:
                close['default']='normal'
                save_btn['default']='active'
            elif InputTimeBox_Window.focus_get()==close:
                save_btn['default']='normal'
                close['default']='active'
            else:
                close['default']='normal'
                save_btn['default']='active'


    if parent!=None:
        parent.attributes('-disabled', 'true')
        InputTimeBox_Window=Toplevel(parent)
        InputTimeBox_Window.wm_transient(parent)
    else:
        InputTimeBox_Window=Tk()
    InputTimeBox_Window.title(title)
    # 设置窗口大小、居中
    width = 280
    height = 170
    screenwidth = InputTimeBox_Window.winfo_screenwidth()
    screenheight = InputTimeBox_Window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    InputTimeBox_Window.geometry(geometry)
    InputTimeBox_Window.resizable(width=False, height=False)
    InputTimeBox_Window.bind('<Return>',save_)
    InputTimeBox_Window.protocol("WM_DELETE_WINDOW", close_handler_cancel)
    InputTimeBox_Window.bind('<Escape>',close_handler_cancel_)
    InputTimeBox_Window.focus_set()


    tkinter.Label(InputTimeBox_Window,text=text_begin,anchor="w", ).place(x=20, y=20, width=80, height=30)

    Entry_begin_hour = Spinbox(InputTimeBox_Window, from_=0,to=23,increment=1,wrap=True)
    Entry_begin_hour.place(x=100, y=20, width=70, height=30)
    Entry_begin_hour.insert(0,str(default[0][0]))


    tkinter.Label(InputTimeBox_Window,text=":",anchor="center", ).place(x=170, y=20, width=20, height=30)

    def zfill_Entry_begin_minute():
        if len(Entry_begin_minute.get().replace(' ',''))==1:
            num_=Entry_begin_minute.get().zfill(2)
            Entry_begin_minute.delete(0,'end')
            Entry_begin_minute.insert(0,num_)
    def zfill_Entry_begin_minute_(nothing):
        zfill_Entry_begin_minute()

    Entry_begin_minute= Spinbox(InputTimeBox_Window, from_=0,to=59,increment=1,wrap=True,command=zfill_Entry_begin_minute)
    Entry_begin_minute.place(x=190, y=20, width=70, height=30)
    Entry_begin_minute.insert(0,str(default[0][1]).zfill(2))
    Entry_begin_minute.bind('<FocusOut>',zfill_Entry_begin_minute_)


    tkinter.Label(InputTimeBox_Window,text=text_end,anchor="w", ).place(x=20, y=70, width=80, height=30)

    Entry_end_hour = Spinbox(InputTimeBox_Window, from_=0,to=23,increment=1,wrap=True)
    Entry_end_hour.place(x=100, y=70, width=70, height=30)
    Entry_end_hour.insert(0,str(default[1][0]))


    tkinter.Label(InputTimeBox_Window,text=":",anchor="center", ).place(x=170, y=70, width=20, height=30)

    def zfill_Entry_end_minute():
        if len(Entry_end_minute.get().replace(' ',''))==1:
            num_=Entry_end_minute.get().zfill(2)
            Entry_end_minute.delete(0,'end')
            Entry_end_minute.insert(0,num_)
    def zfill_Entry_end_minute_(nothing):
        zfill_Entry_end_minute()
    Entry_end_minute = Spinbox(InputTimeBox_Window, from_=0,to=59,increment=1,wrap=True,command=zfill_Entry_end_minute)
    Entry_end_minute.place(x=190, y=70, width=70, height=30)
    Entry_end_minute.insert(0,str(default[1][1]).zfill(2))
    Entry_end_minute.bind('<FocusOut>',zfill_Entry_end_minute_)


    save_btn= Button(InputTimeBox_Window, text="确定",command=save,default='active')
    save_btn.place(x=80, y=120, width=80, height=30)

    close =Button(InputTimeBox_Window, text="取消", command=close_handler_cancel)
    close.place(x=180, y=120, width=80, height=30)
    
    save_btn.bind("<FocusIn>", focus_see_)
    save_btn.bind("<FocusOut>", focus_see_)
    close.bind("<FocusIn>", focus_see_)
    close.bind("<FocusOut>", focus_see_)

    InputTimeBox_Window.wm_iconbitmap(str(res_icon_folder)+'/icon.ico')
    InputTimeBox_Window.wait_window(InputTimeBox_Window)

    return rt

def Message_Box(parent,text,title,icon='none',buttonmode=1,defaultfocus=1):
    def set_image(tk_label, img_path, img_size=None):
        img_open = Image.open(img_path)
        if img_size is not None:
            height, width = img_size
            img_h, img_w = img_open.size

            if img_w > width:
                size = (width, int(height * img_w / img_h))
                img_open = img_open.resize(size, Image.LANCZOS)
 
            elif img_h > height:
                size = (int(width * img_h / img_w), height)
                img_open = img_open.resize(size, Image.LANCZOS)
        img = ImageTk.PhotoImage(img_open)
        tk_label.config(image=img)
        tk_label.image = img
    def close_Message_Box_window():
        if parent!=None:
            parent.attributes('-disabled', 'false')
        Message_Box_window.destroy()
        if parent!=None:
            parent.focus_set()
    def ok():
        global rtn
        rtn=True
        close_Message_Box_window()
    def cancel():
        global rtn
        rtn=False
        close_Message_Box_window()
    def ok_(even):
        if Message_Box_window.focus_get()==ok_button:
            ok()
        elif Message_Box_window.focus_get()==cancel_button:
            cancel()
        else:
            ok()
    def cancel_(even):
        cancel()
    def focus_see_(nothing):
        if Message_Box_window.focus_get()!=None:
            if Message_Box_window.focus_get()==ok_button:
                cancel_button['default']='normal'
                ok_button['default']='active'
            elif Message_Box_window.focus_get()==cancel_button:
                ok_button['default']='normal'
                cancel_button['default']='active'
            else:
                cancel_button['default']='normal'
                ok_button['default']='active'
    if parent!=None:
        parent.attributes('-disabled', 'true')
        Message_Box_window=Toplevel(parent)
        Message_Box_window.wm_transient(parent)
    else:
        Message_Box_window=Tk()
    Message_Box_window.title(str(title))
    width = 420
    height = 190
    screenwidth = Message_Box_window.winfo_screenwidth()
    screenheight = Message_Box_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    Message_Box_window.geometry(geometry)
    Message_Box_window.resizable(width=False, height=False)
    Message_Box_window.protocol("WM_DELETE_WINDOW", close_Message_Box_window)
    Message_Box_window.focus_set()
    
    label_icon=tkinter.Label(Message_Box_window,anchor='center')
    label_icon.place(x=20,y=20,width=50,height=50)

    text_text=tkinter.Label(Message_Box_window,wraplength=310,text=str(text),justify='left',anchor='nw')
    text_text.place(x=85,y=20,width=315,height=100)


    if icon=='none':
        text_text.place(x=20,y=20,width=380,height=130)
    else:
        set_image(label_icon,str(res_icon_folder)+'/'+str(icon)+'.ico',img_size=[50,50])

    if buttonmode==2:
        Message_Box_window.bind('<Escape>',cancel_)
        Message_Box_window.bind('<Return>',ok_)

        ok_button=Button(Message_Box_window,text='确定',command=ok,default='active')
        ok_button.place(x=220,y=140,width=80,height=30)

        cancel_button=Button(Message_Box_window,text='取消',command=cancel)
        cancel_button.place(x=320,y=140,width=80,height=30)
        
        ok_button.bind("<FocusIn>", focus_see_)
        ok_button.bind("<FocusOut>", focus_see_)
        cancel_button.bind("<FocusIn>", focus_see_)
        cancel_button.bind("<FocusOut>", focus_see_)
        
        if defaultfocus==1:
            ok_button.focus()
        elif defaultfocus==2:
            cancel_button.focus()
        else:
            raise ValueError('focus只能为1或2')
    elif buttonmode==1:
        Message_Box_window.bind('<Escape>',ok_)
        Message_Box_window.bind('<Return>',ok_)

        ok_button=Button(Message_Box_window,text='确定',command=ok,default='active')
        ok_button.place(x=320,y=140,width=80,height=30)
        ok_button.focus()
    else:
        raise ValueError('buttonmode只能为1或2')
    
    if icon=='question' or icon=='safe_question':
        win32api.MessageBeep(win32con.MB_ICONQUESTION)
    elif icon=='error' or icon=='modern_error' or icon=='safe_error' or icon=='word_error_red' or icon=='word_deny':
        win32api.MessageBeep(win32con.MB_ICONERROR)
    elif icon=='warning' or icon=='safe_warning' or icon=='word_correct_orange' or icon=='modern_warning' or icon=='uac':
        win32api.MessageBeep(win32con.MB_ICONWARNING)
    elif icon=='info' or icon=='word_correct_green' or icon=='safe_correct' or icon=='modern_correct' or icon=='modern_correct_gray':
        win32api.MessageBeep(win32con.MB_ICONINFORMATION)
    elif icon=='none':
        win32api.MessageBeep()
    else:
        raise ValueError('icon只能为question/error/warning/info/none其中之一')

    Message_Box_window.wm_iconbitmap(str(res_icon_folder)+'/icon.ico')
    Message_Box_window.wait_window(Message_Box_window)

    try:
        return rtn
    except:
        if buttonmode==1:
            return True
        else:
            return False

def InputText(parent,title='',state='normal',default=''):
    rt=None
    def close_command_window():
        parent.attributes('-disabled', 'false')
        command_window.destroy()
        parent.focus_set()
    def close_command_window_(nothing):
        close_command_window()
    parent.attributes('-disabled', 'true')
    command_window=Toplevel(parent)
    command_window.wm_transient(parent)
    command_window.title(title)
    width=620      
    height=470     
    screenwidth=command_window.winfo_screenwidth()
    screenheight=command_window.winfo_screenheight()       
    command_window.geometry('%dx%d+%d+%d'%(width,height,(screenwidth - width)/2,(screenheight - height)/2))
    command_window.resizable(False , False) 
    command_window.protocol("WM_DELETE_WINDOW", close_command_window)
    command_window['bd']=0
    command_window['highlightthickness']=0
    command_window.bind('<Escape>',close_command_window_)
    command_window.focus()

    scrollBary=Scrollbar(command_window,orient='vertical')
    scrollBary.place(x=580,y=20,width=20,height=360)

    scrollBarx=Scrollbar(command_window,orient='horizontal')
    scrollBarx.place(x=20,y=380,width=560,height=20)

    command_input=tkinter.Text(command_window,yscrollcommand=scrollBary.set,xscrollcommand=scrollBarx.set ,font='TkDefaultFont' , takefocus='true' , relief='groove' , bd=2,undo='true' , wrap='none')    
    command_input.place(x=20 , y=20 , width=560 , height=360)
    command_input.focus_set()
    command_input.insert(0.0,default)
    command_input['state']=state

    scrollBary.config(command=command_input.yview)
    scrollBarx.config(command=command_input.xview)

    def ok():
        nonlocal rt
        rt=command_input.get(0.0,'end')
        close_command_window()

    ok=Button(command_window , command=ok  , text='确定' ,)      
    ok.place(x=420 , y=420 , width=80 , height=30)

    cancel=Button(command_window , text='取消' ,command=close_command_window)
    cancel.place(x=520 , y=420 , width=80 , height=30)

    command_window.iconbitmap(str(res_icon_folder)+'/icon.ico')
    command_window.wait_window(command_window)
    return rt

def InputDate(parent,title='',):
    rt=None
    def close_ClassTreeWindow():
        parent.attributes('-disabled','false')
        ClassTreeWindow_CompensatoryHolidays_Window.destroy()
        parent.focus_set()
    def close_ClassTreeWindow_(nothing):
        close_ClassTreeWindow()
    def focus_see_(nothing):
        if ClassTreeWindow_CompensatoryHolidays_Window.focus_get()!=None:
            if ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==save_btn:
                close['default']='normal'
                save_btn['default']='active'
            elif ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==close:
                save_btn['default']='normal'
                close['default']='active'
            else:
                close['default']='normal'
                save_btn['default']='active'

    def save():
        nonlocal rt
        temp=calendar.get_date().split('/')
        rt=[int(temp[0]),int(temp[1]),int(temp[2])]
        close_ClassTreeWindow()

    def save_(nothing):
        nonlocal rt
        try:
            if ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==save_btn:
                save()
            elif ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==close:
                rt=None
                close_ClassTreeWindow()
            else:
                save()
        except:
            save()
    parent.attributes('-disabled','true')
    ClassTreeWindow_CompensatoryHolidays_Window=Toplevel(parent)
    ClassTreeWindow_CompensatoryHolidays_Window.wm_transient(parent)
    ClassTreeWindow_CompensatoryHolidays_Window.title(title)
    # 设置窗口大小、居中
    width = 340
    height = 290
    screenwidth = ClassTreeWindow_CompensatoryHolidays_Window.winfo_screenwidth()
    screenheight = ClassTreeWindow_CompensatoryHolidays_Window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    ClassTreeWindow_CompensatoryHolidays_Window.geometry(geometry)
    ClassTreeWindow_CompensatoryHolidays_Window.resizable(width=False, height=False)
    ClassTreeWindow_CompensatoryHolidays_Window.focus()
    ClassTreeWindow_CompensatoryHolidays_Window.bind('<Escape>',close_ClassTreeWindow_)
    ClassTreeWindow_CompensatoryHolidays_Window.bind('<Return>',save_)
    ClassTreeWindow_CompensatoryHolidays_Window.protocol("WM_DELETE_WINDOW", close_ClassTreeWindow)
    ClassTreeWindow_CompensatoryHolidays_Window_handler=int(pywintypes.HANDLE(int(ClassTreeWindow_CompensatoryHolidays_Window.frame(), 16)))

    calendar=tkcalendar.Calendar(ClassTreeWindow_CompensatoryHolidays_Window,firstweekday='monday',weekendbackground='white',weekendforeground='gray30',othermonthweforeground='gray30',othermonthwebackground='gray93')
    calendar.place(x=20,y=20,width=300,height=200)



    save_btn= Button(ClassTreeWindow_CompensatoryHolidays_Window, text="确定",command=save,default='active')
    save_btn.place(x=140, y=240, width=80, height=30)

    close =Button(ClassTreeWindow_CompensatoryHolidays_Window, text="取消", command=close_ClassTreeWindow)
    close.place(x=240, y=240, width=80, height=30)
    
    save_btn.bind("<FocusIn>", focus_see_)
    save_btn.bind("<FocusOut>", focus_see_)
    close.bind("<FocusIn>", focus_see_)
    close.bind("<FocusOut>", focus_see_)

    ClassTreeWindow_CompensatoryHolidays_Window.iconbitmap(str(res_icon_folder)+'/icon.ico')
    ClassTreeWindow_CompensatoryHolidays_Window.wait_window(ClassTreeWindow_CompensatoryHolidays_Window)
    return rt

def Balloon_Box(parent=None,title=None,text=None,staytime=0):
    def close_Balloon_Message_Box_Window():
        def isdestroyed(window):
            try:
                window.winfo_id()
            except:
                return False
            else:
                return True

        screenwidth = Balloon_Message_Box_Window.winfo_screenwidth()
        screenheight = Balloon_Message_Box_Window.winfo_screenheight()
        x=50
        width = 375
        height = 200
        alpha=1
        for i in range (100):
            if isdestroyed(Balloon_Message_Box_Window)==True:
                alpha=alpha-0.01
                Balloon_Message_Box_Window.attributes('-alpha', alpha)
                geometry = '%dx%d+%d+%d' % (width, height, screenwidth - width+x, screenheight - height-100)
                Balloon_Message_Box_Window.geometry(geometry)
                Balloon_Message_Box_Window.update()
                x=x+5
        if isdestroyed(Balloon_Message_Box_Window)==True:
            Balloon_Message_Box_Window.destroy()

    if parent!=None:
        Balloon_Message_Box_Window=Toplevel(parent)
    else:
        Balloon_Message_Box_Window=Tk()
    Balloon_Message_Box_Window.overrideredirect(True)
    Balloon_Message_Box_Window.title("")
    width = 370
    height = 200

    Balloon_Message_Box_Window.resizable(width=False, height=False)
    Balloon_Message_Box_Window.attributes('-topmost', 'true')
    Balloon_Message_Box_Window.config(cursor='arrow')
    Balloon_Message_Box_Window.bind('<ButtonRelease-1>', lambda event: close_Balloon_Message_Box_Window())

    rtitle = tkinter.Label(Balloon_Message_Box_Window,text=title,anchor="w", font='微软雅黑 12',fg='#0078e4')
    rtitle.place(x=10, y=10, width=350, height=30)
    rtitle.bind('<ButtonRelease-1>', lambda event: close_Balloon_Message_Box_Window())

    rtext= tkinter.Label(Balloon_Message_Box_Window,text=text,anchor="nw")
    rtext.place(x=10, y=40, width=350, height=150)
    rtext.bind('<ButtonRelease-1>', lambda event: close_Balloon_Message_Box_Window())

    
    win32api.MessageBeep(win32con.MB_ICONINFORMATION)
    x=-450
    alpha=0
    for i in range (99):
        alpha=alpha+0.01 
        Balloon_Message_Box_Window.attributes('-alpha', alpha)
        geometry = '%dx%d+%d+%d' % (width, height, Balloon_Message_Box_Window.winfo_screenwidth() - width-x, Balloon_Message_Box_Window.winfo_screenheight() - height-100)
        Balloon_Message_Box_Window.geometry(geometry)
        Balloon_Message_Box_Window.update()
        x=x+5

    if staytime<=0:
        pass
    else:
        Balloon_Message_Box_Window.after(staytime*1000,close_Balloon_Message_Box_Window)
    
    Balloon_Message_Box_Window.wait_window(Balloon_Message_Box_Window)

def Password_Box(title='',text='',parent=None,defaultuser='',defaultpassword='',defaultusertuple=tuple(),savepassword_state=False,usernamestate='normal'):
    rt=(None,None,'CANCEL')
    def close_password_box_window():
        if parent!=None:
            parent.attributes('-disabled','false')
        password_box_window.destroy()
        if parent!=None:
            parent.focus_set()


    def close_password_box_window_(nothing):
        nonlocal rt
        rt=(None,None,'CANCEL')
        close_password_box_window()

    def close_password_box_window_none():
        close_password_box_window_(0)

    def ok():
        nonlocal rt
        if password_box_window.focus_get() is not None:
            if password_box_window.focus_get() == button_ok:
                entry_username__=entry_username.get()
                entry_password__=entry_password.get()
                rt=(entry_username__,entry_password__,save_cb.get())
                close_password_box_window()
            elif password_box_window.focus_get() == button_cancel:
                rt=(None,None,None)
                close_password_box_window()
            else:
                entry_username__=entry_username.get()
                entry_password__=entry_password.get()
                rt=(entry_username__,entry_password__,save_cb.get())
                close_password_box_window()
    def ok_(nothing):
        ok()
    if parent!=None:
        parent.attributes('-disabled','true')
        password_box_window=Toplevel(parent)
        password_box_window.transient(parent)
    else:
        password_box_window=Tk()
    password_box_window.title(title)
    # 设置窗口大小、居中
    width = 480
    height = 395
    screenwidth = password_box_window.winfo_screenwidth()
    screenheight = password_box_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height-100) / 2)
    password_box_window.geometry(geometry)
    password_box_window.config(highlightthickness=0,bd=0)
    password_box_window.protocol("WM_DELETE_WINDOW", close_password_box_window_none)
    password_box_window.resizable(False,False)
    password_box_window.bind('<Escape>',close_password_box_window_)
    password_box_window.bind('<Return>',ok_)

    img=PhotoImage(width=480, height=93,file=res_icon_folder+'\\passwordinput_headpicture.gif')
    tkinter.Label(password_box_window,image=img,borderwidth=0,).place(x=0,y=0,width=480,height=93)

    tkinter.Label(password_box_window,text=text,anchor='w',).place(x=15,y=100,width=455,height=50)

    tkinter.Label(password_box_window,text='用户名(U):',anchor='w',underline=4).place(x=15,y=160,width=155,height=30)
    entry_username=Combobox(password_box_window,values=defaultusertuple)
    entry_username.place(x=170,y=160,width=250,height=30)
    entry_username.insert(0,defaultuser)
    entry_username['state']=usernamestate
    def entry_username_focus_(nothing):
        entry_username.focus()
    password_box_window.bind('<Alt-u>',entry_username_focus_)


    Button(password_box_window,text='...',state='disabled',underline=0).place(x=430,y=160,width=35,height=30)

    tkinter.Label(password_box_window,text='密码(P):',anchor='w',underline=3).place(x=15,y=200,width=155,height=30)
    entry_password=Entry(password_box_window,show='●')
    entry_password.place(x=170,y=200,width=250,height=30)
    entry_password.insert(0,defaultpassword)
    entry_password.focus()
    def entry_password_focus_(nothing):
        entry_password.focus()
    password_box_window.bind('<Alt-p>',entry_password_focus_)

    s=Style()
    s.configure('c.TCheckbutton',anchor='w')
    save_cb=BooleanVar()
    save_cb.set(False)
    savepas=Checkbutton(password_box_window,text='记住我的密码(R)',underline=7,style='c.TCheckbutton',variable=save_cb,onvalue=True,offvalue=False)
    savepas.place(x=170,y=240,width=250,height=30)
    def savepas_focus_(nothing):
        savepas.focus()
    if savepassword_state==True:
        password_box_window.bind('<Alt-r>',savepas_focus_)
    else:
        savepas['state']='disabled'

    button_ok=Button(password_box_window,text='确定',default='active',command=ok)
    button_ok.place(x=225,y=345,width=115,height=35)

    button_cancel=Button(password_box_window,text='取消',command=close_password_box_window)
    button_cancel.place(x=350,y=345,width=115,height=35)

    def focus_see_(nothing):
        if password_box_window.focus_get() != None:
            if password_box_window.focus_get() == button_ok:
                button_cancel['default'] = 'normal'
                button_ok['default'] = 'active'
            elif password_box_window.focus_get() == button_cancel:
                button_ok['default'] = 'normal'
                button_cancel['default'] = 'active'
            else:
                button_cancel['default'] = 'normal'
                button_ok['default'] = 'active'
    
    savepas.bind("<FocusIn>", focus_see_)
    savepas.bind("<FocusOut>", focus_see_)
    button_ok.bind("<FocusIn>", focus_see_)
    button_ok.bind("<FocusOut>", focus_see_)
    button_cancel.bind("<FocusIn>", focus_see_)
    button_cancel.bind("<FocusOut>", focus_see_)

    password_box_window.wm_iconbitmap(str(res_icon_folder)+'\\icon.ico')
    password_box_window.wait_window(password_box_window)
    try:
        return rt
    except:
        return (None,None,'CANCEL')

def Link_Message_Box(title='',bigtext='',text='',text1='',text2='',parent=None,defaultfocus=1,defaultreturn=None):
    def close_Message_Box_window():
        if parent!=None:
            parent.attributes('-disabled', 'false')
        Link_Message_Box_window.destroy()
        if parent!=None:
            parent.focus_set()
    def ok():
        global rtn
        rtn=1
        close_Message_Box_window()
    def cancel():
        global rtn
        rtn=2
        close_Message_Box_window()
    def ok_(even):
        if Link_Message_Box_window==ok_button:
            ok()
        elif Link_Message_Box_window==cancel_button:
            cancel()
        else:
            ok()
    def cancel_(even):
        global rtn
        rtn=defaultreturn
        close_Message_Box_window()
    if parent!=None:
        parent.attributes('-disabled', 'true')
        Link_Message_Box_window=Toplevel(parent)
        Link_Message_Box_window.wm_transient(parent)
    else:
        Link_Message_Box_window=Tk()
    Link_Message_Box_window.title(str(title))
    Link_Message_Box_window.config(background='white')
    width = 400
    height = 205
    screenwidth = Link_Message_Box_window.winfo_screenwidth()
    screenheight = Link_Message_Box_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    Link_Message_Box_window.geometry(geometry)
    Link_Message_Box_window.resizable(width=False, height=False)
    Link_Message_Box_window.protocol("WM_DELETE_WINDOW", close_Message_Box_window)
    Link_Message_Box_window.bind('<Escape>',cancel_)
    Link_Message_Box_window.bind('<Return>',ok_)
    Link_Message_Box_window.focus_set()

    label_bigtext=tkinter.Label(Link_Message_Box_window,bg='white',font='微软雅黑 12',fg='#003399',text=str(bigtext),anchor='w')
    label_bigtext.place(x=10,y=10,width=400,height=30)
    
    label_text=tkinter.Label(Link_Message_Box_window,bg='white',justify='left',wraplength = 320,anchor='w',text=str(text))
    label_text.place(x=10,y=50,width=400,height=30)
    
    ok_button=tkinter.Button(Link_Message_Box_window,text=' → '+str(text1),command=ok,takefocus=True,anchor='w',bg='white',bd=0,activebackground='#bcdcf4',activeforeground='#0078e4',foreground='#0078e4',font='微软雅黑 12')
    ok_button.place(x=10,y=95,width=380,height=50)

    def Enter_ok_button(n):
        ok_button.focus()
        ok_button['bg']='#d9ebf9'
    ok_button.bind('<Enter>',Enter_ok_button)
    def Leave_ok_button(n):
        ok_button['bg']='white'
    ok_button.bind('<Leave>',Leave_ok_button)

    cancel_button=tkinter.Button(Link_Message_Box_window,text=' → '+str(text2),command=cancel,takefocus=True,anchor='w',bg='white',bd=0,activebackground='#bcdcf4',activeforeground='#0078e4',foreground='#0078e4',font='微软雅黑 12')
    cancel_button.place(x=10,y=145,width=380,height=50)
    def Enter_cancel_button(n):
        cancel_button.focus()
        cancel_button['bg']='#d9ebf9'
    cancel_button.bind('<Enter>',Enter_cancel_button)
    def Leave_cancel_button(n):
        cancel_button['bg']='white'
    cancel_button.bind('<Leave>',Leave_cancel_button)

    if defaultfocus==1:
        ok_button.focus()
    elif defaultfocus==2:
        cancel_button.focus()
    else:
        raise ValueError('focus只能为1或2')

    Link_Message_Box_window.wm_iconbitmap(str(res_icon_folder)+'\\icon.ico')
    Link_Message_Box_window.wait_window(Link_Message_Box_window)

    try:
        return rtn
    except:
        return defaultreturn




try:#正常课程表
    with open('DATA/ClassTree/ClassTree.json','r',encoding='utf-8') as f:
        json.loads(f.read())
except Exception as error:
    Message_Box(None, '读取课程表文件失败.\n此错误不影响程序正常运行,但是会导致课程表及其附属功能无法使用.\n错误代码: ' + str(error), '错误',icon='error')


try:#读调休文件
    remove_BOM('DATA/CompensatoryHolidays/CompensatoryHolidays.ini')
    configparser.ConfigParser().read('DATA/CompensatoryHolidays/CompensatoryHolidays.ini',encoding='utf-8')
except Exception as error:
    Message_Box(None, '读取调休表文件失败.\n此错误不影响程序正常运行,但是会导致调休日及其附属功能无法使用.\n错误代码: ' + str(error), '错误', icon='error')











def mousedown(event):
    widget=event.widget
    widget.startx=event.x # 开始拖动时, 记录控件位置
    widget.starty=event.y
def drag(event):
    widget=event.widget
    dx=event.x-widget.startx
    dy=event.y-widget.starty
    # winfo_x(),winfo_y() 方法获取控件的坐标
    if isinstance(widget,Wm):
        widget.geometry("+%d+%d"%(widget.winfo_x()+dx,widget.winfo_y()+dy))#窗口
def draggable(tkwidget):
    # tkwidget为一个控件(Widget)或一个窗口(Wm)
    tkwidget.bind("<Button-1>",mousedown,add='+')
    tkwidget.bind("<B1-Motion>",drag,add='+')

def root_hide():
    root.attributes('-alpha',0.3)
    root.geometry('+%d+%d'%(root.winfo_x(),-root.winfo_height()+30))
    Button_root_show.place(x=0,y=root.winfo_height()-30,width=300,height=30)
def root_hide_(o):
    root_hide()
def root_show():
    root.attributes('-alpha',1)
    root.geometry('+%d+%d' % (root.winfo_x(), 100 ))
    Button_root_show.place(x=0,y=root.winfo_height()-20,width=300,height=30)


root=Tk()
draggable(root)#允许root被拖动
root.title("班级辅助授课系统")
# 设置窗口大小、居中
width = 300
height = root.winfo_screenheight()
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
geometry = '%dx%d+%d+%d' % (width, height, root.winfo_screenwidth()-300 , 100 )
root.geometry(geometry)
root.resizable(width=False, height=False)
root.bind('<Escape>',root_hide_)
root.protocol("WM_DELETE_WINDOW", root_hide)
root.attributes('-topmost','true')
root.update()
root_handler=int(pywintypes.HANDLE(int(root.frame(), 16)))



def DO_SHOW_Var_class_on_notice():
    global class_on_notice
    class_on_notice=Var_class_on_notice.get()
Var_class_on_notice=BooleanVar()
Var_class_on_notice.set(True)
Checkbutton_root_class_on_notice=Checkbutton(root,text='上下课提醒',variable=Var_class_on_notice,onvalue=True,offvalue=False,command=DO_SHOW_Var_class_on_notice)
Checkbutton_root_class_on_notice.place(x=0,y=0,width=150,height=40)
class_on_notice=Var_class_on_notice.get()



def sort_dict_by_time(input_dict):
    # 提取字典中的键和对应的时间列表
    time_list = [(key, value['Time'][0]) for key, value in input_dict.items()]
    
    # 按照时间排序，先按小时排序，如果小时相同则按分钟排序
    sorted_time_list = sorted(time_list, key=lambda x: (x[1][0], x[1][1]))
    
    # 提取排序后的键并返回
    sorted_keys = [item[0] for item in sorted_time_list]
    return sorted_keys
def Button_root_classtree_open_ClassTreeWindow():
    global Permissions
    def close_ClassTreeWindow():
        ClassTreeWindow.destroy()
        root.focus_set()
    def close_ClassTreeWindow_(nothing):
        close_ClassTreeWindow()

    ClassTreeWindow=Toplevel(root)
    #ClassTreeWindow.wm_transient(root)
    ClassTreeWindow.title("课程表")
    # 设置窗口大小、居中
    width = 920
    height = 560
    screenwidth = ClassTreeWindow.winfo_screenwidth()
    screenheight = ClassTreeWindow.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    ClassTreeWindow.geometry(geometry)
    ClassTreeWindow.resizable(width=False, height=False)
    ClassTreeWindow.focus()
    ClassTreeWindow.bind('<Escape>',close_ClassTreeWindow_)
    ClassTreeWindow.protocol("WM_DELETE_WINDOW", close_ClassTreeWindow)

    yscrollbar=Scrollbar(ClassTreeWindow, orient='vertical', )
    yscrollbar.place(x=760,y=20,width=20,height=520)

    Tree_ClassTreeWindow_tree = Treeview(ClassTreeWindow, show="headings",columns=('Class','Time','Monday','Tuesday','Wednesday','Thursday','Friday'),yscrollcommand=yscrollbar.set)
    Tree_ClassTreeWindow_tree.place(x=20, y=20, width=740, height=520)

    yscrollbar.config(command=Tree_ClassTreeWindow_tree.yview)

    Tree_ClassTreeWindow_tree.heading('Class', text="标题", anchor='center')
    Tree_ClassTreeWindow_tree.column('Class', width=100, stretch=False,anchor='w')

    Tree_ClassTreeWindow_tree.heading('Time', text="时间", anchor='center')
    Tree_ClassTreeWindow_tree.column('Time', width=120, stretch=False,anchor='w')

    Tree_ClassTreeWindow_tree.heading('Monday', text="周一", anchor='center')
    Tree_ClassTreeWindow_tree.column('Monday', width=100, stretch=False,anchor='center')

    Tree_ClassTreeWindow_tree.heading('Tuesday', text="周二", anchor='center')
    Tree_ClassTreeWindow_tree.column('Tuesday', width=100, stretch=False,anchor='center')

    Tree_ClassTreeWindow_tree.heading('Wednesday', text="周三", anchor='center')
    Tree_ClassTreeWindow_tree.column('Wednesday', width=100, stretch=False,anchor='center')

    Tree_ClassTreeWindow_tree.heading('Thursday', text="周四", anchor='center')
    Tree_ClassTreeWindow_tree.column('Thursday',  width=100, stretch=False,anchor='center')

    Tree_ClassTreeWindow_tree.heading('Friday', text="周五", anchor='center')
    Tree_ClassTreeWindow_tree.column('Friday',  width=100, stretch=False,anchor='center')

    try:
        with open('DATA/ClassTree/ClassTree.json','r',encoding='utf-8') as f:
            Class_Tree = json.loads(f.read())#Class_Tree={'第一节': {'Time':[[7, 45], [8, 40]],'Monday': '语文','Tuesday': '数学','Wednesday': '信息','Thursday':'英语','Friday': '体育',},'第二节': {'Time':[[7, 45], [8, 40]],'Monday': '美术','Tuesday': '班会','Wednesday': '阅读','Thursday': '交流','Friday': '英语',},}
    except Exception as error:
        Message_Box(ClassTreeWindow, '读取课程表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
        return 0

    for child in Tree_ClassTreeWindow_tree.get_children():
        Tree_ClassTreeWindow_tree.delete(child)
    for item in sort_dict_by_time(Class_Tree):
        Tree_ClassTreeWindow_tree.insert('', values=(item,'%s:%s~%s:%s'%(str(Class_Tree[item]['Time'][0][0]),str(Class_Tree[item]['Time'][0][1]).zfill(2),str(Class_Tree[item]['Time'][1][0]),str(Class_Tree[item]['Time'][1][1]).zfill(2),),Class_Tree[item]['Monday'],Class_Tree[item]['Tuesday'],Class_Tree[item]['Wednesday'],Class_Tree[item]['Thursday'],Class_Tree[item]['Friday']), index='end')


    def new_Temporary_Class_Tree():
        title=InputBox(parent=ClassTreeWindow,title='新建课程表行的标题',text='标题:',canspecialtext=False,canspace=True)
        if title==None: return 0
        try: int(title.replace(' ','')) 
        except: pass
        else:
            Message_Box(ClassTreeWindow,'新建课程表行失败.\n错误代码: 该标题不符合命名规范,禁止纯数字的标题.','错误',icon='error')
            return 0
        if title in Class_Tree:
            Message_Box(ClassTreeWindow,'新建课程表行失败.\n错误代码: 该标题已存在,禁止重复使用同一标题.','错误',icon='error')
            return 0

        Class_Tree[title]={'Time': [[0, 0], [0, 0]], 'Monday': '--', 'Tuesday': '--', 'Wednesday': '--', 'Thursday': '--', 'Friday': '--'}
        #每次课表操作都进行保存
        with open('DATA/ClassTree/ClassTree.json','w',encoding='utf-8') as f:
            f.write(json.dumps(Class_Tree,ensure_ascii=False))  #实用utf格式编码,不使用文本,保证安全,不受编辑(懒得转文本)
        
        #刷新表格
        for child in Tree_ClassTreeWindow_tree.get_children():
            Tree_ClassTreeWindow_tree.delete(child)
        for item in sort_dict_by_time(Class_Tree):
            Tree_ClassTreeWindow_tree.insert('', values=(item,'%s:%s~%s:%s'%(str(Class_Tree[item]['Time'][0][0]),str(Class_Tree[item]['Time'][0][1]).zfill(2),str(Class_Tree[item]['Time'][1][0]),str(Class_Tree[item]['Time'][1][1]).zfill(2),),Class_Tree[item]['Monday'],Class_Tree[item]['Tuesday'],Class_Tree[item]['Wednesday'],Class_Tree[item]['Thursday'],Class_Tree[item]['Friday']), index='end')
    Button_ClassTreeWindow_newclass = Button(ClassTreeWindow, text="新建", command=new_Temporary_Class_Tree)
    Button_ClassTreeWindow_newclass.place(x=800, y=20, width=100, height=40)


    def change_Temporary_Class_Tree(mode):
        try:
            selected_item = Tree_ClassTreeWindow_tree.selection()[0]
            values = Tree_ClassTreeWindow_tree.item(selected_item)
        except:
            if mode=='Hand': win32api.MessageBeep()
            Tree_ClassTreeWindow_tree.focus_set()
            return
        weekday_change=InputComboBox(title='编辑课程表行的项',text='项:',parent=ClassTreeWindow,state='readonly',value=('时间','周一','周二','周三','周四','周五'))
        if weekday_change=='时间': weekday_change='Time';weekday_change_number=1
        elif weekday_change=='周一': weekday_change='Monday';weekday_change_number=2#因为有日期,所以数字是原来的+1
        elif weekday_change=='周二': weekday_change='Tuesday';weekday_change_number=3
        elif weekday_change=='周三': weekday_change='Wednesday';weekday_change_number=4
        elif weekday_change=='周四': weekday_change='Thursday';weekday_change_number=5
        elif weekday_change=='周五': weekday_change='Friday';weekday_change_number=6
        elif weekday_change=='': return 0

        if weekday_change=='Time':
            Class_Time=InputTimeBox(title='编辑项的课程时间(24小时制)',text_begin='上课时间:',text_end='下课时间:',parent=ClassTreeWindow,default=Class_Tree[values['values'][0]][weekday_change])
            if Class_Time==None: return 0
            Class_Tree[values['values'][0]][weekday_change]=Class_Time
        else:
            Class_name=InputBox(title='编辑项的课程',text='课程:',canspecialtext=False,parent=ClassTreeWindow,default=values['values'][weekday_change_number])
            if Class_name==None: return 0
            Class_Tree[values['values'][0]][weekday_change]=Class_name

        #每次课表操作都进行保存
        with open('DATA/ClassTree/ClassTree.json','w',encoding='utf-8') as f:
            f.write(json.dumps(Class_Tree,ensure_ascii=False))  #实用utf格式编码,不使用文本,保证安全,不受编辑(懒得转文本)
        
        #刷新表格
        for child in Tree_ClassTreeWindow_tree.get_children():
            Tree_ClassTreeWindow_tree.delete(child)
        for item in sort_dict_by_time(Class_Tree):
            Tree_ClassTreeWindow_tree.insert('', values=(item,'%s:%s~%s:%s'%(str(Class_Tree[item]['Time'][0][0]),str(Class_Tree[item]['Time'][0][1]).zfill(2),str(Class_Tree[item]['Time'][1][0]),str(Class_Tree[item]['Time'][1][1]).zfill(2),),Class_Tree[item]['Monday'],Class_Tree[item]['Tuesday'],Class_Tree[item]['Wednesday'],Class_Tree[item]['Thursday'],Class_Tree[item]['Friday']), index='end')
    def change_Temporary_Class_Tree_hand():
        change_Temporary_Class_Tree('Hand')
    def change_Temporary_Class_Tree_(nothing):
        change_Temporary_Class_Tree('Key')
    Button_ClassTreeWindow_temporaryclass = Button(ClassTreeWindow, text="编辑", command=change_Temporary_Class_Tree_hand)
    Button_ClassTreeWindow_temporaryclass.place(x=800, y=80, width=100, height=40)

    Tree_ClassTreeWindow_tree.bind('<Double-Button-1>',change_Temporary_Class_Tree_)

    def pop_Temporary_Class_Tree():
        try:
            selected_item = Tree_ClassTreeWindow_tree.selection()[0]
            values = Tree_ClassTreeWindow_tree.item(selected_item)
        except:
            win32api.MessageBeep()
            Tree_ClassTreeWindow_tree.focus_set()
            return 0
        if Message_Box(ClassTreeWindow,'删除课程表行无法撤回,是否继续?','疑问',icon='question',buttonmode=2)!=True: return 0
        try:
            Class_Tree.pop(values['values'][0])
        except Exception as error:
            Message_Box(ClassTreeWindow, '删除课程表行失败.\n错误代码: ' + str(error), '错误', icon='error')
            return 0
    
        #每次课表操作都进行保存
        with open('DATA/ClassTree/ClassTree.json','w',encoding='utf-8') as f:
            f.write(json.dumps(Class_Tree,ensure_ascii=False))  
        
        #刷新表格
        for child in Tree_ClassTreeWindow_tree.get_children():
            Tree_ClassTreeWindow_tree.delete(child)
        for item in sort_dict_by_time(Class_Tree):
            Tree_ClassTreeWindow_tree.insert('', values=(item,'%s:%s~%s:%s'%(str(Class_Tree[item]['Time'][0][0]),str(Class_Tree[item]['Time'][0][1]).zfill(2),str(Class_Tree[item]['Time'][1][0]),str(Class_Tree[item]['Time'][1][1]).zfill(2),),Class_Tree[item]['Monday'],Class_Tree[item]['Tuesday'],Class_Tree[item]['Wednesday'],Class_Tree[item]['Thursday'],Class_Tree[item]['Friday']), index='end')
    Button_ClassTreeWindow_popclass = Button(ClassTreeWindow, text="删除", command=pop_Temporary_Class_Tree)
    Button_ClassTreeWindow_popclass.place(x=800, y=140, width=100, height=40)


    def import_Temporary_Class_Tree():
        file=tkinter.filedialog.askopenfilename(parent=ClassTreeWindow,title='导入课程表文件',filetypes=[('课程表文件','.json')])
        if file=='' or file==None: return 0
        if Message_Box(ClassTreeWindow,'导入课程表文件将会覆盖原课程表文件,且无法撤回,是否继续?','警告',icon='warning',buttonmode=2)!=True: return 0
        try:
            with open(file,'r+',encoding='utf-8') as f:
                Class_Tree_import = json.loads(f.read())
        except Exception as error:
            Message_Box(ClassTreeWindow, '导入课程表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return 0
        else:
            try:
                with open('DATA/ClassTree/ClassTree.json','w',encoding='utf-8') as f:#课表进行覆写
                    f.write(json.dumps(Class_Tree_import,ensure_ascii=False))  
            except Exception as error:
                Message_Box(ClassTreeWindow, '导入课程表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
                return 0
            else:
                for child in Tree_ClassTreeWindow_tree.get_children():
                    Tree_ClassTreeWindow_tree.delete(child)
                for item in sort_dict_by_time(Class_Tree):
                    Tree_ClassTreeWindow_tree.insert('', values=(item,'%s:%s~%s:%s'%(str(Class_Tree_import[item]['Time'][0][0]),str(Class_Tree_import[item]['Time'][0][1]).zfill(2),str(Class_Tree_import[item]['Time'][1][0]),str(Class_Tree_import[item]['Time'][1][1]).zfill(2),),Class_Tree_import[item]['Monday'],Class_Tree_import[item]['Tuesday'],Class_Tree_import[item]['Wednesday'],Class_Tree_import[item]['Thursday'],Class_Tree_import[item]['Friday']), index='end')
                Message_Box(ClassTreeWindow, '导入课程表文件成功.\n文件路径: \n'+file, '信息', icon='info')
    Button_ClassTreeWindow_importclass = Button(ClassTreeWindow, text="导入", command=import_Temporary_Class_Tree)
    Button_ClassTreeWindow_importclass.place(x=800, y=200, width=100, height=40)


    def output_Temporary_Class_Tree():
        file=tkinter.filedialog.asksaveasfilename(parent=ClassTreeWindow,title='导出课程表文件',filetypes=[('课程表文件','.json')],defaultextension='*.json')
        if file=='' or file==None: return 0
        try:
            with open('DATA/ClassTree/ClassTree.json','r',encoding='utf-8') as f:
                Class_Tree = json.loads(f.read())#Class_Tree={'第一节': {'Time':[[7, 45], [8, 40]],'Monday': '语文','Tuesday': '数学','Wednesday': '信息','Thursday':'英语','Friday': '体育',},'第二节': {'Time':[[7, 45], [8, 40]],'Monday': '美术','Tuesday': '班会','Wednesday': '阅读','Thursday': '交流','Friday': '英语',},}
        except Exception as error:
            Message_Box(ClassTreeWindow, '导出课程表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return 0
        else:
            try:
                with open(file,'w',encoding='utf-8') as f:#课表进行覆写
                    f.write(json.dumps(Class_Tree,ensure_ascii=False))  
            except Exception as error:
                Message_Box(ClassTreeWindow, '导出课程表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
                return 0
            else:
                Message_Box(ClassTreeWindow, '导出课程表文件成功.\n文件路径: \n'+file, '信息', icon='info')
    Button_ClassTreeWindow_outputclass = Button(ClassTreeWindow, text="导出", command=output_Temporary_Class_Tree)
    Button_ClassTreeWindow_outputclass.place(x=800, y=260, width=100, height=40)


    def output_Class_Tree_excel():
        file=tkinter.filedialog.asksaveasfilename(parent=ClassTreeWindow,title='导出课程表文件为Excel',filetypes=[('Excel表格(旧版)','.xls'),('Excel表格(新版)','.xlsx')],defaultextension='*.xls')
        Class_Excel=xlwings.App(visible=True, add_book=False)
        Class_Excel.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        Class_Excel.screen_updating = False    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        Class_Excel_book = Class_Excel.books.open(file)
        Class_Excel_book_sheet = Class_Excel_book.sheets.add()


        Class_Excel_book_sheet.range(1,1).value = '标题'
        Class_Excel_book_sheet.range(1,2).value = '时间'
        Class_Excel_book_sheet.range(1,3).value = '周一'
        Class_Excel_book_sheet.range(1,4).value = '周二'
        Class_Excel_book_sheet.range(1,5).value = '周三'
        Class_Excel_book_sheet.range(1,6).value = '周四'
        Class_Excel_book_sheet.range(1,7).value = '周五'
        
        n=2
        for item in sort_dict_by_time(Class_Tree):
            value_excel=(item,'%s:%s~%s:%s'%(str(Class_Tree[item]['Time'][0][0]),str(Class_Tree[item]['Time'][0][1]).zfill(2),str(Class_Tree[item]['Time'][1][0]),str(Class_Tree[item]['Time'][1][1]).zfill(2),),Class_Tree[item]['Monday'],Class_Tree[item]['Tuesday'],Class_Tree[item]['Wednesday'],Class_Tree[item]['Thursday'],Class_Tree[item]['Friday'])
            Class_Excel_book_sheet.range(n,1).value = value_excel[0]
            Class_Excel_book_sheet.range(n,2).value = value_excel[1]
            Class_Excel_book_sheet.range(n,3).value = value_excel[2]
            Class_Excel_book_sheet.range(n,4).value = value_excel[3]
            Class_Excel_book_sheet.range(n,5).value = value_excel[4]
            Class_Excel_book_sheet.range(n,6).value = value_excel[5]
            Class_Excel_book_sheet.range(n,7).value = value_excel[6]
            n=n+1
    Button_ClassTreeWindow_makexlsx = Button(ClassTreeWindow, text="导出为Excel",command=output_Class_Tree_excel)
    Button_ClassTreeWindow_makexlsx.place(x=800, y=320, width=100, height=40)


    def set_CompensatoryHolidays():
        def edit_CompensatoryHolidays(parent,title='新建调休日的属性',normalday=''):
            rt=None
            def close_ClassTreeWindow():
                parent.attributes('-disabled','false')
                ClassTreeWindow_CompensatoryHolidays_Window.destroy()
                parent.focus_set()
            def close_ClassTreeWindow_(nothing):
                close_ClassTreeWindow()
            def focus_see_(nothing):
                if ClassTreeWindow_CompensatoryHolidays_Window.focus_get()!=None:
                    if ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==save_btn:
                        close['default']='normal'
                        save_btn['default']='active'
                    elif ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==close:
                        save_btn['default']='normal'
                        close['default']='active'
                    else:
                        close['default']='normal'
                        save_btn['default']='active'

            def save():
                nonlocal rt
                if to_week.get()=='':
                    to_week.focus()
                    win32api.MessageBeep()
                    return 0
                temp=calendar.get_date().split('/')
                rt=[[int(temp[0]),int(temp[1]),int(temp[2])],to_week.get().replace('周一','Monday').replace('周二','Tuesday').replace('周三','Wednesday').replace('周四','Thursday').replace('周五','Friday')]
                close_ClassTreeWindow()



            def save_(nothing):
                nonlocal rt
                try:
                    if ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==save_btn:
                        save()
                    elif ClassTreeWindow_CompensatoryHolidays_Window.focus_get()==close:
                        rt=None
                        close_ClassTreeWindow()
                    else:
                        save()
                except:
                    save()
            parent.attributes('-disabled','true')
            ClassTreeWindow_CompensatoryHolidays_Window=Toplevel(parent)
            ClassTreeWindow_CompensatoryHolidays_Window.wm_transient(parent)
            ClassTreeWindow_CompensatoryHolidays_Window.title(title)
            # 设置窗口大小、居中
            width = 340
            height = 440
            screenwidth = ClassTreeWindow_CompensatoryHolidays_Window.winfo_screenwidth()
            screenheight = ClassTreeWindow_CompensatoryHolidays_Window.winfo_screenheight()
            geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
            ClassTreeWindow_CompensatoryHolidays_Window.geometry(geometry)
            ClassTreeWindow_CompensatoryHolidays_Window.resizable(width=False, height=False)
            ClassTreeWindow_CompensatoryHolidays_Window.focus()
            ClassTreeWindow_CompensatoryHolidays_Window.bind('<Escape>',close_ClassTreeWindow_)
            ClassTreeWindow_CompensatoryHolidays_Window.bind('<Return>',save_)
            ClassTreeWindow_CompensatoryHolidays_Window.protocol("WM_DELETE_WINDOW", close_ClassTreeWindow)

            tkinter.Label(ClassTreeWindow_CompensatoryHolidays_Window,text='调休日:',anchor='w').place(x=20,y=20,width=120,height=30)
            calendar=tkcalendar.Calendar(ClassTreeWindow_CompensatoryHolidays_Window,firstweekday='monday',weekendbackground='white',weekendforeground='gray30',othermonthweforeground='gray30',othermonthwebackground='gray93')
            calendar.place(x=20,y=70,width=300,height=200)

            tkinter.Label(ClassTreeWindow_CompensatoryHolidays_Window,text='对应工作日:',anchor='w').place(x=20,y=290,width=120,height=30)
            to_week=Combobox(ClassTreeWindow_CompensatoryHolidays_Window,values=('周一','周二','周三','周四','周五'),cursor='arrow')
            to_week.place(x=20,y=340,width=300,height=30)
            to_week.insert(0,normalday)
            to_week['state']='readonly'

            save_btn= Button(ClassTreeWindow_CompensatoryHolidays_Window, text="确定",command=save,default='active')
            save_btn.place(x=140, y=390, width=80, height=30)

            close =Button(ClassTreeWindow_CompensatoryHolidays_Window, text="取消", command=close_ClassTreeWindow)
            close.place(x=240, y=390, width=80, height=30)
            
            save_btn.bind("<FocusIn>", focus_see_)
            save_btn.bind("<FocusOut>", focus_see_)
            close.bind("<FocusIn>", focus_see_)
            close.bind("<FocusOut>", focus_see_)

            ClassTreeWindow_CompensatoryHolidays_Window.iconbitmap(str(res_icon_folder)+'/icon.ico')
            ClassTreeWindow_CompensatoryHolidays_Window.wait_window(ClassTreeWindow_CompensatoryHolidays_Window)
            return rt

        def close_CompensatoryHolidaysWindow():
            CompensatoryHolidaysWindow.destroy()
            ClassTreeWindow.focus_set()
        def close_CompensatoryHolidaysWindow_(nothing):
            close_CompensatoryHolidaysWindow()
        CompensatoryHolidaysWindow=Toplevel(ClassTreeWindow)
        #CompensatoryHolidaysWindow.wm_transient(ClassTreeWindow)
        CompensatoryHolidaysWindow.title("调休表")
        # 设置窗口大小、居中
        width = 630
        height = 340
        screenwidth = CompensatoryHolidaysWindow.winfo_screenwidth()
        screenheight = CompensatoryHolidaysWindow.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        CompensatoryHolidaysWindow.geometry(geometry)
        CompensatoryHolidaysWindow.resizable(width=False, height=False)
        CompensatoryHolidaysWindow.focus()
        CompensatoryHolidaysWindow.bind('<Escape>',close_CompensatoryHolidaysWindow_)
        CompensatoryHolidaysWindow.protocol("WM_DELETE_WINDOW", close_CompensatoryHolidaysWindow)

        # 表头字段 表头宽度

        Scrollbary=Scrollbar(CompensatoryHolidaysWindow,orient='vertical',)
        Scrollbary.place(x=470,y=20,width=20,height=300)
        Tree_CompensatoryHolidaysWindow_tree = Treeview(CompensatoryHolidaysWindow,yscrollcommand=Scrollbary.set ,show="headings",columns=('title','date','to'))
        Tree_CompensatoryHolidaysWindow_tree.place(x=20, y=20, width=450, height=300)
        Scrollbary.config(command=Tree_CompensatoryHolidaysWindow_tree.yview)


        Tree_CompensatoryHolidaysWindow_tree.heading('title', text="标题", anchor='center')
        Tree_CompensatoryHolidaysWindow_tree.column('title', width=150, stretch=False,anchor='w')

        Tree_CompensatoryHolidaysWindow_tree.heading('date', text="调休日期", anchor='center')
        Tree_CompensatoryHolidaysWindow_tree.column('date', width=150, stretch=False,anchor='w')

        Tree_CompensatoryHolidaysWindow_tree.heading('to', text="调休对应工作日", anchor='center')
        Tree_CompensatoryHolidaysWindow_tree.column('to', width=120, stretch=False,anchor='w')

        def New_CompensatoryHolidays():
            name=InputBox(title='新建调休日的标题',text='标题:',parent=CompensatoryHolidaysWindow,canspecialtext=False,canspace=True)
            if name==None: return 0
            try: int(name.replace(' ','')) 
            except: pass
            else:
                Message_Box(CompensatoryHolidaysWindow,'新建调休日失败.\n错误代码: 该标题不符合命名规范,禁止纯数字的标题.','错误',icon='error')
                return 0
            if Compensatory_Holidays_ini.has_section(name)==True: 
                Message_Box(CompensatoryHolidaysWindow,'新建调休日失败.\n错误代码: 该标题已存在,禁止重复使用同一标题.','错误',icon='error')
                return 0
            date_and_more=edit_CompensatoryHolidays(CompensatoryHolidaysWindow)
            if date_and_more==None: return 0
            
            try:
                Compensatory_Holidays_ini.add_section(name)
                Compensatory_Holidays_ini.set(name, 'date',json.dumps(date_and_more[0],ensure_ascii=False))
                Compensatory_Holidays_ini.set(name, 'Compensatory',date_and_more[1])
                with open('DATA/CompensatoryHolidays/CompensatoryHolidays.ini', 'w',encoding='utf-8') as configfile:#不要忘了写入!!
                    Compensatory_Holidays_ini.write(configfile)
            except Exception as error:
                Message_Box(CompensatoryHolidaysWindow, '写入调休日文件失败.\n错误代码: ' + str(error), '错误', icon='error')
                return 0

            for child in Tree_CompensatoryHolidaysWindow_tree.get_children():#更新表格
                Tree_CompensatoryHolidaysWindow_tree.delete(child)
            for sec in Compensatory_Holidays_ini.sections():
                Tree_CompensatoryHolidaysWindow_tree.insert('',index='end',values=(sec,'%s/%s/%s'%(str(json.loads(Compensatory_Holidays_ini[sec]['date'])[0]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[1]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[2])),Compensatory_Holidays_ini[sec]['Compensatory'].replace('Monday','周一').replace('Tuesday','周二').replace('Wednesday','周三').replace('Thursday','周四').replace('Friday','周五')))
        Button_CompensatoryHolidaysWindow_tree_new=Button(CompensatoryHolidaysWindow,text='新建',command=New_CompensatoryHolidays)
        Button_CompensatoryHolidaysWindow_tree_new.place(x=510,y=20,width=100,height=40)

        def Revise_CompensatoryHolidays():
            try:
                selected_item = Tree_CompensatoryHolidaysWindow_tree.selection()[0]
                values = Tree_CompensatoryHolidaysWindow_tree.item(selected_item)
            except:
                win32api.MessageBeep()
                Tree_CompensatoryHolidaysWindow_tree.focus_set()
                return 0
            name=values['values'][0]
            if name==None: return 0
            date_and_more=edit_CompensatoryHolidays(CompensatoryHolidaysWindow,title='编辑调休日 '+str(name)+' 的属性',normalday=values['values'][2])
            if date_and_more==None: return 0
            
            try:
                Compensatory_Holidays_ini.remove_section(name)
                Compensatory_Holidays_ini.add_section(name)
                Compensatory_Holidays_ini.set(name, 'date',json.dumps(date_and_more[0],ensure_ascii=False))
                Compensatory_Holidays_ini.set(name, 'Compensatory',date_and_more[1])
                with open('DATA/CompensatoryHolidays/CompensatoryHolidays.ini', 'w',encoding='utf-8') as configfile:#不要忘了写入!!
                    Compensatory_Holidays_ini.write(configfile)
            except Exception as error:
                Message_Box(CompensatoryHolidaysWindow, '写入调休表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
                return 0

            for child in Tree_CompensatoryHolidaysWindow_tree.get_children():#更新表格
                Tree_CompensatoryHolidaysWindow_tree.delete(child)
            for sec in Compensatory_Holidays_ini.sections():
                Tree_CompensatoryHolidaysWindow_tree.insert('',index='end',values=(sec,'%s/%s/%s'%(str(json.loads(Compensatory_Holidays_ini[sec]['date'])[0]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[1]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[2])),Compensatory_Holidays_ini[sec]['Compensatory'].replace('Monday','周一').replace('Tuesday','周二').replace('Wednesday','周三').replace('Thursday','周四').replace('Friday','周五')))
        Button_CompensatoryHolidaysWindow_tree_revise=Button(CompensatoryHolidaysWindow,text='编辑',command=Revise_CompensatoryHolidays)
        Button_CompensatoryHolidaysWindow_tree_revise.place(x=510,y=80,width=100,height=40)

        def Delete_CompensatoryHolidays():
            try:
                selected_item = Tree_CompensatoryHolidaysWindow_tree.selection()[0]
                values = Tree_CompensatoryHolidaysWindow_tree.item(selected_item)
            except:
                win32api.MessageBeep()
                Tree_CompensatoryHolidaysWindow_tree.focus_set()
                return 0
            if Message_Box(CompensatoryHolidaysWindow,'删除调休日无法撤回,是否继续?','疑问',icon='question',buttonmode=2)!=True: return 0
            try:
                Compensatory_Holidays_ini.remove_section(values['values'][0])
                with open('DATA/CompensatoryHolidays/CompensatoryHolidays.ini', 'w',encoding='utf-8') as configfile:#不要忘了写入!!
                    Compensatory_Holidays_ini.write(configfile)
            except Exception as error:
                Message_Box(CompensatoryHolidaysWindow, '写入调休日文件失败.\n错误代码: ' + str(error), '错误',icon='error')
                return 0

            for child in Tree_CompensatoryHolidaysWindow_tree.get_children():#更新表格
                Tree_CompensatoryHolidaysWindow_tree.delete(child)
            for sec in Compensatory_Holidays_ini.sections():
                Tree_CompensatoryHolidaysWindow_tree.insert('',index='end',values=(sec,'%s/%s/%s'%(str(json.loads(Compensatory_Holidays_ini[sec]['date'])[0]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[1]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[2])),Compensatory_Holidays_ini[sec]['Compensatory'].replace('Monday','周一').replace('Tuesday','周二').replace('Wednesday','周三').replace('Thursday','周四').replace('Friday','周五')))
        Button_CompensatoryHolidaysWindow_tree_delete=Button(CompensatoryHolidaysWindow,text='删除',command=Delete_CompensatoryHolidays)
        Button_CompensatoryHolidaysWindow_tree_delete.place(x=510,y=140,width=100,height=40)

        def Import_CompensatoryHolidays():
            file=tkinter.filedialog.askopenfilename(parent=CompensatoryHolidaysWindow,title='导入调休表文件',filetypes=[('调休表文件','.ini')])
            if file=='' or file==None: return 0
            if Message_Box(CompensatoryHolidaysWindow,'导入调休表文件将会覆盖原调休表文件,且无法撤回!','警告',buttonmode=2,icon='warning')!=True: return 0
            try:
                remove_BOM('DATA/CompensatoryHolidays/CompensatoryHolidays.ini')
                Compensatory_Holidays_ini = configparser.ConfigParser()
                Compensatory_Holidays_ini.read(file,encoding='utf-8')
            except Exception as error:
                Message_Box(CompensatoryHolidaysWindow, '导入调休表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
                return 0
            else:
                try:
                    with open('DATA/CompensatoryHolidays/CompensatoryHolidays.ini','w',encoding='utf-8') as f:#调休进行覆写
                        file_io=open(file,'r',encoding='utf-8')
                        f.write(file_io.read())
                        file_io.close()
                except Exception as error:
                    Message_Box(CompensatoryHolidaysWindow, '导入调休表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
                    return 0
                else:
                    for child in Tree_CompensatoryHolidaysWindow_tree.get_children():#更新表格
                        Tree_CompensatoryHolidaysWindow_tree.delete(child)
                    for sec in Compensatory_Holidays_ini.sections():
                        Tree_CompensatoryHolidaysWindow_tree.insert('',index='end',values=(sec,'%s/%s/%s'%(str(json.loads(Compensatory_Holidays_ini[sec]['date'])[0]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[1]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[2])),Compensatory_Holidays_ini[sec]['Compensatory'].replace('Monday','周一').replace('Tuesday','周二').replace('Wednesday','周三').replace('Thursday','周四').replace('Friday','周五')))
                    Message_Box(CompensatoryHolidaysWindow, '导入调休表文件成功.\n文件路径: \n'+file, '信息', icon='info')
        Button_CompensatoryHolidaysWindow_tree_import=Button(CompensatoryHolidaysWindow,text='导入',command=Import_CompensatoryHolidays)
        Button_CompensatoryHolidaysWindow_tree_import.place(x=510,y=200,width=100,height=40)

        def Output_CompensatoryHolidays():
            file=tkinter.filedialog.asksaveasfilename(parent=CompensatoryHolidaysWindow,title='导出调休表文件',filetypes=[('调休表文件','.ini')],defaultextension='*.ini')
            if file=='' or file==None: return 0
            try:
                with open('DATA/CompensatoryHolidays/CompensatoryHolidays.ini','r',encoding='utf-8') as f:
                    wt=f.read()
                remove_BOM(file)
                with open(file,'w',encoding='utf-8') as f:
                    f.write(wt)
            except Exception as error:
                Message_Box(ClassTreeWindow, '导出课程表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
                return 0
            else:
                Message_Box(ClassTreeWindow, '导出课程表文件成功.\n文件路径: \n'+file, '信息', icon='info')

        Button_CompensatoryHolidaysWindow_tree_output=Button(CompensatoryHolidaysWindow,text='导出',command=Output_CompensatoryHolidays)
        Button_CompensatoryHolidaysWindow_tree_output.place(x=510,y=260,width=100,height=40)

        if Permissions=='STUDENT':
            Button_CompensatoryHolidaysWindow_tree_new['state']='disabled'
            Button_CompensatoryHolidaysWindow_tree_revise['state']='disabled'
            Button_CompensatoryHolidaysWindow_tree_delete['state']='disabled'
            Button_CompensatoryHolidaysWindow_tree_import['state']='disabled'
            Button_CompensatoryHolidaysWindow_tree_output['state']='disabled'



        try:
            remove_BOM('DATA/CompensatoryHolidays/CompensatoryHolidays.ini')
            Compensatory_Holidays_ini = configparser.ConfigParser()
            Compensatory_Holidays_ini.read('DATA/CompensatoryHolidays/CompensatoryHolidays.ini',encoding='utf-8')
        except Exception as error:
            Message_Box(CompensatoryHolidaysWindow, '读取调休表文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return 0

        for child in Tree_CompensatoryHolidaysWindow_tree.get_children():
            Tree_CompensatoryHolidaysWindow_tree.delete(child)
        for sec in Compensatory_Holidays_ini.sections():
            Tree_CompensatoryHolidaysWindow_tree.insert('',index='end',values=(sec,'%s/%s/%s'%(str(json.loads(Compensatory_Holidays_ini[sec]['date'])[0]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[1]),str(json.loads(Compensatory_Holidays_ini[sec]['date'])[2])),Compensatory_Holidays_ini[sec]['Compensatory'].replace('Monday','周一').replace('Tuesday','周二').replace('Wednesday','周三').replace('Thursday','周四').replace('Friday','周五')))

        CompensatoryHolidaysWindow.iconbitmap(str(res_icon_folder)+'/icon.ico')
        CompensatoryHolidaysWindow.mainloop()


        #edit_CompensatoryHolidays(CompensatoryHolidaysWindow)
    Button_ClassTreeWindow_compensatoryholidays = Button(ClassTreeWindow, text="调休表", command=set_CompensatoryHolidays)
    Button_ClassTreeWindow_compensatoryholidays.place(x=800, y=380, width=100, height=40)

    if Permissions=='STUDENT':
        Button_ClassTreeWindow_newclass['state']='disabled'
        Button_ClassTreeWindow_temporaryclass['state']='disabled'
        Button_ClassTreeWindow_popclass['state']='disabled'
        Button_ClassTreeWindow_importclass['state']='disabled'
        Button_ClassTreeWindow_outputclass['state']='disabled'
        Button_ClassTreeWindow_makexlsx['state']='disabled'

    ClassTreeWindow.iconbitmap(str(res_icon_folder)+'/icon.ico')
    ClassTreeWindow.mainloop()
Button_root_classtree=Button(root,text='课程表',command=Button_root_classtree_open_ClassTreeWindow)
Button_root_classtree.place(x=150,y=0,width=150,height=40)


def Thread_check_istime_to_playsound():
    global running_threading,class_on_notice
    while True:
        if running_threading==False: break
        try:#优先考虑调休日
            remove_BOM('DATA/CompensatoryHolidays/CompensatoryHolidays.ini')
            CompensatoryHolidays_ini = configparser.ConfigParser()
            CompensatoryHolidays_ini.read('DATA/CompensatoryHolidays/CompensatoryHolidays.ini',encoding='utf-8')
        except Exception as error:
            #Message_Box(root, '读取调休日文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            continue

        try:#正常课程表
            with open('DATA/ClassTree/ClassTree.json','r',encoding='utf-8') as f:
                Class_Tree = json.loads(f.read())
        except Exception as error:
            #Message_Box(root, '读取课程表文件失败.\n错误代码: ' + str(error), '错误',icon='error')
            continue

        weekday=time.strftime(r'%A')  #如果调休文件为空,则weekday=现在时间的星期
        for sec in CompensatoryHolidays_ini.sections():#调休文件读入
            date=json.loads(CompensatoryHolidays_ini[sec]['date'])
            if int(time.strftime(r'%Y'))==date[0] and int(time.strftime(r'%m'))==date[1] and int(time.strftime(r'%d'))==date[2]:
                weekday=CompensatoryHolidays_ini[sec]['Compensatory']
            else:
                weekday=time.strftime(r'%A')

        if weekday=='Saturday' or weekday=='Sunday':#判断周末(要放在调休文件读入后面)
            for i in range(61):#61
                if running_threading==False:
                    break
            continue


        for item in sort_dict_by_time(Class_Tree):#上课,往右翻,还有一个附加条件->    ->    ->
            if int(Class_Tree[item]['Time'][0][0])==int(time.strftime(r'%H')) and int(Class_Tree[item]['Time'][0][1])==int(time.strftime(r'%M')) and class_on_notice==True:
                win32api.Beep(659,1025);win32api.Beep(523,875);win32api.Beep(587,925);win32api.Beep(392,1250);win32api.Sleep(725);win32api.Beep(392,950);win32api.Beep(587,875);win32api.Beep(659,1025);win32api.Beep(523,800)
                try:
                    Balloon_Box(root,title='上课: %s'%(Class_Tree[item][weekday]),text='🖤再不回到座位坐好就要被罚抄课文了哦~🖤',staytime=0)
                except:
                    pass
        for item in sort_dict_by_time(Class_Tree):#下课,往右翻,还有一个附加条件->    ->    ->
            if int(Class_Tree[item]['Time'][1][0])==int(time.strftime(r'%H')) and int(Class_Tree[item]['Time'][1][1])==int(time.strftime(r'%M')) and class_on_notice==True:
                try:
                    Balloon_Box(root,title='下课: %s'%(Class_Tree[item][weekday]),text='🖤再不下课电脑就要强制关机了哦~🖤',staytime=0)
                except:
                    pass
        for i in range(6):#61
            print(i)
            if running_threading==False:
                break
            time.sleep(1)
Thread_check_istime_to_playsound=threading.Thread(target=Thread_check_istime_to_playsound,daemon=True)
Thread_check_istime_to_playsound.start()






def sort_dict_by_date(data):
    """
    对字典中的日期进行排序，返回一个包含所有经过排序的键的列表。
    :param data: 字典，格式为 {'标题':{'Date':[年，月，日]}
    :return: 包含所有经过排序的键的列表
    """
    sorted_keys = sorted(data.keys(), key=lambda x: data[x]['Date'])
    return sorted_keys
def Button_root_todotree_open_Edit_todo_Window():
    def close_wait_todo_window():
        wait_todo_window.destroy()
        root.focus_set()
    def close_wait_todo_window_(nothing):
        close_wait_todo_window()
    wait_todo_window=Toplevel(root)
    wait_todo_window.title("待办事务")
    width = 640
    height = 440
    screenwidth = wait_todo_window.winfo_screenwidth()
    screenheight = wait_todo_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    wait_todo_window.geometry(geometry)
    wait_todo_window.resizable(width=False, height=False)
    wait_todo_window.bind('<Escape>',close_wait_todo_window_)
    wait_todo_window.protocol("WM_DELETE_WINDOW", close_wait_todo_window)
    wait_todo_window.focus()

    yscroebar=Scrollbar(wait_todo_window,orient='vertical')
    yscroebar.place(x=480,y=20,width=20,height=400)

    tree_wait_todo_window_waittodo = Treeview(wait_todo_window, show="headings", columns=('title','date'),yscrollcommand=yscroebar.set)
    tree_wait_todo_window_waittodo.place(x=20,y=20,width=460,height=400)

    yscroebar.config(command=tree_wait_todo_window_waittodo.yview)

    tree_wait_todo_window_waittodo.heading('title', text="内容", anchor='center')
    tree_wait_todo_window_waittodo.column('title', width=300, stretch=False,anchor='w')

    tree_wait_todo_window_waittodo.heading('date', text="日期", anchor='center')
    tree_wait_todo_window_waittodo.column('date', width=120, stretch=False,anchor='w')


    def COMMAND_button_wait_todo_window_new():
        nonlocal Todo_Dictionary
        title_todo=InputBox(parent=wait_todo_window,title='新建待办事务的内容',text='内容:',canspecialtext=False,canspace=True)
        if title_todo==None: return
        try: int(title_todo.replace(' ','')) 
        except: pass
        else:
            Message_Box(wait_todo_window,'新建待办事务失败.\n错误代码: 命名错误:该内容不符合命名规范,禁止纯数字的内容.','错误',icon='error')
            return
        if title_todo in Todo_Dictionary:
            Message_Box(wait_todo_window,'新建待办事务失败.\n错误代码: 查重错误:该内容已存在,禁止重复使用同一内容.','错误',icon='error')
            return
        date_todo=InputDate(wait_todo_window,title='新建待办事务的日期')
        if date_todo==None: return

        todo_ini[title_todo]={
            'date':json.dumps(date_todo),
        }#这里是新建,也可以用add+set方法,图省事.
        try:
            with open('DATA/Todo/Todo.ini','w',encoding='utf-8') as f:
                todo_ini.write(f)
        except Exception as error:
            Message_Box(wait_todo_window, '写入待办事务文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return

        Todo_Dictionary={}
        for sec in todo_ini.sections():
            Todo_Dictionary[sec]={'Date':json.loads(todo_ini[sec]['date']),}

        for child in tree_wait_todo_window_waittodo.get_children():
            tree_wait_todo_window_waittodo.delete(child)
        for i in sort_dict_by_date(Todo_Dictionary):
            tree_wait_todo_window_waittodo.insert('', values=(i,'%s/%s/%s'%(str(Todo_Dictionary[i]['Date'][0]),str(Todo_Dictionary[i]['Date'][1]),str(Todo_Dictionary[i]['Date'][2]))), index='end')



        for sec in todo_ini.sections():
            Todo_Dictionary[sec]={'Date':json.loads(todo_ini[sec]['date']),}
        for child in tree_wait_todo_window_waittodo.get_children():
            tree_wait_todo_window_waittodo.delete(child)
        
        for i in sort_dict_by_date(Todo_Dictionary):
            tree_wait_todo_window_waittodo.insert('', values=(i,'%s/%s/%s'%(str(Todo_Dictionary[i]['Date'][0]),str(Todo_Dictionary[i]['Date'][1]),str(Todo_Dictionary[i]['Date'][2]))), index='end')

        Listbox_root_todo.delete(0,'end')
        for i in sort_dict_by_date_is_systemdate(Todo_Dictionary):
            Listbox_root_todo.insert('end',i)
    button_wait_todo_window_new = Button(wait_todo_window, text="新建",command=COMMAND_button_wait_todo_window_new )
    button_wait_todo_window_new.place(x=520, y=20, width=100, height=40)

    '''
    def COMMAND_button_wait_todo_window_edit():
        nonlocal Todo_Dictionary
        try:
            selected_item = tree_wait_todo_window_waittodo.selection()[0]
            values = tree_wait_todo_window_waittodo.item(selected_item)
        except:
            win32api.MessageBeep()
            tree_wait_todo_window_waittodo.focus_set()
            return 0
        text_todo=InputText(parent=wait_todo_window,title='编辑待办事务的内容',)
        if text_todo==None: return
        else: text_todo=text_todo.replace('\n',r'\n')
        todo_ini.set(values['values'][0],'text',text_todo)#set方法,字典写入应该也可以
        try:
            with open('DATA/Todo/Todo.ini','w',encoding='utf-8') as f:
                todo_ini.write(f)
        except Exception as error:
            Message_Box(wait_todo_window, '写入待办事务文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return

        for child in tree_wait_todo_window_waittodo.get_children():
            tree_wait_todo_window_waittodo.delete(child)
        for i in sort_dict_by_date(Todo_Dictionary):
            tree_wait_todo_window_waittodo.insert('', values=(i,'%s/%s/%s'%(str(Todo_Dictionary[i]['Date'][0]),str(Todo_Dictionary[i]['Date'][1]),str(Todo_Dictionary[i]['Date'][2]))), index='end')
        text_wait_todo_window_waittodo['state']='normal'
        text_wait_todo_window_waittodo.delete(0.0,'end')
        text_wait_todo_window_waittodo['state']='disabled'
    button_wait_todo_window_edit = Button(wait_todo_window, text="编辑",command=COMMAND_button_wait_todo_window_edit )
    button_wait_todo_window_edit.place(x=500, y=80, width=100, height=40)
    '''

    def COMMAND_button_wait_todo_window_delete():
        try:
            selected_item = tree_wait_todo_window_waittodo.selection()[0]
            values = tree_wait_todo_window_waittodo.item(selected_item)
        except:
            win32api.MessageBeep()
            tree_wait_todo_window_waittodo.focus_set()
            return
        if Message_Box(wait_todo_window,'删除待办事务无法撤回,是否继续?','疑问',icon='question',buttonmode=2)!=True: return 0
        todo_ini.remove_section(values['values'][0])

        try:
            with open('DATA/Todo/Todo.ini','w',encoding='utf-8') as f:
                todo_ini.write(f)
        except Exception as error:
            Message_Box(wait_todo_window, '写入待办事务文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return
        
        Todo_Dictionary={}
        for sec in todo_ini.sections():
            Todo_Dictionary[sec]={'Date':json.loads(todo_ini[sec]['date']),}

        for child in tree_wait_todo_window_waittodo.get_children():
            tree_wait_todo_window_waittodo.delete(child)
        for i in sort_dict_by_date(Todo_Dictionary):
            tree_wait_todo_window_waittodo.insert('', values=(i,'%s/%s/%s'%(str(Todo_Dictionary[i]['Date'][0]),str(Todo_Dictionary[i]['Date'][1]),str(Todo_Dictionary[i]['Date'][2]))), index='end')


        Listbox_root_todo.delete(0,'end')
        for i in sort_dict_by_date_is_systemdate(Todo_Dictionary):
            Listbox_root_todo.insert('end',i)
    button_wait_todo_window_delete = Button(wait_todo_window, text="删除",command=COMMAND_button_wait_todo_window_delete )
    button_wait_todo_window_delete.place(x=520, y=80, width=100, height=40)

    try:
        remove_BOM('DATA/Todo/Todo.ini')
        todo_ini=configparser.ConfigParser()
        todo_ini.read('DATA/Todo/Todo.ini',encoding='utf-8')

        Todo_Dictionary={}
        for sec in todo_ini.sections():
            Todo_Dictionary[sec]={'Date':json.loads(todo_ini[sec]['date']),}

        for child in tree_wait_todo_window_waittodo.get_children():
            tree_wait_todo_window_waittodo.delete(child)
        for i in sort_dict_by_date(Todo_Dictionary):
            tree_wait_todo_window_waittodo.insert('', values=(i,'%s/%s/%s'%(str(Todo_Dictionary[i]['Date'][0]),str(Todo_Dictionary[i]['Date'][1]),str(Todo_Dictionary[i]['Date'][2]))), index='end')
    except Exception as error:
        Message_Box(wait_todo_window, '读取待办事务文件失败.\n错误代码: ' + str(error), '错误', icon='error')
        return

    if Permissions=='STUDENT':
        button_wait_todo_window_new['state']='disabled'
        button_wait_todo_window_delete['state']='disabled'

    wait_todo_window.iconbitmap(str(res_icon_folder)+'/icon.ico')
    wait_todo_window.mainloop()
Button_root_todo=Button(root,text='待办事务',command=Button_root_todotree_open_Edit_todo_Window)
Button_root_todo.place(x=150,y=40,width=150,height=40)










def COMMAND_Button_root_randomcaller():
    with open('DATA/RandomCaller/settings.json','r',encoding='utf-8') as f:
        RandomCaller_settings=json.loads(f.read())
    with open('DATA/RandomCaller/names.json','r',encoding='utf-8') as f:
        RandomCaller_names=json.loads(f.read())

    def close_root_Button_Randomcaller_window():
        root_Button_Randomcaller_window.destroy()
        root.focus_set()

    root_Button_Randomcaller_window=Toplevel(root)
    width = 800
    height = 600
    screenwidth =root_Button_Randomcaller_window.winfo_screenwidth()
    screenheight = root_Button_Randomcaller_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root_Button_Randomcaller_window.geometry(geometry)
    root_Button_Randomcaller_window.resizable(width=False, height=False)
    root_Button_Randomcaller_window.attributes('-topmost', 'true')
    root_Button_Randomcaller_window.title('随机点名器')
    root_Button_Randomcaller_window.protocol("WM_DELETE_WINDOW", close_root_Button_Randomcaller_window)
    root_Button_Randomcaller_window.iconbitmap(str(res_icon_folder)+'/icon.ico')
    list_=[]
    checklist=False
    n=1
    def r():
        nonlocal checklist,list_,n
        mix.config(state='readonly')
        maxx.config(state='readonly')
        try:
            int(maxx.get())
            int(mix.get())
        except:
            text['text']='错误'
            return 0
        if checklist==False:
            rg=int(maxx.get())
            for i in range(rg):
                list_.append(n)
                n=n+1
            
            checklist=True
        try:
            btn.config(state=DISABLED)
            random.randint(int(mix.get()),int(maxx.get()))
        except:
            text['text']='错误'
            btn.config(state='normal')
            return 0
        else:
            btn['state']='disabled'
            for i in range(50):
                try:
                    text['text']=random.randint(int(mix.get()),int(maxx.get()))
                    time.sleep(0.01)
                    root_Button_Randomcaller_window.update()
                except:
                    text['text']='错误'
                    return 0
            if len(list_)==0:
                n=int(mix.get())
                for i in range(int(maxx.get())):
                    list_.append(n)
                    n=n+1
            wdel=random.randint(0,len(list_)-1)

            if RandomCaller_settings['DisplayMode']=='Name':
                text['text']=RandomCaller_names[int(list_[wdel])-1]
            else:
                text['text']=list_[wdel]
                
            del list_[wdel]
            cout.config(state='normal')
            cout.delete(0,END)
            cout.insert(0,int(maxx.get())-len(list_))
            cout.config(state='readonly')
            btn.config(state='normal')
            root_Button_Randomcaller_window.update()



    text = tkinter.Label(root_Button_Randomcaller_window,text='',font=('微软雅黑', 170),state='normal')
    text.place(x=20, y=20, width=760, height=410)

    mix= Entry(root_Button_Randomcaller_window, justify='center')
    mix.insert(0,str(RandomCaller_settings['DefaultFromNumber']))
    mix.place(x=20, y=450, width=240, height=30)

    cout= Entry(root_Button_Randomcaller_window,justify='center',)
    cout.insert(0,'0')
    cout.config(state='readonly')
    cout.place(x=280, y=450, width=240, height=30)

    maxx = Entry(root_Button_Randomcaller_window,justify='center')
    maxx.insert(0,str(RandomCaller_settings['DefaultToNumber']))
    maxx.place(x=540, y=450, width=240, height=30)

    btn = Button(root_Button_Randomcaller_window, text="点名",command=r,)
    btn.place(x=20, y=500, width=760, height=80)
    btn.focus()



    root_Button_Randomcaller_window.mainloop()


    root.after(0,COMMAND_Button_root_randomcaller)
Button_root_randomcaller=Button(root,text='随机点名器',command=COMMAND_Button_root_randomcaller,)
Button_root_randomcaller.place(x=0,y=80,width=150,height=40)








def COMMAND_Button_root_timer():
    if  os.path.exists('Timer.pyw')==True:
        subprocess.Popen(['Timer.pyw'], shell=True)
    elif os.path.exists('Timer.exe')==True:
        subprocess.Popen(['Timer.exe'], shell=True)
    else:
        Message_Box(root,'无法运行Timer.pyw或Timer.exe,请检查文件是否存在.','错误',icon='error')
        return
Button_root_timer=Button(root,text='计时器',command=COMMAND_Button_root_timer)
Button_root_timer.place(x=150,y=80,width=150,height=40)







def COMMAND_Button_root_Keymapping():
    def get_window_handle_from_point(x, y):
        hwnd = win32gui.WindowFromPoint((x, y))
        return hwnd
    def send_key_to_window(hwnd, key):
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.05)
        keyboard.press(key)
        time.sleep(0.02)
        keyboard.release(key)


    def question_root_Button_PPT_showing_window():
        Frame_exit.place(x=0,y=0,width=300,height=100)
    def close_and_exit_PPT_showing_window():
        root_Button_PPT_showing_window.destroy()

    def left():
        hwnd=get_window_handle_from_point(int(screenwidth/2),int(screenheight/2))
        send_key_to_window(hwnd,'left')
    def right():
        hwnd=get_window_handle_from_point(int(screenwidth/2),int(screenheight/2))
        send_key_to_window(hwnd,'right')
    def stop():
        hwnd=get_window_handle_from_point(int(screenwidth/2),int(screenheight/2))
        send_key_to_window(hwnd,'esc')
    root_Button_PPT_showing_window=Toplevel(root)
    width = 300
    height = 100
    screenwidth = root_Button_PPT_showing_window.winfo_screenwidth()
    screenheight = root_Button_PPT_showing_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, 0, (screenheight - height - 60))
    root_Button_PPT_showing_window.geometry(geometry)
    root_Button_PPT_showing_window.resizable(width=False, height=False)
    root_Button_PPT_showing_window.title("键盘映射")
    root_Button_PPT_showing_window.attributes('-topmost', 'true')
    root_Button_PPT_showing_window.attributes("-toolwindow", 1)
    root_Button_PPT_showing_window.protocol("WM_DELETE_WINDOW", question_root_Button_PPT_showing_window)
    root_Button_PPT_showing_window.focus()

    Button_left=Button(root_Button_PPT_showing_window,text='<',command=left)
    Button_left.place(x=0,y=0,width=100,height=100)

    Button_right=Button(root_Button_PPT_showing_window,text='>',command=right)
    Button_right.place(x=100,y=0,width=100,height=100)

    Button_stop=Button(root_Button_PPT_showing_window,text='X',command=stop)
    Button_stop.place(x=200,y=0,width=100,height=100)

    Frame_exit=tkinter.Frame(root_Button_PPT_showing_window,bd=0)
    Frame_exit.place(x=0,y=32000,width=300,height=100)

    Button_exit=Button(Frame_exit,text='关闭窗口',command=close_and_exit_PPT_showing_window)
    Button_exit.place(x=0,y=0,width=150,height=100)

    def showwindow_root_Button_PPT_showing_window():
        Frame_exit.place(x=0,y=32000,width=300,height=100)
    
    Button_donotexit=Button(Frame_exit,text='只是误触了',command=showwindow_root_Button_PPT_showing_window)
    Button_donotexit.place(x=150,y=0,width=150,height=100)

    root_Button_PPT_showing_window.iconbitmap(str(res_icon_folder)+'/icon.ico')

    root_Button_PPT_showing_window.mainloop()
Button_root_Keymapping=Button(root,text='键盘映射',command=COMMAND_Button_root_Keymapping)
Button_root_Keymapping.place(x=0,y=120,width=150,height=40)







def COMMAND_Button_root_Clock():
    running_showtime_threading=True
    def close_Button_root_Clock_window():
        nonlocal running_showtime_threading
        running_showtime_threading=False
        Button_root_Clock_window.destroy()
        root.focus_set()
    def close_Button_root_Clock_window_(nothing):
        close_Button_root_Clock_window()
    Button_root_Clock_window=Toplevel(root)
    Button_root_Clock_window.title("时钟")
    width = 800
    height = 500
    screenwidth = Button_root_Clock_window.winfo_screenwidth()
    screenheight = Button_root_Clock_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, int((screenwidth - width)/2), int((screenheight - height)/2))
    Button_root_Clock_window.geometry(geometry)
    Button_root_Clock_window.attributes('-topmost', 'true')
    Button_root_Clock_window.protocol("WM_DELETE_WINDOW",close_Button_root_Clock_window)
    Button_root_Clock_window.bind('<Escape>',close_Button_root_Clock_window_)
    Button_root_Clock_window.focus()

    Button_root_Clock_window_Label=tkinter.Label(Button_root_Clock_window,text='',font='微软雅黑 110')
    Button_root_Clock_window_Label.pack(expand=True,fill='both')

    def Button_root_Clock_window_Label_SHOWTIME():
        nonlocal running_showtime_threading
        while True:
            if running_showtime_threading==False: break
            try:
                Button_root_Clock_window_Label['text']=time.strftime(r"%H:%M:%S")
                Button_root_Clock_window.update()
            except:
                break
    root.after(0,Button_root_Clock_window_Label_SHOWTIME)

    Button_root_Clock_window.iconbitmap(str(res_icon_folder)+'/icon.ico')
    Button_root_Clock_window.mainloop()
Button_root_Clock=Button(root,text='时钟',command=COMMAND_Button_root_Clock)
Button_root_Clock.place(x=150,y=120,width=150,height=40)










def COMMAND_Button_root_DisabledCompute():
    running_topmost_threading=True
    mode=InputComboBox(parent=root,title='选择禁用屏幕的解锁模式',text='模式:',value=['密码解锁','定时解锁'],default='密码解锁')
    if mode=='': return
    elif mode=='密码解锁':
        password=InputBox(parent=root,title='禁用屏幕的解锁密码',text='密码:',show='●',canspecialtext=True,canempty=False)
        if password==None: return
        try:
            int(password)
        except:
            Message_Box(root,'密码必须是数字.','错误',icon='error')
            return
        if Message_Box(root,'请确认解锁电脑的密码为: '+str(password),'疑问',icon='question',buttonmode=2)!=True: return
    elif mode=='定时解锁':
        disabled_time=InputBox(parent=root,title='禁用屏幕的时长(分钟)',text='时长:',canspecialtext=True,canempty=False)
        if disabled_time==None: return
        try:
            int(disabled_time)
        except:
            Message_Box(root,'时长只能是数字.','错误',icon='error')
            return
        if Message_Box(root,'请确认禁用屏幕的时长为: '+str(disabled_time)+' 分钟.','疑问',icon='question',buttonmode=2)!=True: return
        if Message_Box(root,'在"定时解锁"模式下,未到达解锁时间任何人都无法解锁!','警告',icon='warning',buttonmode=2)!=True: return
    
    def run_window():
        def DisabledCompute_window_destroy():
            nonlocal running_topmost_threading
            running_topmost_threading=False
            DisabledCompute_window.destroy()
        def DisabledCompute_window_pass():
            pass

        DisabledCompute_window=Toplevel(root)
        DisabledCompute_window.resizable(width=False, height=False)
        DisabledCompute_window.attributes('-topmost','true')
        DisabledCompute_window.attributes('-fullscreen',True)
        DisabledCompute_window.overrideredirect(True)
        DisabledCompute_window.protocol("WM_DELETE_WINDOW", DisabledCompute_window_pass)
        DisabledCompute_window['bg']='black'


        def Thread_topmost():
            while True:
                nonlocal running_topmost_threading
                global running_threading
                if running_topmost_threading==False or running_threading==False: break
                try:
                    DisabledCompute_window.focus_force()
                    DisabledCompute_window.attributes('-topmost',True)
                except:
                    break
        Thread_topmost_threading=threading.Thread(target=Thread_topmost,daemon=True)
        Thread_topmost_threading.start()

        if mode=='定时解锁':
            tkinter.Label(DisabledCompute_window,text='保持安静',bg='black',fg='yellow',font='微软雅黑 80').pack(fill='both',expand=True)
            DisabledCompute_window.after(int(disabled_time)*60*1000,DisabledCompute_window_destroy)

        elif mode=='密码解锁':
            password_mistake_counter=0

            Frame_PasswordKeyboard=tkinter.Frame(DisabledCompute_window,bd=0,width=240,height=360)
            Frame_PasswordKeyboard.pack(anchor='center',expand=True)

            Entry_Password=Entry(Frame_PasswordKeyboard,show='●',state='disabled',justify='center')
            Entry_Password.place(x=0,y=0,width=240,height=40)

            def COMMAND_insert_1():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','1')
                Entry_Password['state']='readonly'
            Button_1=Button(Frame_PasswordKeyboard,text='1',command=COMMAND_insert_1)
            Button_1.place(x=0,y=40,width=80,height=80)

            def COMMAND_insert_2():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','2')
                Entry_Password['state']='readonly'
            Button_2=Button(Frame_PasswordKeyboard,text='2',command=COMMAND_insert_2)
            Button_2.place(x=80,y=40,width=80,height=80)

            def COMMAND_insert_3():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','3')
                Entry_Password['state']='readonly'
            Button_3=Button(Frame_PasswordKeyboard,text='3',command=COMMAND_insert_3)
            Button_3.place(x=160,y=40,width=80,height=80)

            def COMMAND_insert_4():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','4')
                Entry_Password['state']='readonly'
            Button_4=Button(Frame_PasswordKeyboard,text='4',command=COMMAND_insert_4)
            Button_4.place(x=0,y=120,width=80,height=80)

            def COMMAND_insert_5():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','5')
                Entry_Password['state']='readonly'
            Button_5=Button(Frame_PasswordKeyboard,text='5',command=COMMAND_insert_5)
            Button_5.place(x=80,y=120,width=80,height=80)

            def COMMAND_insert_6():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','6')
                Entry_Password['state']='readonly'
            Button_6=Button(Frame_PasswordKeyboard,text='6',command=COMMAND_insert_6)
            Button_6.place(x=160,y=120,width=80,height=80)

            def COMMAND_insert_7():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','7')
                Entry_Password['state']='readonly'
            Button_7=Button(Frame_PasswordKeyboard,text='7',command=COMMAND_insert_7)
            Button_7.place(x=0,y=200,width=80,height=80)

            def COMMAND_insert_8():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','8')
                Entry_Password['state']='readonly'
            Button_8=Button(Frame_PasswordKeyboard,text='8',command=COMMAND_insert_8)
            Button_8.place(x=80,y=200,width=80,height=80)

            def COMMAND_insert_9():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','9')
                Entry_Password['state']='readonly'
            Button_9=Button(Frame_PasswordKeyboard,text='9',command=COMMAND_insert_9)
            Button_9.place(x=160,y=200,width=80,height=80)

            def COMMAND_check_password():
                nonlocal password_mistake_counter
                def Unlock_Buttons():
                    nonlocal password_mistake_counter
                    Entry_Password['state']='normal'
                    Entry_Password['show']='●'
                    Entry_Password.delete(0,'end')
                    Entry_Password['state']='readonly'
                    Button_1['state']='normal'
                    Button_2['state']='normal'
                    Button_3['state']='normal'
                    Button_4['state']='normal'
                    Button_5['state']='normal'
                    Button_6['state']='normal'
                    Button_7['state']='normal'
                    Button_8['state']='normal'
                    Button_9['state']='normal'
                    Button_0['state']='normal'
                    Button_ok['state']='normal'
                    Button_delete['state']='normal'
                
                if Entry_Password.get()==password:
                    DisabledCompute_window_destroy()
                else:
                    win32api.MessageBeep()
                    Entry_Password['state']='normal'
                    password_mistake_counter=password_mistake_counter+1
                    Entry_Password.delete(0,'end')
                    Entry_Password['state']='readonly'
                    if password_mistake_counter>=5:
                        password_mistake_counter=0
                        Entry_Password['state']='normal'
                        Entry_Password['show']=''
                        Entry_Password.delete(0,'end')
                        Entry_Password.insert(0,'密码连续错误,2分钟后再试.')
                        Entry_Password['state']='readonly'
                        Button_1['state']='disabled'
                        Button_2['state']='disabled'
                        Button_3['state']='disabled'
                        Button_4['state']='disabled'
                        Button_5['state']='disabled'
                        Button_6['state']='disabled'
                        Button_7['state']='disabled'
                        Button_8['state']='disabled'
                        Button_9['state']='disabled'
                        Button_0['state']='disabled'
                        Button_ok['state']='disabled'
                        Button_delete['state']='disabled'
                        Frame_PasswordKeyboard.after(120000,Unlock_Buttons)#120000
                        
            Button_ok=Button(Frame_PasswordKeyboard,text='确定',command=COMMAND_check_password)
            Button_ok.place(x=0,y=280,width=80,height=80)

            def COMMAND_insert_0():
                Entry_Password['state']='normal'
                Entry_Password.insert('end','0')
                Entry_Password['state']='readonly'
            Button_0=Button(Frame_PasswordKeyboard,text='0',command=COMMAND_insert_0)
            Button_0.place(x=80,y=280,width=80,height=80)

            def COMMAND_delete():
                Entry_Password['state']='normal'
                Entry_Password.delete(len(Entry_Password.get())-1)
                Entry_Password['state']='readonly'
            Button_delete=Button(Frame_PasswordKeyboard,text='删除',command=COMMAND_delete)
            Button_delete.place(x=160,y=280,width=80,height=80)

        DisabledCompute_window.mainloop()
    run_window()
Button_root_DisabledCompute=Button(root,text='禁用屏幕',command=COMMAND_Button_root_DisabledCompute)
Button_root_DisabledCompute.place(x=0,y=160,width=150,height=40)









        
tkinter.Label(root,text='今日待办:',anchor='w').place(x=0,y=200,width=300,height=30)
Listbox_root_todo=Listbox(root,selectmode='browse',activestyle='none',highlightthickness=0,bd=2,relief='groove',selectforeground='black',selectbackground='#cde8ff',selectborderwidth=0)
Listbox_root_todo.place(x=0,y=230,width=300,height=250)

def sort_dict_by_date_is_systemdate(data):
    """
    对字典中的日期进行排序，返回一个包含所有经过排序的键的列表。
    只有年，月，日都与系统时间相同才加入输出的列表里。
    :param data: 字典，格式为 {'标题':{'Date':[年，月，日]},}
    :return: 包含所有经过排序且日期与系统时间相同的键的列表
    """
    # 获取当前系统时间
    current_date = datetime.now().strftime('%Y-%m-%d').split('-')
    current_year, current_month, current_day = int(current_date[0]), int(current_date[1]), int(current_date[2])

    # 筛选出日期与系统时间相同的条目
    filtered_data = {key: value for key, value in data.items() if value['Date'] == [current_year, current_month, current_day]}

    # 对筛选后的字典进行排序
    sorted_keys = sorted(filtered_data.keys(), key=lambda x: filtered_data[x]['Date'])
    return sorted_keys

try:
    remove_BOM('DATA/Todo/Todo.ini')
    todo_ini_for_listbox=configparser.ConfigParser()
    todo_ini_for_listbox.read('DATA/Todo/Todo.ini',encoding='utf-8')

    Todo_Dictionary_for_listbox={}
    for sec in todo_ini_for_listbox.sections():
        Todo_Dictionary_for_listbox[sec]={'Date':json.loads(todo_ini_for_listbox[sec]['date']),}

    Listbox_root_todo.delete(0,'end')
    for i in sort_dict_by_date_is_systemdate(Todo_Dictionary_for_listbox):
        Listbox_root_todo.insert('end',i)
except Exception as error: Message_Box(root, '读取待办事务文件失败.\n此错误不影响程序运行,但是会导致待办事务及其附属功能无法使用.\n错误代码: ' + str(error), '错误', icon='error')
finally: todo_ini_for_listbox=None  #重置Todo_Dictionary_for_listbox

def COMMAND_Button_root_for_Listbox_todo_finish_todo():
    if Listbox_root_todo.curselection()==():
        win32api.MessageBeep()
        Listbox_root_todo.focus()
        return
    if Message_Box(parent=root,text='完成此待办将会删除此待办事务,是否继续?',title='疑问',icon='question',buttonmode=2)!=True: return
    for index in Listbox_root_todo.curselection():
        item = Listbox_root_todo.get(index)
        try:
            remove_BOM('DATA/Todo/Todo.ini')
            todo_ini_for_COMMAND_Button=configparser.ConfigParser()
            todo_ini_for_COMMAND_Button.read('DATA/Todo/Todo.ini',encoding='utf-8')
        except Exception as error:
            Message_Box(root, '读取待办事务文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return
        todo_ini_for_COMMAND_Button.remove_section(item)

        try:
            with open('DATA/Todo/Todo.ini','w',encoding='utf-8') as f:
                todo_ini_for_COMMAND_Button.write(f)
        except Exception as error:
            Message_Box(root, '写入待办事务文件失败.\n错误代码: ' + str(error), '错误', icon='error')
            return

        Todo_Dictionary={}
        for sec in todo_ini_for_COMMAND_Button.sections():
            Todo_Dictionary[sec]={'Date':json.loads(todo_ini_for_COMMAND_Button[sec]['date']),}

        Listbox_root_todo.delete(0,'end')
        for i in sort_dict_by_date_is_systemdate(Todo_Dictionary):
            Listbox_root_todo.insert('end',i)
Button_root_for_Listbox_todo_finish_todo=Button(root,text='完成此待办',command=COMMAND_Button_root_for_Listbox_todo_finish_todo)
Button_root_for_Listbox_todo_finish_todo.place(x=0,y=480,width=150,height=40)






def UpdateWidgetPermissions():
    global Permissions
    if Permissions=='STUDENT':
        Button_exit['state']='disabled'
        Checkbutton_root_class_on_notice['state']='disabled'
        Button_root_DisabledCompute['state']='disabled'
        Button_root_Permissions_set_TEACHERPermissionsPassword['state']='disabled'
        Button_root_for_Listbox_todo_finish_todo['state']='disabled'
        Button_root_randomcaller['state']='disabled'
        Button_root_Keymapping['state']='disabled'
    elif Permissions=='TEACHER':
        Button_exit['state']='normal'
        Checkbutton_root_class_on_notice['state']='normal'
        Button_root_DisabledCompute['state']='normal'
        Button_root_Permissions_set_TEACHERPermissionsPassword['state']='normal'
        Button_root_for_Listbox_todo_finish_todo['state']='normal'
        Button_root_randomcaller['state']='normal'
        Button_root_Keymapping['state']='normal'


def COMMAND_Button_root_Permissions_get_TEACHERPermissions():
    global Permissions
    if Permissions=='STUDENT':
        TEACHERPermissions_password_input=Password_Box(parent=root,title='获得教师权限',text='输入教师权限密码(默认为123456):',defaultuser='TEACHER',defaultusertuple=('TEACHER',),usernamestate='disabled')
        if TEACHERPermissions_password_input[0]==None: return
        with open('DATA/Permissions/TEACHER.json','r',encoding='utf-8') as f:
            TEACHERPermissions_password=json.loads(f.read())
        if TEACHERPermissions_password_input[1]==TEACHERPermissions_password['Password']:
            Permissions='TEACHER'
        else:
            Message_Box(root,'获得教师权限失败.\n密码错误.','错误',icon='error')
    elif Permissions=='TEACHER':
        Permissions='STUDENT'

    UpdateWidgetPermissions()
Button_root_Permissions_get_TEACHERPermissions=Button(root,text='切换权限',command=COMMAND_Button_root_Permissions_get_TEACHERPermissions)
Button_root_Permissions_get_TEACHERPermissions.place(x=150,y=480,width=150,height=40)



def COMMAND_Button_root_Permissions_set_TEACHERPermissionsPassword():
    global Permissions
    TEACHERPermissions_password_input1=Password_Box(parent=root,title='设置教师权限密码',text='输入新的教师权限密码:',defaultuser='TEACHER',defaultusertuple=('TEACHER',),usernamestate='disabled')
    if TEACHERPermissions_password_input1[0]==None: return
    TEACHERPermissions_password_input2=Password_Box(parent=root,title='设置教师权限密码',text='再次输入新的教师权限密码:',defaultuser='TEACHER',defaultusertuple=('TEACHER',),usernamestate='disabled')
    if TEACHERPermissions_password_input2[0]==None: return
    if TEACHERPermissions_password_input1!=TEACHERPermissions_password_input2:
        Message_Box(root,'设置教师权限密码失败.\n两次输入的教师权限密码不一致.','错误',icon='error')
        return
    with open('DATA/Permissions/TEACHER.json','w',encoding='utf-8') as f:
        json.dump({'Password':TEACHERPermissions_password_input1[1]},f)
    Message_Box(root,'设置教师权限密码成功.','信息',icon='info')
Button_root_Permissions_set_TEACHERPermissionsPassword=Button(root,text='设置教师权限密码',command=COMMAND_Button_root_Permissions_set_TEACHERPermissionsPassword)
Button_root_Permissions_set_TEACHERPermissionsPassword.place(x=150,y=520,width=150,height=40)











def question_exit():
    global running_threading
    Frame_question=tkinter.Frame(root)
    Frame_question.place(x=Button_exit.winfo_x(),y=Button_exit.winfo_y(),width=Button_exit.winfo_width(),height=Button_exit.winfo_height())

    def COMMAND_exit_yes():
        global running_threading
        running_threading=False
        root.destroy()
        sys.exit()
    Button_question_yes=Button(Frame_question,text='是',command=COMMAND_exit_yes)
    Button_question_yes.place(x=0,y=0,width=75,height=40)

    Button_question_no=Button(Frame_question,text='否',command=lambda: Frame_question.destroy())
    Button_question_no.place(x=75,y=0,width=75,height=40)
Button_exit=Button(root,text='退出',command=question_exit)
Button_exit.place(x=0,y=520,width=150,height=40)










def COMMAND_root_Button_about():
    def set_image(tk_label, img_path, img_size=None):
        img_open = Image.open(img_path)
        if img_size is not None:
            height, width = img_size
            img_h, img_w = img_open.size

            if img_w > width:
                size = (width, int(height * img_w / img_h))
                img_open = img_open.resize(size, Image.LANCZOS)
 
            elif img_h > height:
                size = (int(width * img_h / img_w), height)
                img_open = img_open.resize(size, Image.LANCZOS)
        img = ImageTk.PhotoImage(img_open)
        tk_label.config(image=img)
        tk_label.image = img

    
    def close_root_Button_about_window():
        root_Button_about_window.destroy()
        root.focus_set()
    def close_root_Button_about_window_(nothing):
        close_root_Button_about_window()
    
    root_Button_about_window=Toplevel(root)
    root_Button_about_window.title("关于")
    # 设置窗口大小、居中
    width = 850
    height = 500
    screenwidth = root_Button_about_window.winfo_screenwidth()
    screenheight = root_Button_about_window.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root_Button_about_window.geometry(geometry)
    root_Button_about_window.resizable(width=False, height=False)
    root_Button_about_window.bind('<Escape>',close_root_Button_about_window_)
    root_Button_about_window.protocol("WM_DELETE_WINDOW", close_root_Button_about_window)
    root_Button_about_window.focus()

    root_Button_about_window_Label_icon = tkinter.Label(root_Button_about_window,anchor="center", )
    root_Button_about_window_Label_icon.place(x=20, y=20, width=100, height=100)
    set_image(root_Button_about_window_Label_icon, str(res_icon_folder)+'/icon.ico', (100, 100))

    scrollbary_FOR_root_Button_about_window_Text_about=Scrollbar(root_Button_about_window,orient='vertical',)
    scrollbary_FOR_root_Button_about_window_Text_about.place(x=610,y=20,width=20,height=350)

    root_Button_about_window_Text_about = Text(root_Button_about_window,bd=2,relief='groove',font='TkDefaultFont',wrap='char',yscrollcommand=scrollbary_FOR_root_Button_about_window_Text_about.set)
    root_Button_about_window_Text_about.place(x=140, y=20, width=470, height=350)
    root_Button_about_window_Text_about.insert(0.0,'''Copyright © 2024 炸图监管者. All rights reserved.
■软件名: 班级辅助授课系统
■开发者: 炸图监管者
■版本号: 0.0.0.0
■版权: Copyright © 2024 炸图监管者. All rights reserved.
▲请使用官方软件(从Github下载),因为软件开源,所以下载第三方软件有安全风险.
●本软件遵守AGPL-3.0开源协议.
●联系方式仅支持E-mail/Discord/WhatsApp.
●反馈BUG请通过E-mail/Discord反馈,不建议通过Github.
●Discord/WhatsApp需要科学上网(Github可能需要),自行上网搜索方法,电子文盲别喷!''')
    root_Button_about_window_Text_about.configure(state='disabled')

    def LinkMessageBox_show_(nothing):
        Link_Message_Box(title='器官资源管理器(彩蛋)',bigtext='大脑 无响应',text='如果关闭此器官,可能会丢失记忆.',text2='等待器官响应',text1='关闭器官',parent=root_Button_about_window,defaultfocus=1,defaultreturn=None)
    root_Button_about_window_Text_about.bind('<Double-Button-1>',LinkMessageBox_show_)
    scrollbary_FOR_root_Button_about_window_Text_about.config(command=root_Button_about_window_Text_about.yview)



    scrollbary_FOR_root_Button_about_window_Text_Acknowledgments=Scrollbar(root_Button_about_window,orient='vertical',)
    scrollbary_FOR_root_Button_about_window_Text_Acknowledgments.place(x=810,y=20,width=20,height=350)

    scrollbarx_FOR_root_Button_about_window_Text_Acknowledgments=Scrollbar(root_Button_about_window,orient='horizontal')
    scrollbarx_FOR_root_Button_about_window_Text_Acknowledgments.place(x=650,y=370,width=160,height=20)

    root_Button_about_window_Text_Acknowledgments = Text(root_Button_about_window,bd=2,relief='groove',font='TkDefaultFont',wrap='none',xscrollcommand=scrollbarx_FOR_root_Button_about_window_Text_Acknowledgments.set,yscrollcommand=scrollbary_FOR_root_Button_about_window_Text_Acknowledgments.set)
    root_Button_about_window_Text_Acknowledgments.place(x=650, y=20, width=160, height=350)
    root_Button_about_window_Text_Acknowledgments.insert(0.0,'''鸣谢名单:
再次感谢各位对开发者的发电与支持!
2024年9月:
●(无)
2024年8月:
●(无)''')
    root_Button_about_window_Text_Acknowledgments.configure(state='disabled')
    scrollbary_FOR_root_Button_about_window_Text_Acknowledgments.config(command=root_Button_about_window_Text_Acknowledgments.yview)
    scrollbarx_FOR_root_Button_about_window_Text_Acknowledgments.config(command=root_Button_about_window_Text_Acknowledgments.xview)













    s1 =Style()
    s1.configure("C.TButton",foreground='#0078e4')

    def COMMAND_root_Button_about_window_Button_link_Github_openurl(): webbrowser.open("https://github.com/zhatujianguanzhe/Class-assisted_teaching_system")
    root_Button_about_window_Button_link_Github = Button(root_Button_about_window, text="Github",style="C.TButton",command=COMMAND_root_Button_about_window_Button_link_Github_openurl)
    root_Button_about_window_Button_link_Github.place(x=20, y=140, width=100, height=40)

    def COMMAND_root_Button_about_window_Button_link_index(): webbrowser.open("https://zhatujianguanzhe.github.io/IndexPage/")
    root_Button_about_window_Button_show_WhatsApp = Button(root_Button_about_window, text="官网",style="C.TButton",command=COMMAND_root_Button_about_window_Button_link_index)
    root_Button_about_window_Button_show_WhatsApp.place(x=20, y=200, width=100, height=40)

    def COMMAND_root_Button_about_window_Button_link_Email_MessageBox(): Message_Box(parent=root_Button_about_window,text='E-mail: 1323738778@QQ.com\nWhatsApp: 炸图监管者\n仅反馈BUG和提出意见,无事勿扰.',title='信息',icon='info',buttonmode=1)
    root_Button_about_window_Button_show_Email = Button(root_Button_about_window, text="E-mail/WhatsApp",style="C.TButton",command=COMMAND_root_Button_about_window_Button_link_Email_MessageBox)
    root_Button_about_window_Button_show_Email.place(x=20, y=260, width=100, height=40)

    def COMMAND_root_Button_about_window_Button_link_Bilibili_openurl(): webbrowser.open('https://space.bilibili.com/1342104465')
    root_Button_about_window_Button_link_Bilibili = Button(root_Button_about_window, text="Bilibili",style="C.TButton",command=COMMAND_root_Button_about_window_Button_link_Bilibili_openurl)
    root_Button_about_window_Button_link_Bilibili.place(x=20, y=320, width=100, height=40)

    def COMMAND_root_Button_about_window_Button_link_Discord_openurl(): webbrowser.open('https://discord.gg/zPNbV64Z')
    root_Button_about_window_Button_link_Discord = Button(root_Button_about_window, text="Discord",style="C.TButton",command=COMMAND_root_Button_about_window_Button_link_Discord_openurl)
    root_Button_about_window_Button_link_Discord.place(x=20, y=380, width=100, height=40)

    def COMMAND_root_Button_about_window_Button_link_YouTube_openurl(): webbrowser.open('https://www.youtube.com/@zhatujianguanzhe')
    root_Button_about_window_Button_link_YouTube = Button(root_Button_about_window, text="YouTube",style="C.TButton",command=COMMAND_root_Button_about_window_Button_link_YouTube_openurl)
    root_Button_about_window_Button_link_YouTube.place(x=20, y=440, width=100, height=40)
    
    def COMMAND_root_Button_about_window_Button_Message_Agreement(): win32api.MessageBox(0,'您已经同意过软件使用协议,请以首次使用软件时同意的协议为准.\n本文是直接照上文的,仅供参考,不具有任何意义!\n'+Agreement,'协议',win32con.MB_ICONINFORMATION|win32con.MB_TOPMOST|win32con.MB_TASKMODAL)
    root_Button_about_window_Button_Agreement = Button(root_Button_about_window, text="软件协议",command=COMMAND_root_Button_about_window_Button_Message_Agreement)
    root_Button_about_window_Button_Agreement.place(x=320, y=440, width=100, height=40)

    root_Button_about_window_Label_donate = tkinter.Label(root_Button_about_window,text="开发者吃饱饭才能做更好的东西,可以的话,请开发者吃个甜甜圈吧.",anchor="w", )
    root_Button_about_window_Label_donate.place(x=140, y=390, width=490, height=30)

    def COMMAND_root_Button_about_window_Button_link_Aifadian_openurl(): webbrowser.open('ifdian.net/a/zhatujianguanzhe')
    root_Button_about_window_Button_donate = Button(root_Button_about_window, text="捐助(爱发电)",style="C.TButton",command=COMMAND_root_Button_about_window_Button_link_Aifadian_openurl)
    root_Button_about_window_Button_donate.place(x=140, y=440, width=160, height=40)
    


    root_Button_about_window_Button_OK = Button(root_Button_about_window, text="确定",command=close_root_Button_about_window)
    root_Button_about_window_Button_OK.place(x=710, y=440, width=120, height=40)

    #.window_create('insert',window=---)
    root_Button_about_window.iconbitmap(str(res_icon_folder)+'/icon.ico')
    root_Button_about_window.mainloop()
root_Button_about=tkinter.Button(root,text='关于',command=COMMAND_root_Button_about,bd=0,relief='groove',activebackground='#bcdcf4',bg='#e1e1e1')
root_Button_about.place(x=0,y=560,width=300,height=40)
#<Enter>
def Enter_root_Button_about(n):
    root_Button_about['text']='关于关于一些关于关于的关于'
    root_Button_about['bg']='#d9ebf9'
root_Button_about.bind('<Enter>',Enter_root_Button_about)
def Leave_root_Button_about(n):
    root_Button_about['text']='关于'
    root_Button_about['bg']='#e1e1e1'
root_Button_about.bind('<Leave>',Leave_root_Button_about)




Button_root_show=Button(root,text='显示',command=root_show)
Button_root_show.place(x=0,y=root.winfo_height()-20,width=300,height=40)





root_hide()#一些BUG,只能这么做了
root_show()

UpdateWidgetPermissions()
root.iconbitmap(str(res_icon_folder)+'/icon.ico')
root.wait_window(root)

running_threading=False

sys.exit()