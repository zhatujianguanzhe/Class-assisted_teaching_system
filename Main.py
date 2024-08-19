# -*- coding: utf-8 -*-
from tkinter import *
import tkinter.filedialog
from tkinter.ttk import *
from PIL import Image,ImageTk
from datetime import datetime
import tkinter,tkcalendar#tkcalendar是日期控件库,推荐在清华镜像源里下载
import win32api,win32con,pywintypes,win32gui
import ctypes,os,sys,configparser,re,json,time,threading,xlwings,random,subprocess,keyboard

if getattr(sys, 'frozen', None):
    res_icon_folder = str(sys._MEIPASS)+'\\icons'
else:
    res_icon_folder = str(os.path.dirname(__file__))+'\\icons'

running_threading=True

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

def InputBox(title='', text='', parent=None, default='', canspace=False, canempty=False, canspecialtext=False):
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

    entry_filename = Entry(Input_Box_window, takefocus=True)
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
    elif icon=='error' or icon=='safe_error' or icon=='word_error_red' or icon=='word_deny':
        win32api.MessageBeep(win32con.MB_ICONERROR)
    elif icon=='warning' or icon=='safe_warning' or icon=='word_correct_orange' or icon=='modern_warning':
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

def InputText(parent,title='',):
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
        Balloon_Message_Box_Window=tkinter.Toplevel(parent)
    else:
        Balloon_Message_Box_Window=tkinter.Tk()
    Balloon_Message_Box_Window.overrideredirect(True)
    Balloon_Message_Box_Window.title("")
    width = 370
    height = 200

    Balloon_Message_Box_Window.resizable(width=False, height=False)
    Balloon_Message_Box_Window.attributes('-topmost', 'true')
    Balloon_Message_Box_Window.config(cursor='arrow')
    Balloon_Message_Box_Window.bind('<ButtonRelease-1>', lambda event: close_Balloon_Message_Box_Window())

    rtitle = tkinter.Label(Balloon_Message_Box_Window,text=title,anchor="w", font='微软雅黑 12')
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


try:
    remove_BOM('settings.ini')
    settings = configparser.ConfigParser()
    settings.read('settings.ini',encoding='utf-8')
    ctypes.windll.shcore.SetProcessDpiAwareness(int(settings.get('DPI','DPI_mode')))
except:
    if win32api.GetSystemMetrics(0)>=1920:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(0)
        except:
            pass
    else:#<1920
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass











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
draggable(root)
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
Checkbutton_root_class_on_notice=Checkbutton(root,text='上课提醒',variable=Var_class_on_notice,onvalue=True,offvalue=False,command=DO_SHOW_Var_class_on_notice)
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

#◀ ▶ ◀ ▶ ◀
    ClassTreeWindow.iconbitmap(str(res_icon_folder)+'/icon.ico')
    ClassTreeWindow.mainloop()

Button_root_classtree=Button(root,text='课程表',command=Button_root_classtree_open_ClassTreeWindow)
Button_root_classtree.place(x=150,y=0,width=150,height=40)







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

    '''
    def update_text_wait_todo_window_waittodo_(nothing):
        nonlocal todo_ini
        text_wait_todo_window_waittodo['state']='normal'
        text_wait_todo_window_waittodo.delete(0.0,'end')
        try:
            selected_item = tree_wait_todo_window_waittodo.selection()[0]
            values = tree_wait_todo_window_waittodo.item(selected_item)
        except:
            win32api.MessageBeep()
            tree_wait_todo_window_waittodo.focus_set()
            return
        text_wait_todo_window_waittodo.insert(0.0,todo_ini[values['values'][0]]['text'].replace(r'\n','\n'))
        text_wait_todo_window_waittodo['state']='disabled'
    '''#tree_wait_todo_window_waittodo.bind('<<TreeviewSelect>>',update_text_wait_todo_window_waittodo_)

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
    def topmost():
        while True:
            nonlocal top
            if top==True:
                try:
                    root_Button_PPT_showing_window.attributes('-topmost','true')
                except:
                    break
            else:
                break
    top=True
    def question_root_Button_PPT_showing_window():
        Frame_exit.place(x=0,y=0,width=300,height=100)
    def close_and_exit_PPT_showing_window():
        nonlocal top
        top=False
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
    root_Button_PPT_showing_window.title("按键映射")
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

    Button_exit=Button(Frame_exit,text='退出程序',command=close_and_exit_PPT_showing_window)
    Button_exit.place(x=0,y=0,width=150,height=100)

    def showwindow_root_Button_PPT_showing_window():
        Frame_exit.place(x=0,y=32000,width=300,height=100)
    
    Button_donotexit=Button(Frame_exit,text='只是误触了',command=showwindow_root_Button_PPT_showing_window)
    Button_donotexit.place(x=150,y=0,width=150,height=100)

    threading.Thread(target=topmost, daemon=True).start()
    root_Button_PPT_showing_window.iconbitmap(str(res_icon_folder)+'/icon.ico')

    root_Button_PPT_showing_window.mainloop()
Button_root_Keymapping=Button(root,text='键盘映射',command=COMMAND_Button_root_Keymapping)
Button_root_Keymapping.place(x=0,y=120,width=150,height=40)





        
tkinter.Label(root,text='今日待办:',anchor='w').place(x=0,y=160,width=300,height=30)
Listbox_root_todo=Listbox(root,selectmode='browse',activestyle='none',highlightthickness=0,bd=2,relief='groove',selectforeground='black',selectbackground='#cde8ff',selectborderwidth=0)
Listbox_root_todo.place(x=0,y=190,width=300,height=250)


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
Button_root_for_Listbox_todo_finish_todo.place(x=0,y=440,width=150,height=40)



#尝试读取,保证程序运行
try:#正常课程表
    with open('DATA/ClassTree/ClassTree.json','r',encoding='utf-8') as f:
        json.loads(f.read())
except Exception as error:
    Message_Box(root, '读取课程表文件失败.\n此错误不影响程序正常运行,但是会导致课程表及其附属功能无法使用.\n错误代码: ' + str(error), '错误',icon='error')


try:#读调休文件
    remove_BOM('DATA/CompensatoryHolidays/CompensatoryHolidays.ini')
    configparser.ConfigParser().read('DATA/CompensatoryHolidays/CompensatoryHolidays.ini',encoding='utf-8')
except Exception as error:
    Message_Box(root, '读取调休表文件失败.\n此错误不影响程序正常运行,但是会导致调休日及其附属功能无法使用.\n错误代码: ' + str(error), '错误', icon='error')







def check_istime_to_playsound():
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


        for item in sort_dict_by_time(Class_Tree):#往右翻,还有一个附加条件->    ->    ->
            if int(Class_Tree[item]['Time'][0][0])==int(time.strftime(r'%H')) and int(Class_Tree[item]['Time'][0][1])==int(time.strftime(r'%M')) and class_on_notice==True:
                win32api.Beep(659,1025);win32api.Beep(523,875);win32api.Beep(587,925);win32api.Beep(392,1250);win32api.Sleep(725);win32api.Beep(392,950);win32api.Beep(587,875);win32api.Beep(659,1025);win32api.Beep(523,800)
                Balloon_Box(root,title='即将上课: %s'%(Class_Tree[item][weekday]),text='别搁那神经了,赶紧滚回位子上准备上课!',staytime=0)


        for i in range(61):#61
            print(i)
            if running_threading==False:
                break
            time.sleep(1)
    
Thread_check_istime_to_playsound=threading.Thread(target=check_istime_to_playsound,daemon=True)
Thread_check_istime_to_playsound.start()









def question_exit():
    global running_threading
    Frame_question=tkinter.Frame(root)
    Frame_question.place(x=0,y=480,width=150,height=40)

    def COMMAND_exit_yes():
        global running_threading
        running_threading=False
        root.destroy()
        sys.exit()
    Button_question_yes=Button(Frame_question,text='是',command=COMMAND_exit_yes)
    Button_question_yes.place(x=0,y=0,width=75,height=40)

    Button_question_no=Button(Frame_question,text='否',command=lambda: Frame_question.destroy())
    Button_question_no.place(x=75,y=0,width=75,height=40)


    '''
    if Message_Box(parent=root,text='是否退出?',title='询问',icon='question',buttonmode=2)!=True: return
    running_threading=False
    root.destroy()
    sys.exit()'''

Button_exit=Button(root,text='退出',command=question_exit)
Button_exit.place(x=0,y=480,width=150,height=40)






Button_root_show=Button(root,text='显示',command=root_show)
Button_root_show.place(x=0,y=root.winfo_height()-20,width=300,height=40)





root_hide()#一些BUG,只能这么做了
root_show()


root.iconbitmap(str(res_icon_folder)+'/icon.ico')
root.wait_window(root)

running_threading=False

sys.exit()



