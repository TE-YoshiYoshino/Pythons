# -*- coding: Shift-JIS -*-#
#
#  File Name: ManHourInput.py
# ------------------------------------------------------------------------------------
#
# History ########################################
#   Aug.17 ,2022: Creation from ManHourCheck.py
##############################################
#**** MODULES ****
import tkinter
from tkinter import *
from tkinter import ttk
from tkcalendar import Calendar, DateEntry
import tkinter as tk 
from tkinter import scrolledtext 
from tkinter import messagebox
import random
import time
from tkinter import filedialog
import ManHourSystemWrapper
import datetime as dt
import locale

#**** GLOBALS ****
MHSobj = None
MHSInpObj = None
gKousuuData = []
gDir='C:\\Users'
#**** FUNCTIONS ****
def input_text():
    text = random.choice(texts)
    text_area.insert(tk.END, text + "\n")

def check_input_value(dt1,dt2,GID,SEIBAN):
    if dt2 < dt1:
        messagebox.showerror ('Error', '日付だめ～')
        return False
    GID.strip()
#    if GID.isspace()==False:
    if len(GID) > 0:
        if str.isdecimal(GID) == False:
            messagebox.showerror('Error', 'GIDが数字じゃないかスペース入ってます')
            return False
        if int(GID) > 9999999999:
            messagebox.showerror('Error', 'GIDが桁あふれ')
            return False
    return True

def click_ok():
    global date_to, date_from,MHSobj
    date_from=main_win.data_entry_date_from.get_date()
    date_to=main_win.data_entry_date.get_date()
    GID=entry_GID .get()
    SEIBAN=entry_SEIBAN .get()
    if check_input_value(date_from,date_to,GID,SEIBAN) == False:
        return

    # -- Replace "-"
    date_from=str(date_from).replace('-','/')
    date_to=str(date_to).replace('-','/')

    try:
        MHSobj  = ManHourSystemWrapper.MHSUtils(date_from,date_to,GID,SEIBAN)
    except Exception as e:
        print(e)
    btn_start["state"]="normal"
    btn_reset["state"]="normal"

def gather_kousuu():
    try:
        Date = MHSobj.get_date()
        Seiban = MHSobj.get_seiban()
        GIDs = MHSobj.get_gids()
        ManHour = MHSobj.get_manhour()
    except Exception as e:
        print(e)
    text_area.insert(tk.END,  "(Date)      " +"\t" + "(GID)     "+ "\t" + "(製番)  "+"\t" +"(工数)     " + "\t\n" )
    i=1
    while i < len(Date) + 1:
        print("Date=",Date[i-1])
        gKousuuData.append([Date[i-1],Seiban[i-1],ManHour[i-1]])
        text_area.insert(tk.END, Date[i-1] + "\t" + GIDs[i-1] + "\t" + Seiban[i-1] + "\t" +ManHour[i-1] +"\n")
        i +=1
    print("Kousu=",gKousuuData)
    return

def click_reset():
    global MHSobj
    del MHSobj
    btn_reset["state"]="disable"
    return

def click_exit():
    global MHSobj,MHSInpObj
    if MHSobj:
        del MHSobj
    if MHSInpObj:
        del MHSInpObj
    quit()

def file_select():
    filetype = [("エクセル","*.xlsx")]
    filePath= tk.filedialog.askopenfilename(initialdir = gDir, filetypes=filetype)
    entry_file.insert(tk.END, filePath)

def click_getkousuu():
    global date_to, date_from,MHSInpObj
    date_from=main_win.data_entry_date_from.get_date()
    GID=entry_GID .get()
    SEIBAN=entry_SEIBAN .get()
#    if check_input_value(date_from,date_to,GID,SEIBAN) == False:
#        return

    # -- Replace "-"
    locale.setlocale(locale.LC_TIME, '')
#    print("date_from=",str(date_from)[0:4],str(date_from)[0:4], str(date_from)[5:7],str(date_from)[8:10])
    date = dt.date(int(str(date_from)[0:4]), int(str(date_from)[5:7]), int(str(date_from)[8:10]))
#    date = dt.date(2022, 8, 1)
    print(date.strftime('%A'))  # => '火曜日'
    print(date.strftime('%a'))  # => '火'
#    date_from=str(date_from).replace('-','/')
    date_from=str(date_from)[5:7]+'/'+str(date_from)[8:10]+'('+date.strftime('%a')+')'
    print("date_from = ", date_from)
    try:
        MHSInpObj  = ManHourSystemWrapper.MHSInputUtils(date_from,GID,SEIBAN)
    except Exception as e:
        print(e)
#    btn_start["state"]="normal"
#    btn_reset["state"]="normal"
    MHSInpObj.get_colums()
    MHSInpObj.get_date()

def set_param():
#    global date_to, date_from,MHSInpObj
    date_from=main_win.data_entry_date_from.get_date()
    GID=entry_GID .get()
    SEIBAN=entry_SEIBAN .get()
    locale.setlocale(locale.LC_TIME, '')
    date = dt.date(int(str(date_from)[0:4]), int(str(date_from)[5:7]), int(str(date_from)[8:10]))
    date_from=str(date_from)[5:7]+'/'+str(date_from)[8:10]+'('+date.strftime('%a')+')'
    date_year=str(date_from)[0:4]
    date_month=str(date_from)[5:7]
    num_colums=MHSInpObj.get_colums()
<<<<<<< HEAD
    [key, kousuu_list]=MHSInpObj.set_date(date_year, date_month,date_from)
=======
    kousuu_list=MHSInpObj.set_date(date_year, date_month,date_from)
>>>>>>> f645258383175c7f535111b147f2d1835a603328
#def sub_window():
    sub_win = tk.Toplevel(background='green')
    sub_win.geometry("300x100")
#    label_sub = tk.Label(sub_win, text="サブウィンドウ")
#    label_sub.pack()

    item1=tk.Label(sub_win,text='1個目')
    item1.grid(column=0,row=0)
    entry_item1 = tk.Entry(sub_win,width=5)
    entry_item1.grid(column=1,row=0)
    entry_item1.insert(0,kousuu_list[0])
    btn_register = tk.Button(sub_win, text='登録', command=lambda:ctrl_sub_win(sub_win))
    btn_register.grid(column=1,row=1)
    sub_win.focus_set()
<<<<<<< HEAD
    print("取得keyは・・",key)
=======
>>>>>>> f645258383175c7f535111b147f2d1835a603328

def ctrl_sub_win(win_obj):
    key="/html/body/form/div[4]/table/tbody/tr[46]/td[5]/input[1]"
    MHSInpObj.send_updated_kousuu(key,0)
    MHSInpObj.click_register()
    win_obj.destroy()


# MAIN ************************************************
#def main():
main_win = tk.Tk() 
main_win.title("工数確認&入力") 
main_win.geometry("600x400")
frame1=tk.Frame(main_win,height=100,width=600)
#frame2=tk.Frame(main_win,height=400,width=600)
frame3=tk.Frame(main_win,height=200,width=600)
frame4=tk.Frame(main_win,height=200,width=600)
frame1.pack()
#frame2.pack()
frame3.pack()
frame4.pack()

# --frame1---------------------------------------------
lbl0=tk.Label(frame1,text='いつにする？')
lbl0.grid(column=0,row=0)
main_win.data_entry_date_from = DateEntry(frame1,style='my.DateEntry')
main_win.data_entry_date_from.grid(column=1,row=0)

lbl2=tk.Label(frame1,text='GID')
lbl2.grid(column=0,row=1)
entry_GID = tk.Entry(frame1,width=10)
entry_GID.grid(column=1,row=1)

lbl3=tk.Label(frame1,text='製番')
lbl3.grid(column=3,row=1)
entry_SEIBAN = tk.Entry(frame1,width=10)
entry_SEIBAN.grid(column=4,row=1)

btn_ok = tk.Button(frame1, text='工数確認', command=click_getkousuu)
btn_ok.grid(column=1,row=2)

btn_start = tk.Button(frame1, text='Start', command=gather_kousuu)
btn_start.grid(column=4,row=2)

btn_setparam = tk.Button(frame1, text='日付設定', command=set_param)
btn_setparam.grid(column=3,row=0)

#--frame2----------------------------------------------
#text_area = scrolledtext.ScrolledText(frame2,  
#    wrap = tk.WORD, #単語単位で改行
#    width = 80,
#    height = 20,  
#) 
#text_area.pack()
#text_area.focus() 

#--frame3----------------------------------------------
btn_exit = tk.Button(frame3, text='Exit', command=click_exit)
#btn_exit.place(x=400, y=400)
btn_exit.pack()

#--frame4----------------------------------------------
text_area = scrolledtext.ScrolledText(frame4,  
    wrap = tk.WORD, #単語単位で改行
    width = 80,
    height = 20,  
) 
text_area.pack()
text_area.focus() 


# --配列用意----------------------------------------------- 
texts = ['python', 'java', 'ruby']

# --ループ開始---------------------------------------------
main_win.mainloop() 


# END OF MAIN ************************************************
