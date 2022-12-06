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
from decimal import Decimal

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
    global date_to, date_from,MHSInpObj,btn_setparam
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
        if MHSInpObj == None:
            MHSInpObj  = ManHourSystemWrapper.MHSInputUtils(date_from,GID,SEIBAN)
    except Exception as e:
        print(e)
#    btn_start["state"]="normal"
#    btn_reset["state"]="normal"
    [seiban_list,seibanName_list,seibanSubItem_list]=MHSInpObj.get_colums()
    MHSInpObj.get_date(len(seiban_list))

    btn_setparam["state"] = "normal"

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
#    num_colums=MHSInpObj.get_colums()
#    seiban_list=MHSInpObj.get_colums()
    print("対象日=",date_from)
    [seiban_list,seibanName_list,seibanSubItem_list]=MHSInpObj.get_colums()
    print("seiban_list=",seiban_list)
    [key_list,key, kousuu_list,lysithea_hours]=MHSInpObj.set_date(date_year, date_month,date_from,len(seiban_list))

    print("key_list=",key_list)

    num_items=len(kousuu_list)

    sub_win = tk.Toplevel()
    sub_win.geometry("1000x100")


    entry_item=[]
    kousuu_total =0
    for loop in range(0,len(kousuu_list)):
        item = tk.Label(sub_win,text=seiban_list[loop])
        item.grid(column=loop+1+1,row=0)

        item = tk.Label(sub_win,text=seibanName_list[loop])
        item.grid(column=loop+1+1,row=1,padx=10)

        temp=tk.Entry(sub_win,width=15)
        entry_item.append(temp)
        entry_item[loop].grid(column=loop+1+1,row=2)
        entry_item[loop].insert(0,kousuu_list[loop])
#        kousuu_total += int(kousuu_list[loop])
        kousuu_total += Decimal(kousuu_list[loop])
        print("kousuu_total=",kousuu_total)
        print("--")
    """
    item1=tk.Label(sub_win,text=seiban_list[0])
    item1.grid(column=0,row=0)
    entry_item1 = tk.Entry(sub_win,width=5)
    entry_item1.grid(column=1,row=0)
    entry_item1.insert(0,kousuu_list[0])

    item2=tk.Label(sub_win,text=seiban_list[1])
    item2.grid(column=2,row=0)
    entry_item2 = tk.Entry(sub_win,width=5)
    entry_item2.grid(column=3,row=0)
    entry_item2.insert(0,kousuu_list[1])


    item3=tk.Label(sub_win,text=seiban_list[2])
    item3.grid(column=4,row=0)
    entry_item3 = tk.Entry(sub_win,width=5)
    entry_item3.grid(column=5,row=0)
    entry_item3.insert(0,kousuu_list[2])

    item4=tk.Label(sub_win,text=seiban_list[3])
    item4.grid(column=6,row=0)
    entry_item4 = tk.Entry(sub_win,width=5)
    entry_item4.grid(column=7,row=0)
    entry_item4.insert(0,kousuu_list[3])
    """

    print("kousuu_list[0]=>", kousuu_list[0])
    print("entry_item0=",entry_item[0])

    item = tk.Label(sub_win,text=date_from,bg="LightSteelBlue")
    item.grid(column=0,row=0,rowspan = 3, sticky=tk.N+tk.S)

    item = tk.Label(sub_win,text="リシテア")
    item.grid(column=1,row=0,sticky=tk.N+tk.S)

    item = tk.Entry(sub_win,text=str(lysithea_hours))
    item.delete(0,tk.END)
    item.insert(0,str(lysithea_hours))
    item.grid(column=1,row=1,sticky=tk.N+tk.S)

    item = tk.Entry(sub_win,text=str(kousuu_total))
    item.delete(0,tk.END)
    item.insert(0,str(kousuu_total))
    item.grid(column=1,row=2,sticky=tk.N+tk.S)

    btn_register = tk.Button(sub_win, text='Save', command=lambda:ctrl_sub_win(sub_win,key_list,entry_item))
    btn_register.grid(column=len(kousuu_list)+1+1,row=2)

    btn_register = tk.Button(sub_win, text='Close', command=sub_win.destroy)
    btn_register.grid(column=len(kousuu_list)+1+1,row=3)

#    btn_register = tk.Button(sub_win, text='Reload', command=lambda:ctrl_sub_win_reload(sub_win))
    btn_register = tk.Button(sub_win, text='合計確認', command=lambda:ctrl_sub_win_recalculate(sub_win,entry_item))
    btn_register.grid(column=len(kousuu_list)+1+1,row=1)

    sub_win.focus_set()
    print("取得keyは・・",key)

def ctrl_sub_win(win_obj,key_list,entry_items):
#    key="/html/body/form/div[4]/table/tbody/tr[46]/td[5]/input[1]"
    print("entry_item1=",entry_items[0].get())
#    MHSInpObj.send_updated_kousuu(key,0)
#    MHSInpObj.send_updated_kousuu(key_list[0],entry_items[0].get())
#    MHSInpObj.send_updated_kousuu(key_list[1],entry_items[1].get())
    
    for loop in range(0,len(entry_items)):
        MHSInpObj.send_updated_kousuu(key_list[loop],entry_items[loop].get())

    MHSInpObj.click_register()
    win_obj.destroy()

def ctrl_sub_win_reload(win_obj):
    win_obj.destroy()
    set_param()
    return

def ctrl_sub_win_recalculate(win_obj,kousuu):
    kousuu_total=0
    for loop in range(0,len(kousuu)):
        kousuu_total += Decimal(kousuu[loop].get())
    messagebox.showinfo('工数トータル', str(kousuu_total))
    return

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

btn_ok = tk.Button(frame1, text='事前準備', command=click_getkousuu)
btn_ok.grid(column=1,row=2)

btn_start = tk.Button(frame1, text='Start', command=gather_kousuu)
btn_start.grid(column=4,row=2)

btn_setparam = tk.Button(frame1, text='日付設定', command=set_param)
btn_setparam.grid(column=3,row=0)
btn_setparam["state"] = "disable"

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
