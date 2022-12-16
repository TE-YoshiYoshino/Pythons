# -*- coding: Shift-JIS -*-#
#
#  File Name: ManHoursGetSet3Sys.py
# ------------------------------------------------------------------------------------
#
# History ########################################
#   Dec. 8th, 2022: Creation from ManHourGetSet3Sys.py
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
from concurrent.futures import ThreadPoolExecutor

#-- GLOBALS ------------------------------------------------------------------------


#-- LOCAL FUNCTIONS ------------------------------------------------------------
def main():
    global main_win,text_area
    tpe = ThreadPoolExecutor(max_workers=10)
    futures=[]
    main_win.title("工数確認&入力") 
    main_win.geometry("600x400")
    frame1=tk.Frame(main_win,height=100,width=600)
    frame2=tk.Frame(main_win,height=300,width=600)
    frame3=tk.Frame(main_win,height=300,width=600)
    frame1.pack()
    frame2.pack()
    frame3.pack()
    main_win.data_entry_date_from = DateEntry(frame1,style='my.DateEntry')
    main_win.data_entry_date_from.grid(column=1,row=0)
    btn_showKousuu = tk.Button(frame2, text='Show', command=lambda:showKousuu(tpe))
    btn_showKousuu.pack()
    btn_exit = tk.Button(frame2, text='Exit', command=click_exit)
    btn_exit.pack()

    text_area = scrolledtext.ScrolledText(frame2,  
        wrap = tk.WORD, #単語単位で改行
        width = 80,
        height = 20,  
    ) 

    #text_area.grid(column = 0, row=3, padx = 10, pady = 10) 
    text_area.pack()
    text_area.focus() 

    main_win.mainloop() 

def showKousuu(executor):
    print("こんちわ")
    executor.submit(getReadyAndShowList)
    messagebox.showinfo('Infomation', "Please wait for a while until system initiation finished.")

def click_exit():
    global MHSInpObj
    if MHSInpObj:
        del MHSInpObj
    quit()

def getReadyAndShowList():
    global date_to, date_from,MHSInpObj,btn_setparam,main_win,text_area
    print("こんちわ2")
    date_from=main_win.data_entry_date_from.get_date()
    # -- Replace "-"
    locale.setlocale(locale.LC_TIME, '')
    print("date_from=",str(date_from)[0:4],str(date_from)[0:4], str(date_from)[5:7],str(date_from)[8:10])
    date = dt.date(int(str(date_from)[0:4]), int(str(date_from)[5:7]), int(str(date_from)[8:10]))
    date = dt.date(2022, 8, 1)
    print(date.strftime('%A'))  # => '火曜日'
    print(date.strftime('%a'))  # => '火'
    date_from=str(date_from).replace('-','/')
    date_from=str(date_from)[5:7]+'/'+str(date_from)[8:10]+'('+date.strftime('%a')+')'
    print("date_from = ", date_from)
    try:
        if MHSInpObj == None:
            MHSInpObj  = ManHourSystemWrapper.MHSInputUtils(date_from,0,0)
    except Exception as e:
        print(e)
    [seiban_list,seibanName_list,seibanSubItem_list]=MHSInpObj.get_colums()
#    MHSInpObj.get_date(len(seiban_list))
    [allKousuu_List,Date_List] = MHSInpObj.get_allKouusuu(len(seiban_list))

    print("Nof seibanName_list=", seiban_list )
#    text_area.insert(tk.END,  "(Date)      " + "\t" + "(製番)  "+"\t" +"(工数)     " + "\t\n" )
    text_area.insert(tk.END,  "(Date)      " )
    i=0
    while i < len(seiban_list):
        text_area.insert(tk.END, "\t"+seibanName_list[i])
        i+=1
    
    text_area.insert(tk.END, "\n")
    i=1
    n=0
    while i < len(allKousuu_List) + 1:
        j = 1
        text_area.insert(tk.END, Date_List[n]+"\t" )
        n+=1
        while j < len(seiban_list) +1 :
#        print("Date=",Date[i-1])
#            gKousuuData.append([Date[i-1],Seiban[i-1],ManHour[i-1]])
#            text_area.insert(tk.END, Date[i-1] + "\t" + seibanName[i-1] + "\t" +ManHour[i-1] +"\n")
            text_area.insert(tk.END, "                 "+allKousuu_List[i-1] )
            text_area.insert(tk.END, "\t" )
            j += 1
            i+=1
        text_area.insert(tk.END, "\n" )
    btn_setparam["state"] = "normal"
    print("こんちわ3")

#-- MAIN -----------------------------------------------------------------------------
if __name__ == "__main__":
    MHSInpObj = None
    text_area = None
    main_win = tk.Tk() 
    main()
