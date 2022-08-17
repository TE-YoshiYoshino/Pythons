# -*- coding: Shift-JIS -*-#
#
#  File Name: ManHourCheck.py
# ------------------------------------------------------------------------------------
#
# History ########################################
#	Aug.1 ,2022: Creation
#   Aug.2, 2022: Start to manage with GitHub
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

#**** GLOBALS ****
MHSobj = None
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
#    text_area.insert(tk.END, GID + "\n")
#    text_area.insert(tk.END, date_from)
#    text_area.insert(tk.END, date_to)
    SEIBAN=entry_SEIBAN .get()
#    print("date_from", date_from)
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
#        WS.cell(i,1,Date_List[i-1])
#        WS.cell(i,2,Seiban_List[i-1])
#        WS.cell(i,3,ManPower_List[i-1])
        print("Date=",Date[i-1])
        gKousuuData.append([Date[i-1],Seiban[i-1],ManHour[i-1]])
#        gKousuuData[i-1][1] = Seiban[i-1]
#        gKousuData[i-1][2] = GIDs[i-1]
#        gKousuuData[i-1][3] = ManHour[i-1]
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
    global MHSobj
    del MHSobj
    quit()

def file_select():
    filetype = [("エクセル","*.xlsx")]
    filePath= tk.filedialog.askopenfilename(initialdir = gDir, filetypes=filetype)
    entry_file.insert(tk.END, filePath)

# MAIN ************************************************
#def main():
main_win = tk.Tk() 
main_win.title("工数確認") 
main_win.geometry("600x400")
frame1=tk.Frame(main_win,height=100,width=600)
frame2=tk.Frame(main_win,height=400,width=600)
frame3=tk.Frame(main_win,height=200,width=600)
frame1.pack()
frame2.pack()
frame3.pack()

#frame1.grid(column=)
# --frame1---------------------------------------------
"""
btn = tk.Button(main_win, text='input', command=input_text)
btn.place(x=10, y=70)
btn_ok = tk.Button(main_win, text='Ok', command=click_ok)
btn_ok.place(x=100, y=70)
btn_start = tk.Button(main_win, text='Start', command=gather_kousuu)
btn_start.place(x=300, y=70)
btn_start["state"]="disable"
btn_reset = tk.Button(main_win, text='Clear', command=click_reset)
btn_reset.place(x=400, y=70)
btn_reset["state"]="disable"
"""
lbl0=tk.Label(frame1,text='From')
#lbl0.pack(side=tk.LEFT)
lbl0.grid(column=0,row=0)
main_win.data_entry_date_from = DateEntry(frame1,style='my.DateEntry')
#main_win.data_entry_date_from.place(x=30, y=5)
#main_win.data_entry_date_from.pack(side=tk.LEFT)
main_win.data_entry_date_from.grid(column=1,row=0)


lbl1=tk.Label(frame1,text='To')
#lbl1.pack(side=tk.LEFT)
lbl1.grid(column=2,row=0)

lbl12=tk.Label(frame1,text='～')
#lbl1.pack(side=tk.LEFT)
lbl12.grid(column=3,row=0)

main_win.data_entry_date = DateEntry(frame1,style='my.DateEntry')
#main_win.data_entry_date.pack(side=tk.LEFT)
main_win.data_entry_date.grid(column=4,row=0)

lbl2=tk.Label(frame1,text='GID')
#lbl2.place(x=10,y=30)
lbl2.grid(column=0,row=1)
entry_GID = tk.Entry(frame1,width=10)
#entry_GID.place(x=40,y=30)
entry_GID.grid(column=1,row=1)

lbl3=tk.Label(frame1,text='製番')
#lbl3.place(x=120,y=30)
lbl3.grid(column=3,row=1)
entry_SEIBAN = tk.Entry(frame1,width=10)
#entry_SEIBAN.place(x=150,y=30)
entry_SEIBAN.grid(column=4,row=1)

btn_ok = tk.Button(frame1, text='Ok', command=click_ok)
#btn_ok.pack(side=tk.LEFT, padx=10)
btn_ok.grid(column=1,row=2)

btn_start = tk.Button(frame1, text='Start', command=gather_kousuu)
#btn_start.pack(side=tk.LEFT, padx=10)
btn_start.grid(column=4,row=2)


#--frame2----------------------------------------------
# --スクロールテキスト設定---------------------------------
"""
text_area = scrolledtext.ScrolledText(main_win,  
    wrap = tk.WORD, #単語単位で改行
    width = 80,
    height = 20,  
) 
"""
text_area = scrolledtext.ScrolledText(frame2,  
    wrap = tk.WORD, #単語単位で改行
    width = 80,
    height = 20,  
) 

#text_area.grid(column = 0, row=3, padx = 10, pady = 10) 
text_area.pack()
text_area.focus() 

#--frame3----------------------------------------------
lbl4=tk.Label(frame3, text='保存ファイル')
#lbl4.place(x=10,y=350)
lbl4.pack(side=tk.LEFT)

btn_refFile=tk.Button(frame3, text='参照', command=file_select)
#btn_refFile.place(x=400, y=100)
btn_refFile.pack(side=tk.LEFT, padx=2)

entry_file = tk.Entry(frame3,width=60)
#entry_file.place(x=80,y=350)
entry_file.pack(side=tk.LEFT, padx=2)

btn_exit = tk.Button(frame3, text='Exit', command=click_exit)
#btn_exit.place(x=400, y=400)
btn_exit.pack()

# --配列用意----------------------------------------------- 
texts = ['python', 'java', 'ruby']

"""
lbl0=tk.Label(text='From')
lbl0.place(x=0,y=5)
main_win.data_entry_date_from = DateEntry(style='my.DateEntry')
main_win.data_entry_date_from.place(x=30, y=5)

lbl1=tk.Label(text='To')
lbl1.place(x=120,y=5)
main_win.data_entry_date = DateEntry(style='my.DateEntry')
main_win.data_entry_date.place(x=170, y=5)

lbl2=tk.Label(text='GID')
lbl2.place(x=10,y=30)
entry_GID = tk.Entry(width=10)
entry_GID.place(x=40,y=30)

lbl3=tk.Label(text='製番')
lbl3.place(x=120,y=30)
entry_SEIBAN = tk.Entry(width=10)
entry_SEIBAN.place(x=150,y=30)
"""

# --ループ開始---------------------------------------------
main_win.mainloop() 
print("date_from", date_from)
print("date_to", date_to)

    # Get KOUSU Date

    # Show Results

# END OF MAIN ************************************************
