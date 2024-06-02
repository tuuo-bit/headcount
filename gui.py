import time
import subprocess
import pandas
from datetime import date
import os
from tkinter import *
import tkinter as tk
from tkinter import ttk
class headcount:
      def __init__(s):
            s.color=["#CCC5B9", "#DBD9D1","#EDECE8", "#003554", "#ED2D07", "#D92C0A", "#000000", "#FFFFFF"]
            s.font=[ "Bahnschrift 18 bold", "Bahnschrift 16", "Haettenschweiler 106"]
            s.window = Tk()
            s.window.configure( bg = s.color[0])
            s.window.state( "zoomed")
            s.window.title( "headcount")
            s.imgArrowR = PhotoImage(file = "resource\\arrow.png")
            s.imgArrowL = PhotoImage(file = "resource\\worra.png")
            s.imgReload = PhotoImage(file = "resource\\reload.png")
            s.currentDate = date.today().strftime("%d.%m.%Y")
            #
            s.uiOpenExistingFileList =[]
            s.updateFileList()
            s.uiOpenExistingFileName = StringVar()
            s.uiCreateNewFileName = StringVar()
            s.uiRegisteredSSIDList = StringVar()
            s.uiNewSSIDList = StringVar()
            s.uiAbsentList = StringVar()
            s.uiPresentList = StringVar()
            s.uiSummaryDateList = []
            s.uiSummaryDate = StringVar()
            s.uiFileName = StringVar()
            #
            s.frame2 = tk.Frame(s.window,height = 600,width = 1230,background = s.color[1])
            s.frame1 = tk.Frame(s.window,height = 600,width = 1230,background = s.color[1])
            #
            s.frame2.place(relx = 0.5,rely = 0.5,anchor="center")
            s.frame1.place(relx = 0.5,rely = 0.5,anchor="center")
            return
      def validateFileName(s,str):
            if len(str)==0 or str.isspace() or str=="None":
                  return False
            for ch in "/\\?:<*>|\"":
                  if ch in str:
                        return False
            return True
      def updateFileList(s):
            s.uiOpenExistingFileList = [f for f in os.listdir() if ".xlsx" in f]+["None"]
            return
      def restart(s):
            s.updateFileList()
            s.combobox["values"] = s.uiOpenExistingFileList
            s.frame1.tkraise()
            return
      def UI(s):
            # Heading
            ttk.Label( s.frame1, text = "headcount",foreground=s.color[3] ,background = s.color[1], font=s.font[2]).place( relx = 0.43, rely = 0.63, anchor="center")            # Create new
            ttk.Label( s.frame1, text="Create New\n\nOpen Existing", foreground = "#FFFFFF", background = s.color[1], font = s.font[0]).place( relx = 0.25, rely = 0.3, anchor="nw")
            ttk.Label( s.frame1, text=".xlsx", foreground = "#FFFFFF", background = s.color[1], font = s.font[0]).place( relx = 0.65, rely = 0.3, anchor="nw")
            ttk.Entry( s.frame1, font = s.font[1], textvariable = s.uiCreateNewFileName).place( relx = 0.4, rely = 0.3, anchor="nw", width = 300)
            # Open existing
            s.combobox = ttk.Combobox( s.frame1, textvariable = s.uiOpenExistingFileName, font = s.font[1], state="readonly")
            s.combobox["values"] = s.uiOpenExistingFileList
            s.combobox.place(relx = 0.4,rely = 0.4,anchor="nw",width = 300)
            # Open/Next Button
            win1_button = tk.Button( s.frame1, image = s.imgArrowR, command = s.open)
            win1_button.config(background = s.color[4], activebackground = s.color[5], borderwidth = 0)
            win1_button.place(relx = 0.66,rely = 0.65,width = 83, height=83,  anchor="center")
            # File Name Display
            tk.Label( s.frame2, textvariable = s.uiFileName, foreground = s.color[3], background = s.color[2], font = s.font[0], justify="center").place( relx = 0.5, rely = 0.04, height = 51, width = 1230,anchor="n")
            # Close/Back Button
            win2_button = tk.Button( s.frame2,image = s.imgArrowL,command = s.restart)
            win2_button.config(background = s.color[4], activebackground = s.color[5], borderwidth = 0)
            win2_button.place( relx = 0.02, rely = 0.04, anchor="nw")
            # Reload button
            rel_button = tk.Button( s.frame2,image = s.imgReload,command = s.scan)
            rel_button.config(background = s.color[4], activebackground = s.color[5], borderwidth = 0)
            rel_button.place( relx = 0.98, rely = 0.04, anchor="ne")
            # Summary Frame
            s.frame20 = tk.Frame( s.frame2,background = s.color[2],height = 475,width = 577)
            s.frame20.place( relx = 0.02, rely = 0.17, anchor="nw")
            s.datesComboBox = ttk.Combobox( s.frame20, textvariable = s.uiSummaryDate, values = s.uiSummaryDateList, font = s.font[1], state="readonly")
            s.datesComboBox.place( relx = 0.3, rely = 0.05, height = 40, width = 160, anchor="nw")
            sum_button = tk.Button( s.frame20, text="Summarize", font = s.font[1], command = s.summarize)
            sum_button.config( background = s.color[4], activebackground = s.color[5], borderwidth = 0)
            sum_button.place( relx = 0.04, rely = 0.05, anchor="nw")
            tk.Listbox( s.frame20, listvariable = s.uiAbsentList, font = s.font[1],
                        selectmode="multiple", background = s.color[1], foreground=s.color[7],borderwidth = 0, relief="flat",
                        selectborderwidth = 0, selectbackground = s.color[1], selectforeground=s.color[3]
                        ).place( relx = 0.04, rely = 0.18,height = 168,width = 527,anchor="nw")
            s.absentLabel = ttk.Label( s.frame20, text="Absent", background = s.color[1], font = s.font[1], foreground=s.color[3])
            s.absentLabel.place( relx = 0.94, rely = 0.53, anchor="se")
            tk.Listbox( s.frame20,listvariable = s.uiPresentList, foreground=s.color[7],
                        font = s.font[1],selectmode="multiple",background = s.color[1], borderwidth = 0, relief="flat",
                        selectborderwidth = 0, selectbackground = s.color[1], selectforeground=s.color[3]
                        ).place(relx = 0.04,rely = 0.95,height = 168,width = 527,anchor="sw")
            s.presentLabel = ttk.Label( s.frame20, text="Present", background = s.color[1], font = s.font[1], foreground=s.color[3])
            s.presentLabel.place( relx = 0.94, rely = 0.945, anchor="se")
            # Show Networks Notebook
            s.nb = ttk.Notebook( s.frame2)
            # Registered SSID
            s.frame21 = tk.Frame( s.nb,background = s.color[2], height = 475, width = 577)
            s.listbox21 = tk.Listbox( s.frame21, listvariable = s.uiRegisteredSSIDList, font = s.font[1], foreground=s.color[7],
                        selectmode="multiple", background = s.color[1], borderwidth = 0, relief="flat",
                        selectborderwidth = 0, selectbackground = s.color[0], selectforeground=s.color[4])
            s.listbox21.place( relx = 0.04, rely = 0.05, height = 335, width = 527, anchor="nw")
            rem_button = tk.Button( s.frame21, text="Remove", font = s.font[1], command = s.remove)
            rem_button.config( background = s.color[4], activebackground = s.color[5], borderwidth = 0)
            rem_button.place( relx = 0.96, rely = 0.95, anchor="se")
            take_button = tk.Button( s.frame21, text="Take Attendance",font = s.font[1], command = s.takeAttendance)
            take_button.config( background = s.color[4], activebackground = s.color[5], borderwidth = 0)
            take_button.place( relx = 0.04, rely = 0.95, anchor="sw")
            # New SSID
            s.frame22 = tk.Frame( s.nb,background = s.color[2], height = 475, width = 577)
            s.listbox22 = tk.Listbox( s.frame22, listvariable = s.uiNewSSIDList, foreground=s.color[7],
                        font = s.font[1], selectmode="multiple", background = s.color[1], borderwidth = 0, relief="flat",
                        selectborderwidth = 0, selectbackground = s.color[0], selectforeground=s.color[4])
            s.listbox22.place( relx = 0.04, rely = 0.05, height = 335, width = 527, anchor="nw")
            add_button = tk.Button( s.frame22, text="Add", font = s.font[1], command = s.add)
            add_button.config( background = s.color[4], activebackground = s.color[5], borderwidth = 0)
            add_button.place(relx = 0.96, rely = 0.95, anchor="se")
            # Notebook
            s.nb.add( s.frame21,text="Registered Networks")
            s.nb.add( s.frame22,text="New Networks")
            s.nb.place( relx = 0.51, rely = 0.17, height = 475, width = 577, anchor="nw")
            return
      def open(s):
            s.updateFileList()
            s.frame2.tkraise()
            v1 = s.validateFileName(s.uiOpenExistingFileName.get())
            v2 = s.validateFileName(s.uiCreateNewFileName.get())
            if v1==True and v2==True:
                  s.classFileName = s.uiOpenExistingFileName.get()
            elif v1==True and v2==False:
                  s.classFileName = s.uiOpenExistingFileName.get()
            elif v1==False and v2==True:
                  s.classFileName = s.uiCreateNewFileName.get()+".xlsx"
                  s.create(s.classFileName)
            else:
                  s.classFileName = "Class Of "+s.currentDate+".xlsx"
                  s.create(s.classFileName)
            s.uiFileName.set(s.classFileName)
            s.classDF = pandas.read_excel(s.classFileName)
            print(f"~ Opened: {s.classFileName[:-5]} ~")
            s.scan()
            #s.frame2.tkraise()
            return
      def create(s,filename):
            pandas.DataFrame({"ssid":[],"id":[]}).to_excel(filename,index = False)
            return
      def close(s):
            return
      def scan(s):
            subprocess.run("netsh interface set interface name=\"Wi-Fi\" admin = disabled")
            time.sleep(1)
            subprocess.run("netsh interface set interface name=\"Wi-Fi\" admin = enabled")
            time.sleep(1)
            s.rawtext = subprocess.check_output("netsh wlan show networks").decode(encoding="UTF-8",errors="ignore").replace('\r','').split('\n')
            s.update()
            return
      def update(s):
            s.visibleAll = int(s.rawtext[2].split(" ")[2])
            s.visibleAllSSID=[]
            for i in range(s.visibleAll):
                  t = " ".join(s.rawtext[5*i+4].split(" ")[3:])
                  s.visibleAllSSID.append(t)
            s.visibleNewSSID = list(filter(lambda x: x not in s.classDF["ssid"].to_list(),s.visibleAllSSID))
            s.visibleNew = len(s.visibleNewSSID)
            s.visibleRegSSID = list(filter(lambda x: x not in s.visibleNewSSID, s.visibleAllSSID))
            s.visibleReg = len(s.visibleRegSSID)
            s.uiRegisteredSSIDList.set(s.visibleRegSSID)
            s.uiNewSSIDList.set(s.visibleNewSSID)
            s.datesComboBox["values"] = s.classDF.columns.to_list()[2:][::-1]
            return
      def add(s):
            s.ssidToAdd = list(map(s.visibleNewSSID.__getitem__,s.listbox22.curselection()))
            s.uiDisplaySSID = StringVar()
            if s.ssidToAdd:
                  s.frame3 = tk.Frame( s.frame2, height = 475, width = 577, background = s.color[0])
                  s.frame3.place(relx = 0.51,rely = 0.17)
                  s.frame3.tkraise()
                  s.listbox3 = tk.Listbox( s.frame3, listvariable = s.uiDisplaySSID, foreground=s.color[3],
                        font=s.font[1], selectmode="multiple",background = s.color[1], borderwidth = 0, relief="flat",
                        selectborderwidth = 0, selectbackground = s.color[2], selectforeground=s.color[6])
                  s.uiDisplaySSID.set(s.ssidToAdd)
                  s.listbox3.place(relx = 0.04, rely = 0.05, height = 335, width = 250, anchor="nw")
                  rollEntry = tk.Text( s.frame3, borderwidth = 0, relief = "flat", background = s.color[1], foreground=s.color[4], font = s.font[1])
                  rollEntry.place( rely = 0.05, relx = 0.96, height = 335, width = 251, anchor="ne")
                  s.rollForSSID = rollEntry
                  reg_button = tk.Button( s.frame3, text="Register", font = s.font[1], command = s.addRoll)
                  reg_button.config( background = s.color[4], activebackground = s.color[5], borderwidth = 0)
                  reg_button.place( relx = 0.96, rely = 0.95, anchor="se")
            return
      def addRoll(s):
            s.roll = list(map(int,s.rollForSSID.get("1.0","end").split()))[:len(s.ssidToAdd)]
            s.nb.tkraise()
            s.classDF = s.classDF._append(pandas.DataFrame({"ssid":s.ssidToAdd,"id":s.roll}),ignore_index = True).sort_values("id")
            s.classDF.to_excel(s.classFileName,index = False)
            s.scan()
            return
      def remove(s):
            ssidToRemove = list(map(s.visibleRegSSID.__getitem__,s.listbox21.curselection()))
            if ssidToRemove:
                  s.classDF = s.classDF.drop(s.classDF[s.classDF["ssid"].isin(ssidToRemove)].index,axis = 0).sort_values("id").reset_index(drop = True)
                  s.classDF.to_excel(s.classFileName,index = False)
                  s.scan()
            return
      def takeAttendance(s):
            reg = s.classDF["ssid"].to_list()
            avai = s.visibleRegSSID
            dic={1:"P",0:"A"}
            l = []
            for ssid in reg:
                  l.append(dic[int(ssid in avai)])
            s.classDF[s.currentDate]=l
            s.classDF.to_excel(s.classFileName,index = False)
            s.datesComboBox["values"] = s.classDF.columns.to_list()[2:][::-1]
            return
      def summarize(s):
            s.summaryDate = s.uiSummaryDate.get()
            if not s.summaryDate:
                  return
            sdate = s.summaryDate
            adL = s.classDF[sdate]
            iL= s.classDF["id"]
            s.presentL,s.absentL=[],[]
            for i in range(len(iL)):
                  if adL[i]=="A":
                        s.absentL.append(iL[i])
                        continue
                  if adL[i]=="P":
                        s.presentL.append(iL[i])
                        continue
            s.present,s.absent = len(s.presentL),len(s.absentL)
            s.uiPresentList.set(s.presentL)
            s.uiAbsentList.set(s.absentL)
            s.absentLabel["text"] = str(len(s.absentL))+" Absent"
            s.presentLabel["text"] = str(len(s.presentL))+" Present"
            return
h = headcount()
h.UI()
h.window.mainloop()