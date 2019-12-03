''' -----------------------
program: Sh-StartUp
version: 2.2.20191018

This Python program acts as an alternative (to autoexec*) method to start number of programs in Windows.
It selects programs batch from .JSON file automatically (based on WiFi network SSID available) or 
manually and starts them one by one.
May be useful when you work in different locations (home/office) or different environments.
Although the main mode is autostart at PC login, it may also be useful to start the new environment.
JSON configuration file may contain unlimited number of configurations and applications.

This program uses sh_messagebox module.
------------------------'''

import json
from tkinter import *  
from tkinter.ttk import Combobox
import os.path
import sys
import subprocess
import sh_messagebox as shmb
import time

ECODE_SUCCESS = 0
ECODE_UNKNOWN = 1
ECODE_CONFIGURATION_NOT_FOUND = 2
ECODE_OPERATION_CANCELLED = 3
ECODE_PROGRAM_RECORD_NOT_FOUND = 4
ECODE_ERROR_RUNNING_PROGRAM = 5

MESSAGE_DELAY = 15000

ecode_str = ["Success", "Unknown error", "Configuration not found: \n", 
    "Operation cancelled", "Program record not found: \n", "Error running program: "]
shAppIcon = os.path.dirname(__file__) + "\\Sh-StartUp.ico"
shAppTitle = "Sh StartUp"

class shStartUpGUI:
    def __init__(self, guiListData, guiConfigName):

        self.listData = guiListData 
        self.configName = guiConfigName
        self.prgList = self.listData["Programs"]
        self.configList = list(self.listData["Configurations"].keys())  

        self.exitCode = ECODE_SUCCESS
        self.msgboxFlag = 1

        if self.configName == "":
            self.configName = self.select_network_config()

        self.window = Tk()
        self.window.title(shAppTitle)
        self.window.minsize(width=400, height=50)
        self.window.iconbitmap(shAppIcon) 
        #- put window into screen center
        self.x = (self.window.winfo_screenwidth() - self.window.winfo_reqwidth()) / 2
        self.y = (self.window.winfo_screenheight() - self.window.winfo_reqheight()) / 2
        self.window.wm_geometry("+%d+%d" % (self.x, self.y))
        self.window.protocol('WM_DELETE_WINDOW', self.cmd_cancel)

        self.checkAll = IntVar()

        #--- Panel Buttons
        self.frameButtons = Frame(self.window)
        self.frameButtons.pack(side=BOTTOM, expand=YES, fill=BOTH)
        self.lblFake = Label(self.frameButtons, width=4)
        self.lblFake.pack(side=RIGHT)
        self.buttonCancel = Button(self.frameButtons, text="Cancel", width=10, command=self.cmd_cancel)
        self.buttonCancel.pack(side=RIGHT, padx=5, pady=5)
        self.buttonRun = Button(self.frameButtons, text="Run", width=10, command=self.cmd_run)
        self.buttonRun.pack(side=RIGHT, padx=5, pady=5)

        #---Panel PrgList
        self.framePrgList = LabelFrame(self.window, bd=4, relief=RIDGE, text="Select programs to run:")
        self.framePrgList.pack(side=TOP, expand=YES, fill=BOTH) 

        lbl = Label(self.framePrgList, bd=4, relief=RAISED)
        chk = Checkbutton(lbl, text = "All", variable=self.checkAll, command=self.cmd_process_checkAll)
        chk.pack(side=LEFT, anchor="w", padx = 20)
        lbl.pack(side=TOP, anchor="w", fill=X)

        lbl2 = Label(lbl, bd=4, relief=SUNKEN)
        lbl2.pack(anchor="w", fill=BOTH)
        lblinfo = Label(lbl2, text="Configuration: ")
        lblinfo.pack(side=LEFT, anchor="w", padx = 5, fill=BOTH)

        self.combo1 = Combobox(lbl2)
        self.combo1["values"] = self.configList
        self.combo1.set(self.configName)
        self.combo1.pack(side=LEFT, anchor="w")
        self.combo1.bind('<<ComboboxSelected>>', self.cmd_combo1_selection_changed)

        self.generate_programs_panel()
    
    def generate_programs_panel(self):
      
        if self.configName not in self.configList:
            self.exitCode = ECODE_CONFIGURATION_NOT_FOUND
            ecode_str[ECODE_CONFIGURATION_NOT_FOUND] += self.configName
            self.window.withdraw()
            return
        self.configRecord = self.listData["Configurations"][self.configName]

        self.lblPrgList = Label(self.framePrgList)
        self.lblPrgList.pack(side=TOP, anchor="w", fill=BOTH)

        self.checks = []        
        self.checkAll.set(0) 

        for prg in self.configRecord:
            var = IntVar()
            var.set(prg["Checked"])

            prgRecord = self.get_program_record(prg["Program"])
            if not prgRecord:
                self.exitCode = ECODE_PROGRAM_RECORD_NOT_FOUND
                ecode_str[ECODE_PROGRAM_RECORD_NOT_FOUND] += prg["Program"]
                self.window.withdraw()
                return
            chk2 = Checkbutton(self.lblPrgList, text = prgRecord["DisplayName"], variable=var, command=self.cmd_clear_checkAll)
            self.checks.append(var)
            chk2.pack(side=TOP, anchor="w", padx=25)
    
    def select_network_config(self):
        config = "Default"
        strNetworks = subprocess.check_output(["netsh", "wlan", "show", "networks"])
        if type(strNetworks) == bytes:
            strNetworks = strNetworks.decode('ascii')
        netProfiles = self.listData["NetProfiles"]
        for rec in netProfiles:
            if rec["Network"] in strNetworks:
                config = rec["Configuration"]
                break
        return config

    def get_program_record(self, prgID):
        record = None
        for rec in self.prgList:
            if rec["Program"] == prgID:
                record = rec
                break
        return record

    def cmd_cancel(self):
        self.exitCode = ECODE_OPERATION_CANCELLED
        self.window.destroy()

    def cmd_run(self):
        self.exitCode = ECODE_SUCCESS
        var = 0
        for rec in self.configRecord:
            prgRecord = self.get_program_record(rec["Program"])
            if self.checks[var].get():
                try:
                    subprocess.Popen(prgRecord["Command"], shell=True)
                except:
                    self.exitCode = ECODE_ERROR_RUNNING_PROGRAM
                    ecode_str[ECODE_ERROR_RUNNING_PROGRAM] += "\n" + rec["Command"]
            var +=1
            time.sleep(0.2)
        self.window.destroy()

            
    def cmd_clear_checkAll(self):
        self.checkAll.set(0)

    def cmd_process_checkAll(self):
        for chk_val in self.checks:
            chk_val.set(1 if self.checkAll.get() else 0)

    def cmd_combo1_selection_changed(self, param):
        self.lblPrgList.destroy()
        self.configName = self.combo1.get()
        self.generate_programs_panel()
    
    def main_cycle(self):
        if self.exitCode > ECODE_SUCCESS: 
            return
        self.exitCode = ECODE_UNKNOWN
        self.window.mainloop()
        self.msgboxFlag = 0
    

if __name__ == "__main__":

    try:
        configName = sys.argv[1]
    except:
        configName = ""

    fileName = "Sh-StartUp-List.json"
    with open(fileName, "r") as list_file:
        listData = json.load(list_file)

    objGUI = shStartUpGUI(listData, configName)
    objGUI.main_cycle()

    if objGUI.exitCode:
        shmb.shShowMessage(objGUI.msgboxFlag, "Sh-StartUp ERROR: " + str(objGUI.exitCode), 
            "ERROR: " + str(objGUI.exitCode) + "\n" + ecode_str[objGUI.exitCode], 
            shmb.SH_MESSAGE_ERROR, MESSAGE_DELAY, shAppIcon)
    sys.exit(objGUI.exitCode)
