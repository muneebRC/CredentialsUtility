#!/usr/bin/python3
from email.message import Message
import tkinter as tk
from tkinter import EXCEPTION, simpledialog
import tkinter.ttk as ttk
from tkinter.scrolledtext import ScrolledText
import subprocess
from tkinter import messagebox
import random
from docxtpl import DocxTemplate
import string
import os
import threading
from tkinter import filedialog
from selenium import webdriver
import time

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

from configparser import ConfigParser
import sys

def update_block():
    try:
        if not os.path.exists(r'info\versioninfo.ini'):
            os.makedirs('info')
            version_file = ConfigParser()
            version_file['version'] = {'current': '1.0'}
            with open(r'info\versioninfo.ini',"w") as file_object:
                version_file.write(file_object) 
        update_ini = r'PATH'
        parser = ConfigParser()
        parser.read(r'info\versioninfo.ini')
        current = float(parser.get('version', 'current'))
        parser.read(update_ini)
        latest = float(parser.get('version', 'current'))
        if current < latest:
            response = messagebox.askyesno('Update Available', 'Would you like to update to the latest version?\n\nCurrent Version: {0} > New version: {1}'.format(current, latest))
            if response == 1:
                subprocess.Popen(r'updater.exe')
                sys.exit()
            else:
                pass
    except Exception as e:
        messagebox.showwarning('Update Warning',f'Updating is disabled, please contact help@domain_here.com\n\n{e}')
        
class S4CredutiluiApp:
    def __init__(self, master=None):
        # build ui
        self.Main_window = tk.Tk() if master is None else tk.Toplevel(master)
        self.frame23 = ttk.Frame(self.Main_window)
        self.label2 = tk.Label(self.frame23)
        self.label2.configure(font="{Calibri} 24 {bold}", text="Credentials Utility")
        self.label2.pack(pady="5 0", side="top")
        self.User_Frame = ttk.Labelframe(self.frame23)
        self.frame16 = tk.Frame(self.User_Frame)
        self.frame12 = tk.Frame(self.frame16)
        self.first_name_entry = ttk.Entry(self.frame12)
        self.first_name_str = tk.StringVar(value="First")
        self.first_name_entry.configure(textvariable=self.first_name_str, width=15)
        _text_ = "First"
        self.first_name_entry.delete("0", "end")
        self.first_name_entry.insert("0", _text_)
        self.first_name_entry.pack(expand="true", fill="x", side="top")
        self.frame12.configure(height=200, width=200)
        self.frame12.pack(expand="true", fill="x", padx=2, side="left")
        self.frame13 = tk.Frame(self.frame16)
        self.middle_name_entry = ttk.Entry(self.frame13)
        self.middle_name_entry.configure(justify="left", state="disabled", width=8)
        _text_ = "M.Initial"
        self.middle_name_entry["state"] = "normal"
        self.middle_name_entry.delete("0", "end")
        self.middle_name_entry.insert("0", _text_)
        self.middle_name_entry["state"] = "disabled"
        self.middle_name_entry.pack(expand="true", fill="x", side="top")
        self.frame13.configure(height=200, width=200)
        self.frame13.pack(expand="true", fill="x", side="left")
        self.frame14 = tk.Frame(self.frame16)
        self.last_name_entry = ttk.Entry(self.frame14)
        self.last_name_str = tk.StringVar(value="Last")
        self.last_name_entry.configure(textvariable=self.last_name_str, width=23)
        _text_ = "Last"
        self.last_name_entry.delete("0", "end")
        self.last_name_entry.insert("0", _text_)
        self.last_name_entry.pack(expand="true", fill="x", side="top")
        self.frame14.configure(height=200, width=200)
        self.frame14.pack(expand="true", fill="x", padx=2, side="left")
        self.frame16.configure(height=200, width=200)
        self.frame16.pack(fill="x", padx=2, pady=1, side="top")
        self.frame9 = tk.Frame(self.User_Frame)
        self.unlock_btn = ttk.Button(self.frame9)
        self.unlock_btn.configure(state="disabled", text="UN-LOCK")
        self.unlock_btn.pack(expand="true", fill="x", padx="0 1", side="right")
        self.unlock_btn.configure(command=self.unlock_function)
        self.lock_btn = ttk.Button(self.frame9)
        self.lock_btn.configure(text="LOCK")
        self.lock_btn.pack(expand="true", fill="x", padx="1 0", side="left")
        self.lock_btn.configure(command=self.lock_function)
        self.frame9.configure(height=200, width=200)
        self.frame9.pack(fill="x", padx=2, pady="1 5", side="top")
        self.User_Frame.configure(height=200, text="User Information", width=200)
        self.User_Frame.pack(fill="x", padx=5, pady="0 3", side="top")
        self.labelframe18 = ttk.Labelframe(self.frame23)
        self.frame43 = tk.Frame(self.labelframe18)
        self.frame44 = tk.Frame(self.frame43)
        self.label11 = ttk.Label(self.frame44)
        self.label11.configure(
            font="{Calibri Light} 12 {underline}", text="User logon name:"
        )
        self.label11.pack(anchor="w", side="top")
        self.frame47 = tk.Frame(self.frame44)
        self.ad_name_entry = ttk.Entry(self.frame47)
        self.ad_user_str = tk.StringVar(value="")
        self.ad_name_entry.configure(
            justify="right", state="disabled", textvariable=self.ad_user_str
        )
        self.ad_name_entry.pack(expand="true", fill="x", padx="2 0", side="left")
        self.label14 = ttk.Label(self.frame47)
        self.label14.configure(font="{Calibri} 12 {}", text="@domain_here.com")
        self.label14.pack(side="left")
        self.frame47.configure(height=200, width=200)
        self.frame47.pack(expand="true", fill="x", pady="2 4", side="top")
        self.frame44.configure(height=200, width=200)
        self.frame44.pack(fill="x", side="top")
        self.frame43.configure(height=200, width=200)
        self.frame43.pack(anchor="w", fill="x", padx=5, side="top")
        self.frame41 = tk.Frame(self.labelframe18)
        self.label10 = ttk.Label(self.frame41)
        self.label10.configure(
            font="{Calibri Light} 12 {underline}", text="OU Container:"
        )
        self.label10.pack(side="left")
        self.s4for_radio = ttk.Radiobutton(self.frame41)
        self.ou_container = tk.StringVar(value="")
        self.s4for_radio.configure(
            state="disabled", text="S4FOR", value="S4FOR", variable=self.ou_container
        )
        self.s4for_radio.pack(anchor="w", padx="2 0", side="left")
        self.s4mur_radio = ttk.Radiobutton(self.frame41)
        self.s4mur_radio.configure(
            state="disabled", text="S4MUR", value="S4MUR", variable=self.ou_container
        )
        self.s4mur_radio.pack(anchor="w", padx=10, side="left")
        self.s4lyn_radio = ttk.Radiobutton(self.frame41)
        self.s4lyn_radio.configure(
            state="disabled", text="S4LYN", value="S4LYN", variable=self.ou_container
        )
        self.s4lyn_radio.pack(anchor="w", side="left")
        self.frame41.configure(height=200, highlightbackground="#9d9d9d", width=200)
        self.frame41.pack(anchor="w", expand="true", padx=5, side="top")
        self.frame45 = tk.Frame(self.labelframe18)
        self.frame46 = tk.Frame(self.frame45)
        self.ad_create_btn = ttk.Button(self.frame46)
        self.ad_create_btn.configure(state="disabled", text="Create")
        self.ad_create_btn.pack(fill="x", side="top")
        self.ad_create_btn.configure(command=self.validate_ou)
        self.frame46.configure(height=200, width=200)
        self.frame46.pack(fill="x", padx=5, pady="2 5", side="top")
        self.frame45.configure(height=200, width=200)
        self.frame45.pack(fill="x", side="top")
        self.labelframe18.configure(height=200, text="Active Directory User", width=200)
        self.labelframe18.pack(
            anchor="n", expand="false", fill="x", padx=5, pady="0 3", side="top"
        )
        self.labelframe19 = ttk.Labelframe(self.frame23)
        self.frame48 = tk.Frame(self.labelframe19)
        self.label16 = ttk.Label(self.frame48)
        self.label16.configure(
            font="{Calibri Light} 12 {underline}", text="Email Address:"
        )
        self.label16.pack(anchor="w", side="top")
        self.frame49 = tk.Frame(self.frame48)
        self.email_address_entry = ttk.Entry(self.frame49)
        self.email_str = tk.StringVar(value="")
        self.email_address_entry.configure(
            justify="right", state="disabled", textvariable=self.email_str
        )
        self.email_address_entry.pack(
            expand="true", fill="x", padx="2 0", pady=0, side="left"
        )
        self.label15 = ttk.Label(self.frame49)
        self.label15.configure(font="{Calibri} 12 {}", text="@domain_here.com")
        self.label15.pack(side="left")
        self.frame49.configure(height=200, width=200)
        self.frame49.pack(fill="x", pady=1, side="top")
        self.frame48.configure(height=200, width=200)
        self.frame48.pack(anchor="w", fill="x", padx=5, side="top")
        self.frame53 = tk.Frame(self.labelframe19)
        self.frame54 = tk.Frame(self.frame53)
        self.add_email_btn = ttk.Button(self.frame54)
        self.add_email_btn.configure(state="disabled", text="Create")
        self.add_email_btn.pack(fill="both", side="top")
        self.add_email_btn.configure(command=self.validate_email)
        self.dist_groups_btn = ttk.Button(self.frame54)
        self.dist_groups_btn.configure(state="disabled", text="Dist. Groups")
        self.dist_groups_btn.pack(fill="both", side="top")
        self.dist_groups_btn.configure(command=self.add_dist_window)
        self.frame54.configure(height=200, width=200)
        self.frame54.pack(fill="x", padx=5, pady="2 5", side="top")
        self.frame53.configure(height=200, width=200)
        self.frame53.pack(fill="x", side="top")
        self.labelframe19.configure(height=200, text="0365 User", width=200)
        self.labelframe19.pack(
            anchor="n", expand="false", fill="x", padx=5, pady="0 3", side="top"
        )
        self.labelframe1 = ttk.Labelframe(self.frame23)
        self.frame11 = tk.Frame(self.labelframe1)
        self.label7 = ttk.Label(self.frame11)
        self.label7.configure(
            font="{Calibri Light} 12 {underline}", text="DSX Username:"
        )
        self.label7.pack(anchor="w", side="top")
        self.frame15 = tk.Frame(self.frame11)
        self.dsx_user_entry = ttk.Entry(self.frame15)
        self.dsx_user_entry.configure(
            justify="right", state="disabled", textvariable=self.email_str
        )
        self.dsx_user_entry.pack(
            expand="true", fill="x", padx="2 0", pady=0, side="left"
        )
        self.label8 = ttk.Label(self.frame15)
        self.label8.configure(font="{Calibri} 12 {}", text="@domain_here.com")
        self.label8.pack(side="left")
        self.frame15.configure(height=200, width=200)
        self.frame15.pack(fill="x", pady=1, side="top")
        self.frame11.configure(height=200, width=200)
        self.frame11.pack(anchor="w", fill="x", padx=5, side="top")
        self.frame19 = tk.Frame(self.labelframe1)
        self.label9 = ttk.Label(self.frame19)
        self.label9.configure(font="{Calibri Light} 12 {underline}", text="User Role:")
        self.label9.pack(anchor="w", side="top")
        self.user_role_box = ttk.Combobox(self.frame19)
        self.user_role_str = tk.StringVar(value="")
        self.user_role_box.configure(
            justify="center",
            state="disabled",
            textvariable=self.user_role_str,
            values="Viewer Shipper Manager",
        )
        self.user_role_box.pack(fill="x", pady="1 0", side="top")
        self.frame19.configure(height=200, width=200)
        self.frame19.pack(anchor="w", fill="x", padx=6, pady=2, side="top")
        self.frame18 = tk.Frame(self.labelframe1)
        self.frame8 = tk.Frame(self.frame18)
        self.headless_dsx = ttk.Checkbutton(self.frame8)
        self.headless_dsx_var = tk.IntVar()
        self.headless_dsx.configure(
            state="disabled", text="Check to disable headless mode.", variable=self.headless_dsx_var 
        )
        self.headless_dsx.pack(side="top")
        self.frame8.configure(height=200, width=200)
        self.frame8.pack(anchor="w", padx=10, side="top")
        self.add_dsx_btn = ttk.Button(self.frame18)
        self.add_dsx_btn.configure(state="disabled", text="Create")
        self.add_dsx_btn.pack(fill="both", side="top")
        self.add_dsx_btn.configure(command=self.dsx_function)
        self.frame18.configure(height=200, width=200)
        self.frame18.pack(fill="x", padx=5, pady="2 5", side="top")
        self.labelframe1.configure(height=200, text="DSX User", width=200)
        self.labelframe1.pack(
            anchor="n", expand="false", fill="x", padx=5, pady="0 3", side="top"
        )
        self.labelframe3 = ttk.Labelframe(self.frame23)
        self.frame1 = tk.Frame(self.labelframe3)
        self.frame2 = tk.Frame(self.frame1)
        self.nav_user_entry = ttk.Entry(self.frame2)
        self.nav_user_str = tk.StringVar(value="")
        self.nav_user_entry.configure(
            justify="center", state="disabled", textvariable=self.nav_user_str
        )
        self.nav_user_entry.pack(
            expand="true", fill="x", padx="2 0", pady=0, side="right"
        )
        self.label3 = ttk.Label(self.frame2)
        self.label3.configure(font="{Calibri} 12 {}", text="Username:")
        self.label3.pack(side="left")
        self.frame2.configure(height=200, width=200)
        self.frame2.pack(fill="x", padx=1, pady=1, side="top")
        self.frame1.configure(height=200, width=200)
        self.frame1.pack(anchor="w", fill="x", padx=5, side="top")
        self.frame6 = tk.Frame(self.labelframe3)
        self.frame10 = tk.Frame(self.frame6)
        self.headless_nav = ttk.Checkbutton(self.frame10)
        self.headless_nav_var = tk.IntVar()
        self.headless_nav.configure(
            state="disabled", text="Check to disable headless mode.", variable=self.headless_nav_var
        )
        self.headless_nav.pack(side="top")
        self.frame10.configure(height=200, width=200)
        self.frame10.pack(anchor="w", padx=10, side="top")
        self.add_nav_btn = ttk.Button(self.frame6)
        self.add_nav_btn.configure(state="disabled", text="Create")
        self.add_nav_btn.pack(fill="both", side="top")
        self.add_nav_btn.configure(command=self.nav_function)
        self.frame6.configure(height=200, width=200)
        self.frame6.pack(fill="x", padx=5, pady="2 5", side="top")
        self.labelframe3.configure(height=200, text="Nav User", width=200)
        self.labelframe3.pack(
            anchor="n", expand="false", fill="x", padx=5, pady="0 3", side="top"
        )
        self.labelframe17 = ttk.Labelframe(self.frame23)
        self.frame20 = tk.Frame(self.labelframe17)
        self.generate_doc_btn = ttk.Button(self.frame20)
        self.generate_doc_btn.configure(state="disabled", text="Generate")
        self.generate_doc_btn.pack(padx="1 0", pady=1, side="top")
        self.generate_doc_btn.configure(command=self.generate_doc_function)
        self.frame20.configure(height=200, width=200)
        self.frame20.pack(anchor="n", padx="2 0", pady="1 5", side="left")
        self.frame21 = tk.Frame(self.labelframe17)
        self.doc_name_entry = ttk.Entry(self.frame21)
        self.doc_title_str = tk.StringVar(value="")
        self.doc_name_entry.configure(
            justify="center", state="disabled", textvariable=self.doc_title_str
        )
        self.doc_name_entry.pack(
            expand="true", fill="x", padx=1, pady="2 1", side="top"
        )
        self.open_doc_btn = ttk.Button(self.frame21)
        self.open_doc_btn.configure(state="disabled", text="Open")
        self.open_doc_btn.pack(expand="true", fill="x", side="top")
        self.open_doc_btn.configure(command=self.open_doc_func)
        self.frame21.configure(height=200, width=200)
        self.frame21.pack(expand="true", fill="x", padx="0 2", pady="1 5", side="top")
        self.labelframe17.configure(height=200, text="Credentials Doc", width=200)
        self.labelframe17.pack(anchor="n", fill="x", padx=5, pady="0 5", side="top")
        self.frame23.configure(height=200, width=200)
        self.frame23.pack(fill="y", side="left")
        self.frame24 = ttk.Frame(self.Main_window)
        self.labelframe22 = ttk.Labelframe(self.frame24)
        self.frame40 = tk.Frame(self.labelframe22)
        self.console_display = ScrolledText(self.frame40)
        self.console_display.configure(
            autoseparators="false",
            background="#012456",
            font="systemfixed",
            foreground="#ffffff",
        )
        self.console_display.configure(relief="flat", state="disabled", wrap="word")
        self.console_display.pack(expand="true", fill="both", side="left")
        self.frame40.configure(
            height=200, highlightbackground="#c0c0c0", highlightthickness=2, width=200
        )
        self.frame40.pack(expand="true", fill="both", padx=5, pady=2, side="top")
        self.frame26 = ttk.Frame(self.labelframe22)
        self.clear_btn = ttk.Button(self.frame26)
        self.clear_btn.configure(text="Clear")
        self.clear_btn.pack(side="left")
        self.clear_btn.configure(command=self.clear_function)
        self.export_btn = ttk.Button(self.frame26)
        self.export_btn.configure(text="Export")
        self.export_btn.pack(side="left")
        self.export_btn.configure(command=self.export_function)
        self.exit_btn = ttk.Button(self.frame26)
        self.exit_btn.configure(text="Exit", command= self.Main_window.destroy)
        self.exit_btn.pack(side="left")
        self.frame26.configure(height=200, width=200)
        self.frame26.pack(anchor="e", padx=4, pady=2, side="top")
        self.labelframe22.configure(
            height=200, labelanchor="ne", text="Console", width=200
        )
        self.labelframe22.pack(
            expand="true", fill="both", padx="0 5", pady="0 5", side="top"
        )
        self.frame24.configure(height=200, width=200)
        self.frame24.pack(anchor="n", expand="true", fill="both", side="right")
        self.Main_window.geometry("1100x739")
        self.Main_window.title("S4 Credentials Utility")
      

        # Main widget
        self.mainwindow = self.Main_window

        self.first_name_entry.bind("<Button>", self.click)
        self.first_name_entry.bind("<Tab>", self.leave_last)
        self.first_name_entry.bind("<Leave>", self.leave)

        self.last_name_entry.bind("<Button>", self.click_last)
        self.last_name_entry.bind("<Tab>", self.leave_last)
        self.last_name_entry.bind("<Leave>", self.leave_last)

        self.user_role_box.current(0)
        
        self.console_display.tag_config('red', foreground='red')
        self.console_display.tag_config('yellow', foreground='yellow')
        self.console_display.tag_config('green', foreground='#00ff00')
        
        self.setup_function()

    def run(self):
        self.mainwindow.mainloop()
  
    def state_toggle(self, state_new):
        self.unlock_btn.config(state=state_new)
        self.generate_doc_btn.config(state=state_new)
        self.doc_name_entry.config(state=state_new)
        self.open_doc_btn.config(state=state_new)
        self.ad_name_entry.config(state=state_new)
        self.s4for_radio.config(state=state_new)
        self.s4mur_radio.config(state=state_new)
        self.s4lyn_radio.config(state=state_new)
        self.ad_create_btn.config(state=state_new)
        self.email_address_entry.config(state=state_new)
        self.add_email_btn.config(state=state_new)
        # self.run_all_btn.config(state=state_new)
        self.dsx_user_entry.config(state=state_new)
        self.add_dsx_btn.config(state=state_new)
        self.add_nav_btn.config(state=state_new)
        self.nav_user_entry.config(state=state_new)
        self.headless_nav.config(state=state_new)
        self.headless_dsx.config(state=state_new)
        self.dist_groups_btn.config(state=state_new)


    def run_all_window(self):
        pass
        # self.Run_all_window = tk.Toplevel()
        # self.labelframe2 = ttk.Labelframe(self.Run_all_window)
        # self.frame22 = tk.Frame(self.labelframe2)
        # self.doc_ch_btn = ttk.Checkbutton(self.frame22)
        # self.doc_ch_btn_var = tk.IntVar()
        # self.doc_ch_btn.configure(
        #     state="normal", text="Generate Doc", variable=self.doc_ch_btn_var
        # )
        # self.doc_ch_btn.pack(anchor="w", side="top")
        # self.ad_ch_btn = ttk.Checkbutton(self.frame22)
        # self.ad_ch_btn_var = tk.IntVar()
        # self.ad_ch_btn.configure(
        #     state="normal", text="AD User", variable=self.ad_ch_btn_var
        # )
        # self.ad_ch_btn.pack(anchor="w", side="top")
        # self.O365_ch_btn = ttk.Checkbutton(self.frame22)
        # self.O365_ch_btn_var = tk.IntVar()
        # self.O365_ch_btn.configure(
        #     state="normal", text="O365 User", variable=self.O365_ch_btn_var
        # )
        # self.O365_ch_btn.pack(anchor="w", side="top")
        # self.dsx_ch_btn = ttk.Checkbutton(self.frame22)
        # self.dsx_ch_btn_var = tk.IntVar()
        # self.dsx_ch_btn.configure(
        #     state="normal", text="DSX User", variable=self.dsx_ch_btn_var
        # )
        # self.dsx_ch_btn.pack(anchor="w", side="top")
        # self.frame22.configure(height=200, width=200)
        # self.frame22.pack(padx=20, pady=5, side="top")
        # self.frame4 = ttk.Frame(self.labelframe2)
        # self.button8 = ttk.Button(self.frame4)
        # self.button8.configure(text="Select All")
        # self.button8.pack(fill="x", padx=1, pady="2 0", side="top")
        # self.button8.configure(command=self.select_all)
        # self.frame4.configure(height=200, width=200)
        # self.frame4.pack(fill="x", side="top")
        # self.frame3 = ttk.Frame(self.labelframe2)
        # self.button7 = ttk.Button(self.frame3)
        # self.button7.configure(text="Run")
        # self.button7.pack(fill="x", padx=1, pady="0 2", side="top")
        # self.button7.configure(command=self.run_all_function)
        # self.frame3.configure(height=200, width=200)
        # self.frame3.pack(fill="x", side="top")
        # self.labelframe2.configure(height=200, text="Run Selected", width=200)
        # self.labelframe2.pack(padx=10, pady=10, side="top")
        # self.Run_all_window.configure(height=200, width=200)
        # self.Run_all_window.title("Run All")
        
        # self.Run_all_window.grab_set()
    global default_groups
    default_groups = None
    def add_dist_window(self):
        global default_groups
        if any(n in self.email_str.get().strip() for n in [" ","`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "–", "_", "=", "+", "[", "]", "{", "}", "\\", "|", ";", ":", "‘", "“", ",", "/", "<", ">", "?"]):
            messagebox.showerror('Warning','Please ensure that the email provided is valid')
            return
        
        global ref
        ref = []
        def add_dist_entry(event):
            if not event.keysym in ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']:
                return
            focused_widget = self.add_dist_window.focus_get()
            if not len(focused_widget.get()) > 0:
                if len(focused_widget.get()) > 1:
                    pass
                else:
                    self.dist_entry = ttk.Entry(self.frame28)
                    self.dist_entry.configure(justify="center", width=30)
                    self.dist_entry.pack(pady="2 0", side="top")
                    self.dist_entry.bind('<BackSpace>', delete_dist_entry)
                    self.dist_entry.bind('<KeyPress>', add_dist_entry)
                    ref.append(self.dist_entry)
            # self.dist_entry = ttk.Entry(self.frame28)
            # self.dist_entry.configure(justify="center", width=30)
            # self.dist_entry.pack(pady="2 0", side="top")
            # ref.append(self.dist_entry)
            
        def delete_dist_entry(event):
        
            focused_widget = self.add_dist_window.focus_get()
            if len(ref) < 2:
                pass
            elif len(focused_widget.get()) <2:
                focused_widget.destroy()
                ref.remove(focused_widget)
                ref[-1].focus_set()

            for i, e in enumerate(ref):
                if not e.get() == '':
                    bool = True
                else:
                    bool = False
            if bool == True:  
                self.dist_entry = ttk.Entry(self.frame28)
                self.dist_entry.configure(justify="center", width=30)
                self.dist_entry.pack(pady="2 0", side="top")
                self.dist_entry.bind('<BackSpace>', delete_dist_entry)
                self.dist_entry.bind('<KeyPress>', add_dist_entry)
                ref.append(self.dist_entry) 
            
            # if len(ref) < 2:
            #     pass
            # else:
            #     ref[-1].destroy()
            #     del ref[-1]
        
        self.add_dist_window = tk.Toplevel()
        self.labelframe4 = ttk.Labelframe(self.add_dist_window)
        self.frame28 = ttk.Frame(self.labelframe4)
        # self.dist_entry = ttk.Entry(self.frame28)
        # self.dist_entry.configure(justify="center", width=30)
        # self.dist_entry.pack(pady="1 0", side="top")
        self.frame28.configure(height=200, width=200)
        self.frame28.pack(padx=3, pady=3, side="top")
        # self.frame30 = ttk.Frame(self.labelframe4)
        # self.add_entry_dist_btn = ttk.Button(self.frame30)
        # self.add_entry_dist_btn.configure(text="Add +")
        # self.add_entry_dist_btn.pack(fill="x", side="top")
        # self.add_entry_dist_btn.configure(command=add_dist_entry)
        # self.delete_entry_dist_btn = ttk.Button(self.frame30)
        # self.delete_entry_dist_btn.configure(text="Delete (-)")
        # self.delete_entry_dist_btn.pack(fill="x", side="bottom")
        # self.delete_entry_dist_btn.configure(command=delete_dist_entry)
        # self.frame30.configure(height=200, width=200)
        # self.frame30.pack(anchor="s", fill="x", padx=3, pady=3, side="top")
        self.labelframe4.configure(height=200, text="Add Dist. Groups", width=200)
        self.labelframe4.pack(padx=5, pady="5 0", side="top")
        self.frame27 = ttk.Frame(self.add_dist_window)
        self.start_dist_btn = ttk.Button(self.frame27)
        self.start_dist_btn.configure(text="Run", command=self.add_dist_group)
        self.start_dist_btn.pack(fill="x", padx=5, pady="2 5", side="top")
        self.frame27.configure(height=200, width=200)
        self.frame27.pack(fill="x", side="top")
        self.add_dist_window.configure(height=200, width=200)
        self.add_dist_window.title("Add Dist. Groups")
        self.add_dist_window.geometry(f'+{self.Main_window.winfo_x()+50}+{self.Main_window.winfo_y()+280}')
        
        self.add_dist_window.grab_set()
        if default_groups == None:
            default_groups = ['GROUP1', 'GROUP2', '']
            for group in default_groups:
                self.dist_entry = ttk.Entry(self.frame28)
                self.dist_entry.configure(justify="center", width=30)
                self.dist_entry.pack(pady="2 0", side="top")
                self.dist_entry.delete("0", "end")
                self.dist_entry.insert("0", group)
                self.dist_entry.bind('<BackSpace>', delete_dist_entry)
                self.dist_entry.bind('<KeyPress>', add_dist_entry)
                ref.append(self.dist_entry)
        else:
            if not default_groups[-1] == '':
                default_groups.append('')
            for group in default_groups:
                self.dist_entry = ttk.Entry(self.frame28)
                self.dist_entry.configure(justify="center", width=30)
                self.dist_entry.pack(pady="2 0", side="top")
                self.dist_entry.delete("0", "end")
                self.dist_entry.insert("0", group)
                self.dist_entry.bind('<BackSpace>', delete_dist_entry)
                self.dist_entry.bind('<KeyPress>', add_dist_entry)
                ref.append(self.dist_entry)
    
    def select_all(self):
        self.O365_ch_btn_var.set(1)
        self.dsx_ch_btn_var.set(1)
        self.doc_ch_btn_var.set(1)
        self.ad_ch_btn_var.set(1)
        
    def unlock_function(self):
        self.state_toggle("disabled")
        self.user_role_box.config(state='disabled')
        self.lock_btn.config(state="enabled")
        self.first_name_entry.config(state='enable')
        self.last_name_entry.config(state='enable')
        self.print_console(message=f'User information unlocked>>>', tag='yellow')
        
    def lock_function(self):
        if self.last_name_entry.get() == 'Last' or self.first_name_entry.get() == 'First' or len(self.last_name_entry.get()) == 0 or len(self.first_name_entry.get()) == 0:
            messagebox.showwarning('Warning','Please enter a valid first and last name.')
            return
        elif ' ' in self.first_name_entry.get() or ' ' in self.last_name_entry.get():
            messagebox.showwarning('Warning','Please enter a valid first and last name.\n\nTip: Make sure to delete any leading or trailing spaces.')
            return
        elif any(n in self.last_name_entry.get() for n in ["`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "–", "_", "=", "+", "[", "]", "{", "}", "\\", "|", ";", ":", "‘", "“", ",", ".", "/", "<", ">", "?"]):
            messagebox.showerror("Warning","A name can't contain any of the following characters:\n< > : \ / | ? *")
            return
        elif any(n in self.first_name_entry.get() for n in ["`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "–", "_", "=", "+", "[", "]", "{", "}", "\\", "|", ";", ":", "‘", "“", ",", ".", "/", "<", ">", "?"]):
            messagebox.showerror("Warning","A name can't contain any of the following characters:\n` ~ ! @ # $ % ^ & * ( ) – _ = + [ ] { } \ | ; : ‘ “ "" ., /, <, >, ?")
            return
        else:

            self.gen_credentials(self.first_name_entry.get().strip().capitalize(),self.last_name_entry.get().strip().capitalize())
            
            self.first_name_str.set(self.first_name_entry.get().strip().capitalize())
            self.last_name_str.set(self.last_name_entry.get().strip().capitalize())
            self.first_name_entry.config(state='readonly')
            self.last_name_entry.config(state='readonly')
            self.user_role_box.config(state='readonly')

            self.print_console(message=f'User information locked>>>', tag='yellow')
            self.print_console(message=f'Preview:\n\n░░░░Full Name: {display_name}\n\n░░░░domain_here.com Domain\n\n   Username: {domain_user}\n   Password: {password}\n\n░░░░O365 Email\n\n   Email Address: {email}\n   Password: {password}\n\n░░░░Place_holder\n\n   Username: {accpac_user}\n   Password: {password}\n\n░░░░Place_holder\n\n   Username: {nav_user}\n   Password: {password}\n\n░░░░Place_holder\n\n   Username: {wims_user}\n\n░░░░Place_holder\n\n   Username: {email}\n   Password: {password}', tag='')

            self.state_toggle("enabled")

            self.doc_title_str.set(file_name)
            self.ad_user_str.set(domain_user)
            self.email_str.set(f"{first_name.lower()}.{last_name.lower()}")
            self.nav_user_str.set(nav_user)
            
            self.lock_btn.config(state="disabled")

    def gen_credentials(self, first_name_raw, last_name_raw):
        global first_name, last_name, display_name, nav_user, accpac_user, password, domain_user, email, wims_user, file_name
        first_name = first_name_raw
        last_name = last_name_raw
        display_name = f"{first_name_raw} {last_name_raw}"
        nav_user = first_name_raw[0].upper() + last_name_raw.upper()
        accpac_user = first_name_raw[0:2].upper() + last_name_raw[0:2].upper()
        password =  (''.join(random.choice(string.ascii_letters) for x in range(3)) + str(random.randint(12115,98315))).capitalize()
        domain_user = first_name_raw.upper() + last_name_raw[0].upper()
        email = f"{first_name_raw.lower()}.{last_name_raw.lower()}@domain_here.com"
        wims_user = accpac_user
        file_name = f"{first_name} {last_name}.docx"
        
    def clear_entry(self):
        self.first_name_entry.delete(0, 'end')
        self.last_name_entry.delete(0, "end")
    
    def generate_doc_function(self):
        if any(n in self.doc_title_str.get().strip() for n in ["<",'>',":","\\","/","|","?","*"]):
            messagebox.showerror("Error","A file name can't contain any of the following characters:\n< > : \ / | ? *")
            return
        if self.doc_title_str.get().strip().endswith('.docx') == False:
            self.doc_title_str.set(self.doc_title_str.get().strip() + ".docx")
        
        file_title = self.doc_title_str.get().strip()

        self.state_toggle("disabled")
        try:
                doc = DocxTemplate(r"PATH")
                context = {'nav_user': self.nav_user_str.get(), 'accpac_user': accpac_user, 'password_user': password, 'domain_user': self.ad_user_str.get(), 'email_user': self.email_str.get() + '@domain_here.com', 'wims_user': wims_user}
                doc.render(context)
                if os.path.exists(rf"PATH"):
                    choice = messagebox.askyesno('File Conflict', f'The following document "{file_title}" already exists, would you like to overwrite it? If no, please re-name the file and try again')
                    if choice == True:
                        try:
                            doc.save(rf"PATH")
                            self.print_console(f'File "{file_title}" overwritten', "yellow")

                        except Exception as e:
                            self.print_console(message=f'File not saved: {e}', tag='red')
                            messagebox.showerror('Error', f'The file could not be saved, please refer to the below error:\n\n{e}')
                    elif choice == False:
                        
                        self.print_console(f'File not saved, please re-name the file and try again','yellow')
                    
                else:
                    try:
                        doc.save(rf"PATH")
                        self.print_console(f'File "{file_title}" saved successfully', 'green')
                    except Exception as e:
                            self.print_console(message=f'File not saved: {e}', tag='red')
                            messagebox.showerror('Error', f'The file could not be saved, please refer to the below error:\n\n{e}')
                        
        except Exception as e:
            
            self.print_console(f'Task failed>>>', 'red')
            
            self.print_console(f'Exception: {e}', 'red')       
        
        self.state_toggle("enabled")

    def open_doc_func(self):
        file_title = self.doc_title_str.get().strip()
        
        try:
            os.startfile(rf"PATH")
    
            self.print_console(f'File "{file_title}" opened successfully.', 'green')
        
        except Exception as e:
    
                self.print_console(f'File could not be opened: {e}', 'red')
    
                messagebox.showerror('Error', f'The file could not be opened, please refer to the below error:\n\n{e}')
                
    def powershell_sp(self, args):
        self.state_toggle("disabled")
        for i in args:
            self.print_console(f'{i[0]}', "yellow")
            self.print_console(f'Command:\n{i[1]}', "yellow")
            try:
                def stream_process(proc):
                    go = proc.poll() is None
                    if not proc.stdout == '' or None:
                        self.print_console(proc.stdout,'')
                    if not proc.stderr == '' or None:
                        self.print_console(proc.stderr,'red')
                    if not proc.returncode is None:
                        if proc.returncode == 0:
                                pass
                        if not proc.returncode == 0:
                                self.print_console(f'Task failed/Completed with errors>>> Returncode={proc.returncode}','red')
                        
                    return go
                
                proc = subprocess.Popen(['powershell', '-windowstyle', 'hidden', '-command', i[1]], stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.DEVNULL, text=True)
                
                    
                while stream_process(proc):
                    time.sleep(0.1)
                    
            except Exception as e:
                self.print_console(f'Exception: {e}', 'red')  
                return 
            
        self.state_toggle("enabled")
       

    def powershell_function(self, input):
        start_thread=threading.Thread(target=self.powershell_sp, args=[input])
        start_thread.start() 
        
    def print_console(self, message, tag):
        
        self.console_display.configure(state='normal')
        # self.console_display.insert('end', '\n', tag)
        for line in message:
            
            self.console_display.insert('end', f'{line}', tag)
            self.console_display.see('end')
            # print(message)
            
        self.console_display.insert('end','\n\n')
        # self.console_display.delete('end',-4)
        self.console_display.see('end')
        self.console_display.configure(state='disabled')
       
        
    def run_all_command(self):
        pass
        summary = []
        self.state_toggle("disabled")
        self.user_role_box.config(state='disabled')
        
        if self.doc_ch_btn_var.get() == 1:
            try:
                self.generate_doc_function()
            except:
                summary.append(('Credentials Doc not saved','red'))
            else:
                summary.append(('Credentials Doc saved successfully','green'))
        
        if self.ad_ch_btn_var.get() == 1:
            if not self.ou_container.get() in ['S4FOR','S4MUR','S4LYN']:
                messagebox.showerror('Warning','Please select one of the OU Containers and try again')
            
            elif any(n in self.ad_user_str.get().strip() for n in [" ","`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "–", "_", "=", "+", "[", "]", "{", "}", "\\", "|", ";", ":", "‘", "“", ",", ".", "/", "<", ">", "?"]):
                messagebox.showerror('Warning','Please ensure your AD logon name is valid')
            else:
                try:
                    self.powershell_sp([('Creating user in ActiveDirectory...',f'New-ADUser -SamAccountName "{self.ad_user_str.get().strip().lower()}" -Name "{display_name}" -GivenName "{first_name}" -Surname "{last_name}" -DisplayName "{display_name}" -PasswordNeverExpires $true -UserPrincipalName "{self.ad_user_str.get().strip()}@domain_here.com" -Path "OU={self.ou_container.get()},OU=MSS,DC=domain_here,DC=com" -AccountPassword(ConvertTo-SecureString "{password}" -AsPlainText -Force) -ScriptPath "login.bat" -Enabled $true -Verbose'),
                ('Adding User to ActiveDirectory groups...',f'get-aduser -Identity TemplateUser -properties memberof -Verbose | Select-Object -ExpandProperty memberof -Verbose |  Add-ADGroupMember -Members {self.ad_user_str.get().strip().lower()} -Verbose')])
                    
                except:
                    summary.append(('AD account not created','red'))
                else:
                    summary.append(('AD account created successfully','green'))
        if self.O365_ch_btn_var.get() == 1:
            if any(n in self.email_str.get().strip() for n in [" ","`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "–", "_", "=", "+", "[", "]", "{", "}", "\\", "|", ";", ":", "‘", "“", ",", "/", "<", ">", "?"]):
                messagebox.showerror('Warning','Please ensure your AD logon name is valid')
            else:
                try:
                    self.powershell_sp([(f'Adding "{self.email_str.get().strip() + "@domain_here.com"}" to O365 tenant...',
                    f'Connect-MsolService -Verbose ; Get-MsolAccountSku -Verbose ; New-MsolUser -DisplayName "{display_name}" -FirstName {first_name} -LastName {last_name} -UserPrincipalName {self.email_str.get().strip() + "@domain_here.com"} -UsageLocation US -LicenseAssignment License_here -Password {password} -Verbose'),
                    (f'Adding "{self.email_str.get().strip() + "@domain_here.com"}" to distribution groups...', f'Import-Module ExchangeOnlineManagement ; Connect-ExchangeOnline -UserPrincipalName $name -ShowProgress $true  ; Add-DistributionGroupMember -Identity "Lynchburg" -Member "{self.email_str.get().strip() + "@domain_here.com"}" -Verbose ; Add-DistributionGroupMember -Identity "Group1" -Member "{self.email_str.get().strip() + "@domain_here.com"}" -Verbose')])
                    
                except:
                    summary.append(('0365 account not created','red'))
                else:
                    summary.append(('0365 account created successfully','green'))
        if self.dsx_ch_btn_var.get() == 1:
            try:
                self.dsx_user()
                
            except:
                summary.append(('DSX account not created','red'))
            else:
                summary.append(('DSX account created successfully','green'))
                
        self.print_console('Summary Report:','yellow')
        for i in summary:
            message,tag = i
            self.print_console(message,tag)
            
        self.user_role_box.config(state='readonly')
        self.state_toggle("enabled")
    

    def run_all_function(self):
        start_thread=threading.Thread(target=self.run_all_command)
        start_thread.start() 
    
    def clear_function(self):
        self.console_display.configure(state='normal')
        self.console_display.delete(1.0,'end')
        self.print_console(message='Ready>>>', tag='')
        self.console_display.configure(state='disabled')
        
    
    def validate_ou(self):
        if not self.ou_container.get() in ['OU1','OU2','OU3']:
            messagebox.showerror('Warning','Please select one of the OU Containers and try again')
            
        elif any(n in self.ad_user_str.get().strip() for n in [" ","`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "–", "_", "=", "+", "[", "]", "{", "}", "\\", "|", ";", ":", "‘", "“", ",", ".", "/", "<", ">", "?"]):
            messagebox.showerror('Warning','Please ensure your AD logon name is valid')
            
        else:
            self.powershell_function([('Creating user in ActiveDirectory...',f'Import-Module ActiveDirectory ; New-ADUser -SamAccountName "{self.ad_user_str.get().strip().lower()}" -Name "{display_name}" -GivenName "{first_name}" -Surname "{last_name}" -DisplayName "{display_name}" -PasswordNeverExpires $true -UserPrincipalName "{self.ad_user_str.get().strip()}@domain_here.com" -Path "OU={self.ou_container.get()},OU=MSS,DC=domain_here,DC=com" -AccountPassword(ConvertTo-SecureString "{password}" -AsPlainText -Force) -ScriptPath "login.bat" -Enabled $true -Verbose ; if ($?) {{ Write-Host ; Write-Host Adding user to default security groups ; Write-Host ; get-aduser -Identity TemplateUser -properties memberof -Verbose | Select-Object -ExpandProperty memberof -Verbose |  Add-ADGroupMember -Members {self.ad_user_str.get().strip().lower()} -Verbose ; If($?) {{Write-Host ; Write-Host Domain user account account created Successfully and added to default security groups. ; Write-Host }} Else {{ Write-Host ; Write-Host $UPN - Error: Domain user account could not be added to the default security groups... ; Write-Host }} }} Else {{  Write-Host ; Write-Host $UPN - Error: Domain user account was not created... ; Write-Host }}  '),
            ]) 

# ('Adding User to ActiveDirectory groups...',f'get-aduser -Identity TemplateUser -properties memberof -Verbose | Select-Object -ExpandProperty memberof -Verbose |  Add-ADGroupMember -Members {self.ad_user_str.get().strip().lower()} -Verbose')
    def validate_email(self):
        
        if any(n in self.email_str.get().strip() for n in [" ","`", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "–", "_", "=", "+", "[", "]", "{", "}", "\\", "|", ";", ":", "‘", "“", ",", "/", "<", ">", "?"]):
            messagebox.showerror('Warning','Please ensure that the email provided is valid')
            
        else:
            self.powershell_function([(f'Adding "{self.email_str.get().strip() + "@domain_here.com"}" to O365 tenant...',
            f'Connect-MsolService -Verbose ; Get-MsolAccountSku -Verbose ; New-MsolUser -DisplayName "{display_name}" -FirstName {first_name} -LastName {last_name} -UserPrincipalName {self.email_str.get().strip() + "@domain_here.com"} -UsageLocation US -LicenseAssignment License_here -Password {password} -Verbose ; If($?) {{ Write-Host ; Write-Host ""{self.email_str.get().strip() + "@domain_here.com"}" successfully added to O365 tenant." ; Write-Host ; Write-Host "Note: Please wait a minimum of 3-5 minutes before adding any distribution groups to this account" ; Write-Host }} Else {{  Write-Host ; Write-Host $UPN - Error: ""{self.email_str.get().strip() + "@domain_here.com"}" could not be added to O365 tenant..." ; Write-Host }}')
            ])
        
    def add_dist_group(self):
        global default_groups
        dist_cmd = []
        default_groups = []
        for self.e in ref:
            if self.e.get() == '' or None:
                pass
            else:
                default_groups.append(self.e.get().strip())
                cmd_string = f'Add-DistributionGroupMember -Identity "{self.e.get().strip()}" -Member "{self.email_str.get().strip() + "@domain_here.com"}" -Verbose ; If($?) {{ Write-Host ; Write-Host "Successfully added to ""{self.e.get().strip()}"""; Write-Host }} Else {{  Write-Host ; Write-Host $UPN "- Error: Dist. group ""{self.e.get().strip()}"" was not added..." ; Write-Host }}'
                dist_cmd.append(cmd_string)
        resp = messagebox.askyesno('Adding Distribution Groups',f'{self.email_str.get().strip() + "@domain_here.com"} will be added to the distribution groups listed below, would you like to continue?\n\n{chr(10).join(default_groups)}')
        if resp == 1:
            
            full_dist_cmd = ' ; '.join(dist_cmd)   
            full_cmd = 'Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Verbose ; Import-Module ExchangeOnlineManagement ; Connect-ExchangeOnline -UserPrincipalName $name -ShowProgress $true  ; ' + full_dist_cmd + ' ; Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser'
            self.powershell_function([(f'Adding "{self.email_str.get().strip() + "@domain_here.com"}" to distribution groups...', full_cmd)
            ])

        else:
            return
        
    def dsx_function1(self):
        start_thread=threading.Thread(target=self.dsx_user)
        start_thread.start() 
        
    def dsx_function(self):
        def login_tk_dsx():
            global admin_pass_dsx, admin_user_dsx
            admin_pass_dsx = self.admin_dsx_password.get()
            admin_user_dsx = self.admin_dsx_user.get()
            self.toplevel1.destroy()
            self.dsx_function1()
        global admin_user, admin_pass
        if admin_pass_dsx == None or admin_user_dsx == None:
            self.toplevel1 = tk.Toplevel()
            self.labelframe4 = ttk.Labelframe(self.toplevel1)
            self.frame29 = ttk.Frame(self.labelframe4)
            self.frame27 = tk.Frame(self.frame29)
            self.entry8 = ttk.Entry(self.frame27)
            self.admin_dsx_user = tk.StringVar()
            self.entry8.configure(
                justify="left", state="normal", textvariable=self.admin_dsx_user
            )
            self.entry8.pack(expand="true", fill="x", padx="2 0", pady=0, side="right")
            self.label18 = ttk.Label(self.frame27)
            self.label18.configure(font="{Calibri} 12 {}", text="Username:")
            self.label18.pack(side="left")
            self.frame27.configure(height=200, width=200)
            self.frame27.pack(fill="x", padx=1, pady=1, side="top")
            self.frame28 = tk.Frame(self.frame29)
            self.entry9 = ttk.Entry(self.frame28)
            self.admin_dsx_password = tk.StringVar()
            self.entry9.configure(
                show="•", state="normal", textvariable=self.admin_dsx_password
            )
            self.entry9.pack(expand="true", fill="x", padx="2 0", pady=0, side="right")
            self.label19 = ttk.Label(self.frame28)
            self.label19.configure(font="{Calibri} 12 {}", text="Password:")
            self.label19.pack(padx="4 0", side="left")
            self.frame28.configure(height=200, width=200)
            self.frame28.pack(fill="x", padx=1, pady=1, side="top")
            self.button3 = ttk.Button(self.frame29)
            self.button3.configure(text="Login", command=login_tk_dsx)
            self.button3.pack(fill="x", pady=5, side="top")
            self.frame29.configure(height=200, width=200)
            self.frame29.pack(fill="x", padx=5, pady=3, side="top")
            self.labelframe4.configure(height=200, text="Admin Credentials", width=200)
            self.labelframe4.pack(padx=5, pady=5, side="top")
            self.toplevel1.configure(height=200, width=200)
            self.toplevel1.geometry(f'+{self.Main_window.winfo_x()+40}+{self.Main_window.winfo_y()+480}')
            
            self.toplevel1.grab_set()
        else:
            self.dsx_function1()

    global admin_user_dsx, admin_pass_dsx
    admin_user_dsx = None
    admin_pass_dsx = None   
    def dsx_user(self):
        
        global admin_user_dsx, admin_pass_dsx
        # options = Options()
        if admin_user_dsx == '' and admin_pass_dsx == '':
            messagebox.showerror('Blank Credentials','Please enter valid admin credentials and try again.')
            admin_user_dsx = None
            admin_pass_dsx = None
            self.dsx_function()
            return        
        
        options = Options()
        
        if self.headless_dsx_var.get() == 1:
            options = webdriver.ChromeOptions()
            options.add_experimental_option("detach", True)
            options.headless = False
        else:
            options = Options()
            options.headless = True
        
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        try:
            self.print_console("Beginning DSX account creation:\n\nLogging into DSX Admin portal...","yellow")
            
            driver.get("URL")
            element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, 'Username'))
            )
            element.send_keys(f"{admin_user_dsx}")
            
            driver.find_element(By.ID,"Password").send_keys(f"{admin_pass_dsx}")
            
            driver.find_element(By.ID,"login-button").click()
            
            element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, 'layout_users'))
            )
            element.click()
            
        except Exception as e:   
            self.print_console(message=f'Login attempt for user "{self.admin_dsx_user.get()}" failed...', tag='red')  
            messagebox.showerror('Error', f'Login attempt failed, please ensure your login credentials are correct.')
            admin_user_dsx = None
            admin_pass_dsx = None
            self.dsx_function()
            return

        try:
            self.print_console("Login successful...","green")
            
            element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="manageUsers"]/div/div/div[1]/div/button'))
            )
            element.click()

            self.print_console("Inputting User credentials...","yellow")
            
            element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="new_user_modal"]/div/div[1]/form/div[2]/div/div[1]/div/div[1]/div[2]/div/input'))
            )
            element.send_keys(f'{self.email_str.get().strip() + "@domain_here.com"}')

            driver.find_element(By.XPATH,
            '//*[@id="new_user_modal"]/div/div[1]/form/div[2]/div/div[1]/div/div[1]/div[3]/div/input').send_keys(password)
            driver.find_element(By.XPATH,
            '//*[@id="new_user_modal"]/div/div[1]/form/div[2]/div/div[1]/div/div[1]/div[4]/div/input').send_keys(password)

            select = Select(driver.find_element(By.XPATH,'//*[@id="new_user_modal"]/div/div[1]/form/div[2]/div/div[1]/div/div[1]/div[5]/div/select'))

            select.select_by_visible_text(self.user_role_str.get())

            driver.find_element(By.XPATH,
            '//*[@id="new_user_modal"]/div/div[1]/form/div[2]/div/div[1]/div/div[2]/div[2]/div/input').send_keys(first_name)
            driver.find_element(By.XPATH,
            '//*[@id="new_user_modal"]/div/div[1]/form/div[2]/div/div[1]/div/div[2]/div[3]/div/input').send_keys(last_name)
            driver.find_element(By.XPATH,
            '//*[@id="new_user_modal"]/div/div[1]/form/div[2]/div/div[1]/div/div[2]/div[4]/div/div/input').send_keys('434-385-1900')
            
            
        except Exception as e:
            self.print_console(message=f'Error: {e}', tag='red')
            messagebox.showerror('Error', f'The user account was not created, please refer to the below error:\n\n{e}')
        
        if self.headless_nav_var.get() == 1:
            try:
                messagebox.showinfo('DSX account','Please submit user creation request manually from the webpage.')
                self.print_console("Please submit user creation request manually from the webpage...","yellow")
                
                
            except Exception as e:
                self.print_console(message=f'Account was not created...\n\n{e}', tag='red')
                messagebox.showerror('Error', f'The user account was not created, please refer to the below error:\n\n{e}')
        else:
            
            try:
                self.print_console("Submitting user creation request...","yellow")
                
                driver.find_element(By.XPATH,
                '//*[@id="new_user_modal"]/div/div[1]/form/div[3]/button[1]').click()
                
                text_message = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, 'noty_topRight_layout_container')) 
                )
                
                if text_message.text == f'Successfully created login for user: {self.email_str.get().strip()+"@domain_here.com"}!':
                    self.print_console(message=f'Account Status: {text_message.text}',tag="green")
                else:
                    self.print_console(message=f'Account Status: {text_message.text}',tag="red")
                
                driver.quit()
            except Exception as e:
                self.print_console(message=f'Account was not created...\n\n{e}', tag='red')
                messagebox.showerror('Error', f'The user account was not created, please refer to the below error:\n\n{e}')

    def nav_function1(self):
            start_thread=threading.Thread(target=self.nav_user)
            start_thread.start()   
    
    def nav_function(self):
        def login_tk():
            global admin_pass, admin_user
            admin_pass = self.admin_nav_password.get()
            admin_user = self.admin_nav_user.get()
            self.toplevel1.destroy()
            self.nav_function1()
        global admin_user, admin_pass
        if admin_pass == None or admin_user == None:
            self.toplevel1 = tk.Toplevel()
            self.labelframe4 = ttk.Labelframe(self.toplevel1)
            self.frame29 = ttk.Frame(self.labelframe4)
            self.frame27 = tk.Frame(self.frame29)
            self.entry8 = ttk.Entry(self.frame27)
            self.admin_nav_user = tk.StringVar()
            self.entry8.configure(
                justify="left", state="normal", textvariable=self.admin_nav_user
            )
            self.entry8.pack(expand="true", fill="x", padx="2 0", pady=0, side="right")
            self.label18 = ttk.Label(self.frame27)
            self.label18.configure(font="{Calibri} 12 {}", text="Username:")
            self.label18.pack(side="left")
            self.frame27.configure(height=200, width=200)
            self.frame27.pack(fill="x", padx=1, pady=1, side="top")
            self.frame28 = tk.Frame(self.frame29)
            self.entry9 = ttk.Entry(self.frame28)
            self.admin_nav_password = tk.StringVar()
            self.entry9.configure(
                show="•", state="normal", textvariable=self.admin_nav_password
            )
            self.entry9.pack(expand="true", fill="x", padx="2 0", pady=0, side="right")
            self.label19 = ttk.Label(self.frame28)
            self.label19.configure(font="{Calibri} 12 {}", text="Password:")
            self.label19.pack(padx="4 0", side="left")
            self.frame28.configure(height=200, width=200)
            self.frame28.pack(fill="x", padx=1, pady=1, side="top")
            self.button3 = ttk.Button(self.frame29)
            self.button3.configure(text="Login", command=login_tk)
            self.button3.pack(fill="x", pady=5, side="top")
            self.frame29.configure(height=200, width=200)
            self.frame29.pack(fill="x", padx=5, pady=3, side="top")
            self.labelframe4.configure(height=200, text="Admin Credentials", width=200)
            self.labelframe4.pack(padx=5, pady=5, side="top")
            self.toplevel1.configure(height=200, width=200)
            self.toplevel1.geometry(f'+{self.Main_window.winfo_x()+40}+{self.Main_window.winfo_y()+580}')
            
            self.toplevel1.grab_set()
        else:
            self.nav_function1()
        
    global admin_user, admin_pass
    admin_user = None
    admin_pass = None             
    def nav_user(self):
        global admin_user, admin_pass
        # options = Options()
        if admin_pass == '' and admin_user == '':
            messagebox.showerror('Blank Credentials','Please enter valid admin credentials and try again.')
            admin_user = None
            admin_pass = None
            self.nav_function()
            return
        if self.headless_nav_var.get() == 1:
            options = webdriver.ChromeOptions()
            options.add_experimental_option("detach", True)
            options.headless = False
        else:
            options = Options()
            options.headless = True
        
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


        try:
            self.print_console("Beginning NAV account creation:\n\nLogging into NAV web portal...","yellow")
            
            driver.get("PATH")
    
            user_name_field = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, 'UserName'))
            )
            user_name_field.send_keys(f"{self.admin_nav_user.get()}")
            
            driver.find_element(By.ID,"Password").send_keys(f"{self.admin_nav_password.get()}")
            
            driver.find_element(By.ID,"submitButton").click()
            
            
            WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.CLASS_NAME,"designer-client-frame")))
            
            search_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="systemaction-container"]/a[2]')))
            search_field.click()
            
        except Exception as e:
            self.print_console(message=f'Login attempt for user "{self.admin_nav_user.get()}" failed...', tag='red')
            
            messagebox.showerror('Error', f'Login attempt failed, please ensure your login credentials are correct.')
            admin_user = None
            admin_pass = None
            self.nav_function()
            return
        try:
            self.print_console("Login successful...","green")
            
            search_text_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[3]/form/div/div/div[3]/div[1]/div/div[3]/div[1]/div/div/input')))
            search_text_box.send_keys("user")

            user_option = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[3]/form/div/div/div[3]/div[1]/div/div[3]/div[4]/div[2]/table/tbody/tr[7]/td[3]/a')))
            user_option.click()
            
            new_option = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div/div[2]/ul/li[1]/span[1]/span[1]/span/span/span/a/span[1]/span/img')))
            new_option.click()
            
            user_name_create = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[4]/form/div[4]/div[1]/div[3]/div/div[1]/div[2]/div/div[1]/div/input')))
            user_name_create.click()
            time.sleep(0.2)
            user_name_create.send_keys(self.nav_user_str.get())
            
            fullname_create = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/form/div[4]/div[1]/div[3]/div/div[1]/div[2]/div/div[2]/div/input")
            fullname_create.click()
            time.sleep(0.2)
            fullname_create.send_keys(display_name)
            
            select = Select(driver.find_element(By.XPATH,'/html/body/div[1]/div[4]/form/div[4]/div[1]/div[3]/div/div[1]/div[2]/div/div[3]/div/select'))

            select.select_by_index(1)
            time.sleep(0.2)
            driver.find_element(By.XPATH,'/html/body/div[1]/div[4]/form/div[4]/div[1]/div[3]/div/div[4]/div[1]/span').click()
            time.sleep(1)
            password_set = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[4]/form/div[4]/div[1]/div[3]/div/div[4]/div[2]/div/div[1]/div/input')))
            password_set.click()
            
            password_type = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[5]/form/div/div/div[3]/div[1]/div/div[3]/div[1]/div/input')))
            password_type.click()
            time.sleep(0.2)
            password_type.send_keys(password)
            
            password_retype = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[5]/form/div/div/div[3]/div[1]/div/div[3]/div[2]/div/input')))
            password_retype.click()
            time.sleep(0.2)
            password_retype.send_keys(password)
            
            password_ok = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[5]/form/div/div/div[4]/button[1]')))
            password_ok.click()
            time.sleep(1)
            password_change_checkbox = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[4]/form/div[4]/div[1]/div[3]/div/div[4]/div[2]/div/div[2]/div/input')))
            password_change_checkbox.click()
            
            pw_1_all = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[4]/form/div[4]/div[1]/div[3]/div/div[8]/div[2]/div/div/div[2]/table/tbody/tr[1]/td[3]/input[1]')))
            pw_1_all.click()
            time.sleep(0.2)
            actions = ActionChains(driver)
            actions.send_keys('PW-1 All').perform()
    
            # actions.perform()
            actions = ActionChains(driver)
            actions.send_keys(Keys.ARROW_DOWN)
            actions.perform()
        
            time.sleep(1)
            
            actions.send_keys('PW-1-LITE').perform()
            
            # actions.perform()
            actions = ActionChains(driver)
            actions.send_keys(Keys.ARROW_DOWN)
            actions.perform()
            
            time.sleep(1)
            # actions.send_keys('PW-1 All').perform()
            
            # # actions.perform()
            # actions = ActionChains(driver)
            # actions.send_keys(Keys.ARROW_DOWN)
            # actions.perform()
            # time.sleep(1)
            if self.headless_nav_var.get() == 1:
                pass
            else:
                driver.quit()
            self.print_console(message=f'Please confirm that the account was created successfully', tag='yellow')
        except Exception as e:
            self.print_console(message=f'Task failed>>> Account was not created...\n\n', tag='red')
            messagebox.showerror('Error', f'The user account was not created')

    def export_function(self):
        log_path = filedialog.asksaveasfile(mode='w', defaultextension=".txt")
        if log_path is None: 
            return
        log_path.write(self.console_display.get(1.0, 'end'))
        log_path.close()

    def click(self, event=None):
        if self.first_name_entry.get() == 'First':
            self.first_name_entry.delete(0, 'end')
            self.Main_window.focus()

    def leave(self, event=None):
        if len(self.first_name_entry.get()) == 0:
            self.first_name_entry.delete(0, 'end')
            self.first_name_entry.insert(0, 'First')
            self.Main_window.focus()

    def click_last(self, event=None):
        if self.last_name_entry.get() == 'Last':
            self.last_name_entry.delete(0, 'end')
            self.Main_window.focus()

    def leave_last(self, event=None):
        if len(self.last_name_entry.get()) == 0:
            self.last_name_entry.delete(0, 'end')
            self.last_name_entry.insert(0, 'Last')
            self.Main_window.focus()
    
    def setup(self):
        
        if not os.path.exists(r'info\setup.ini'):
            resp = messagebox.askyesnocancel('Setup','Setup file not found; you will need to install all required dependencies if not already installed. If yes, make sure the application is launched with administrative privileges.\n\nPress “cancel” to close the application and relaunch with the correct user account control level.')
            if resp == True:
                try:
                    time.sleep(0.5)
                    self.print_console('Installing required dependencies...','yellow')
                    time.sleep(1)
                    args = [('Installing Active Directory RSAT tools...','Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Verbose ; Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Verbose ; Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online -Verbose ; get-ADDomain -Verbose ; Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser -Verbose'), ('Installing Azure AD Module...','Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Verbose ; Install-Module MSOnline -Force -Verbose ; Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser -Verbose'), ('Installing ExchangeOnlineManagement Module...','Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Verbose ; Install-Module -Name ExchangeOnlineManagement -Force -Verbose ; Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope CurrentUser -Verbose')]
                    for i in args:
                        self.print_console(f'{i[0]}', "yellow")
                        self.print_console(f'Command:\n{i[1]}', "yellow")
                        
                        def stream_process(proc):
                            go = proc.poll() is None
                            if not proc.stdout == '' or None:
                                self.print_console(proc.stdout,'')
                            if not proc.stderr == '' or None:
                                self.print_console(proc.stderr,'red')
                            if not proc.returncode is None:
                                if proc.returncode == 0:
                                        pass
                                if not proc.returncode == 0:
                                        self.print_console(f'Task failed/Completed with errors>>> Returncode={proc.returncode}','red')
            
                            return go
                        
                        proc = subprocess.Popen(['powershell', '-windowstyle', 'hidden', '-command', i[1]], stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.DEVNULL, text=True)
                        
                        while stream_process(proc):
                            time.sleep(0.1)
    
                    self.state_toggle("disabled")
                except Exception as e:
                    self.print_console(f'Required dependencies failed to install:\n\n{e}','red')
                    self.state_toggle("disabled")
                    return
                else:
                    try:
                        if not os.path.exists('info'):
                            os.mkdir('info')
                        with open(r'info\setup.ini',"w") as file_object:
                            file_object.write('')
                        self.print_console(message='Ready>>>\n\nMessage: Please check for error before continuing, If errors are present you can attempt to reinstall all dependencies by deleting the setup.ini file located in ~ S4 Credentials Utility\info', tag='yellow') 
                        self.state_toggle("disabled")
                    except Exception as e:
                        self.print_console(e,'red')
                        self.state_toggle("disabled")
            if resp == False:
                self.print_console(message='Dependency installation status unknown>>>\n\nWarning: some program features may not work as intended.', tag='red') 
                self.state_toggle("disabled")
            if resp == None:
                os._exit(1)
        else:
            self.print_console(message='Ready>>>', tag='') 
            self.state_toggle("disabled")
    
    def setup_function(self):
        start_thread=threading.Thread(target=self.setup)
        start_thread.start() 

if __name__ == "__main__":
    password = None
    update_block()
    while password != 'password':
        password = simpledialog.askstring('Security Prompt', 'Enter password to launch Credentials Utility:', show="\u2B24")
        if password == None:
            os._exit(1)
        elif password == 'password':
            break
    app = S4CredutiluiApp()
    app.run()
    
