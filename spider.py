import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import numpy as np
import pandas as pd
import time
import re


url = "https://law.moj.gov.tw/Law/LawSearchLaw.aspx"
href_dep = "javascript:void(0);"
url_prefix = "https://law.moj.gov.tw/Law/"
rules_prefix = "https://law.moj.gov.tw/LawClass/LawAll.aspx?pcode="
#href_eye = "LawSearchLaw.aspx?TY=04013001"

response = requests.get(url)
#print(response.text)
with open('output.html', 'w', encoding = 'utf-8') as f :
    f.write(response.text)

soup = BeautifulSoup(response.text, "html.parser")
depts = soup.find_all("a", href=" javascript:void(0);")

def get_depts():  
    soup = BeautifulSoup(response.text, "html.parser")
    depts = soup.find_all("a", href=" javascript:void(0);")
    return depts

def get_eyes(dept):
    eyes = dept.find_next_sibling("ul")
    eyes = eyes.find_all("a")
    return eyes

def get_rules(eye):
    eye_url = url_prefix + eye['href']

    response_ = requests.get(eye_url)
    with open('output_.html', 'w', encoding = 'utf-8') as f :
        f.write(response_.text)
    soup_ = BeautifulSoup(response_.text, "html.parser")

    rules = soup_.find_all("a", id = "hlkLawName")
    return rules

def trace_dept(dept):
    return soup.find("a",string = dept)
    
def trace_eye(dept, eye):
    eyes = soup.find("a", string=dept)

    ul = eyes.find_next_sibling("ul")
    if not ul:
        return None

    for a in ul.find_all("a"):
        if eye in a.get_text():  # 用 get_text() 判斷是否包含 eye
            return a
    return None

list_depts = []
list_eyes = []
list_rules = []
def export_dept_list():
    global list_depts
    list_depts = []
    list_rules = []
    depts = get_depts()
    for i,a in enumerate(depts):
        list_depts.append(a.text)

def export_eyes_list(eyes):
    global list_eyes
    list_eyes = []
    if eyes:
        for i, a in enumerate(eyes):
            list_eyes.append(a.text)

def export_rules_list(rules):
    global list_rules
    if rules:
        for i, a in enumerate(rules):
            list_rules.append(a.text)

class SpiderGui:
    def __init__(self, root):
        self.root = root
        self.root.title("法規爬蟲工具")
        self.depts = get_depts()
        self.eyes = []
        self.rules = []
        export_dept_list()
        ttk.Label(root, text = "選擇部會:").grid(row=0, column=0, padx=10, pady=10, sticky = "W")
        self.depts_menu = ttk.Combobox(root, values = list_depts)
        self.depts_menu.grid(row=0, column=1, padx=10, pady=10)
        self.depts_menu.bind("<<ComboboxSelected>>", self.refresh_depts_menu)

        ttk.Label(root, text = "選擇目:").grid(row=1, column=0, padx=10, pady=10, sticky = "W")
        self.eyes_menu = ttk.Combobox(root, values = list_eyes)
        self.eyes_menu.grid(row=1, column=1, padx=10, pady=10)
        self.eyes_menu.bind("<<ComboboxSelected>>", self.refresh_eyes_menu)

        self.frame = tk.Frame(root, bg='#09c')#blue
        self.frame.grid(row = 2,column=1, padx=10, pady=10, sticky="W")
        self.frame.grid_rowconfigure(0,weight=1)
        self.frame.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.frame)
        self.scrollbar = tk.Scrollbar(self.frame, orient="vertical", command = self.canvas.yview)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1,sticky="ns")
        self.scrollable_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0,0),window=self.scrollable_frame,anchor="nw") 
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        def on_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.scrollable_frame.bind("<Configure>", on_configure)

        self.checks = []
        self.button = tk.Button(root, text = "選擇所需法規", command=self.press_button)
        self.button.grid(row=2, column=0, padx=10)

        self.mode_val = tk.BooleanVar()
        radio_botton_1 = tk.Radiobutton(root, text="建立新表格", variable=self.mode_val, value=True)
        radio_botton_1.grid(row=3, column=0)
        radio_botton_2 = tk.Radiobutton(root, text="使用現有表格", variable=self.mode_val, value=False)
        radio_botton_2.grid(row=3,column=1)

        self.export_button = tk.Button(root,text = "Export", command=self.export)
        self.export_button.grid(row = 4, column=0)

        self.powered_label = tk.Label(root, text='maintained by Sam Chen - sqfh' )
        self.powered_label.grid(row = 5, column = 1,sticky="e")
        


    def refresh_depts_menu(self, event):
        #print(trace_dept(self.depts_menu.get()))
        self.eyes = get_eyes(trace_dept(self.depts_menu.get()))
        export_eyes_list(self.eyes)
        self.eyes_menu['values'] = list_eyes
        return
    def refresh_eyes_menu(self, event):
        self.eyes = get_eyes(trace_dept(self.depts_menu.get()))
        export_eyes_list(self.eyes)
        self.eyes_menu['values'] = list_eyes

    def press_button(self):
        self.checks.clear()
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.rules = get_rules(trace_eye(self.depts_menu.get(),self.eyes_menu.get()))
        #print(self.rules)
        for i,a  in enumerate(self.rules):
            var = tk.BooleanVar()
            self.check_botton = tk.Checkbutton(self.scrollable_frame, text = a.text, variable = var)
            self.check_botton.grid(row = (i), column=1, padx=10, pady=10, sticky = "W")
            self.checks.append((var,a))
    def export(self):
        mode = self.mode_val.get()
        selected_items = [a for var,a in self.checks if var.get()]
        data = [] #存放Pcode
        table = []
        for a in selected_items:
            href = a['href']
            match = re.search(r'PCode=([A-Z0-9]+)', href)
            data.append(match.group(1))
        for a in data:
            time.sleep(2)
            response_rules = requests.get(rules_prefix + a)
            with open('output.html', 'w', encoding = 'utf-8') as f :
                f.write(response_rules.text)
            soup_rules = BeautifulSoup(response_rules.text, "html.parser")

            rules = soup_rules.find_all("div", class_ = "row")
            name = [soup_rules.find("a", id = "hlLawName").text]
            for i,a in enumerate(rules):
                a = a.text.split("\n",2)
                table.append(name + a[:-1])
            if not table[-1][-1]:
                table.pop()

        df = pd.DataFrame(table)
        if mode:
            folder_pth = filedialog.askdirectory(title="選擇資料夾")
            filename = time.strftime("%Y%m%d_%H%M%S.xlsx", time.localtime())
            df.to_excel((folder_pth+"/"+filename), index=False, header=False)
        else:
            file_pth = filedialog.askopenfilename(title="選擇檔案", filetypes=[("Excel 檔案", "*.xlsx"), ("所有檔案", "*.*")])
            df_exist = pd.read_excel(file_pth, header = None)
            df = pd.concat([df_exist,df])
            df.to_excel(file_pth, index=False, header=False)


    
#show
root = tk.Tk()
app = SpiderGui(root)
root.mainloop()