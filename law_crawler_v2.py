import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import numpy as np
import pandas as pd
import time
import re
from tkinter import messagebox

url = "https://law.moj.gov.tw/Law/LawSearchLaw.aspx"
href_dep = "javascript:void(0);"
url_prefix = "https://law.moj.gov.tw/Law/"
rules_prefix = "https://law.moj.gov.tw/LawClass/LawAll.aspx?pcode="
#href_eye = "LawSearchLaw.aspx?TY=04013001"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Connection": "keep-alive"
}


response = requests.get(url, headers=headers)
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

    response_ = requests.get(eye_url, headers=headers)
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
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("法規爬蟲工具")

        #共用狀態(跨分頁共享，例如查詢字串)
        #self.shared = {:}
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand = True)
        self.after(100, self.show_notification)
        #各分頁
        self.page_find = FunctionFind(self.nb, self)
        self.page_import = FunctionImport(self.nb, self)
        self.nb.add(self.page_find, text="尋找法規並匯出")
        self.nb.add(self.page_import, text="從excel清單匯入")
        
    def show_notification(self):
        messagebox.showinfo("通知", "請將電腦連接個人行動熱點再執行此應用程式\n若出現Not Responding為爬蟲過程中的正常現象，請使用者耐心等候")
class FunctionFind(ttk.Frame):
    def __init__(self, parent, controller: App):
        super().__init__(parent)
        #self.root = root
        #self.root.title("法規爬蟲工具")
        self.depts = get_depts()
        self.eyes = []
        self.rules = []
        export_dept_list()
        ttk.Label(self, text = "選擇部會:").grid(row=0, column=0, padx=10, pady=10, sticky = "W")
        self.depts_menu = ttk.Combobox(self, values = list_depts)
        self.depts_menu.grid(row=0, column=1, padx=10, pady=10)
        self.depts_menu.bind("<<ComboboxSelected>>", self.refresh_depts_menu)

        ttk.Label(self, text = "選擇目:").grid(row=1, column=0, padx=10, pady=10, sticky = "W")
        self.eyes_menu = ttk.Combobox(self, values = list_eyes)
        self.eyes_menu.grid(row=1, column=1, padx=10, pady=10)
        self.eyes_menu.bind("<<ComboboxSelected>>", self.refresh_eyes_menu)

        self.frame = tk.Frame(self, bg='#09c')#blue
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
        self.button = tk.Button(self, text = "選擇所需法規", command=self.press_button)
        self.button.grid(row=2, column=0, padx=10)

        self.mode_val = tk.BooleanVar()
        radio_botton_1 = tk.Radiobutton(self, text="建立新表格", variable=self.mode_val, value=True)
        radio_botton_1.grid(row=3, column=0)
        radio_botton_2 = tk.Radiobutton(self, text="使用現有表格", variable=self.mode_val, value=False)
        radio_botton_2.grid(row=3,column=1)

        self.export_button = tk.Button(self,text = "Export", command=self.export)
        self.export_button.grid(row = 4, column=0)

        self.powered_label = tk.Label(self, text='maintained by Sam Chen - sqfh' )
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
            response_rules = requests.get(rules_prefix + a, headers=headers)
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

class FunctionImport(ttk.Frame):
    def __init__(self, parent, controller : App):
        super().__init__(parent)
        self.excel_location = tk.StringVar()
        button_location = tk.Button(self, text = "選擇excel", command=self.get_file_location).grid(row=0, column=0)
        #entry_location = tk.Entry(self).grid(row=0, column=1)
        self.mode2_val = tk.IntVar(value=0)
        radio_botton2_1 = tk.Radiobutton(self, text="建立新的表格", variable=self.mode2_val, value=1)
        radio_botton2_1.grid(row=1, column=0)
        radio_botton2_2 = tk.Radiobutton(self, text="加入現有表格", variable=self.mode2_val, value=0)
        radio_botton2_2.grid(row=1,column=1)
        button_export = tk.Button(self, text = "匯出", command=self.export_by_excellist)
        button_export.grid(row=2, column=0)
    def get_file_location(self):
            folder_pth = filedialog.askopenfilename(title="選擇excel")
            self.excel_location.set(folder_pth)
            return folder_pth
    
    def export_by_excellist(self):
        #file_pth = self.get_file_location()
        data = [] #存放法規內容
        mode2 = self.mode2_val.get()
        if mode2==1:
            folder_pth = filedialog.askdirectory(title="選擇資料夾")
        else:
            file_pth = filedialog.askopenfilename(title="選擇檔案", filetypes=[("Excel 檔案", "*.xlsx"), ("所有檔案", "*.*")])
        df = pd.read_excel(self.excel_location.get())
        #print(df)
        for i in range(len(df)):
            depts = soup.find_all("a", href=" javascript:void(0);")
            target_dept = ""
            target_eye = ""
            target_rule = ""
            for b in depts :
                if df.iloc[i,0] in b.get_text(): 
                    target_dept = b
            #print(target_dept)
            eyes = get_eyes(target_dept)
            for c in eyes :
                if df.iloc[i,1] in c.get_text():
                    target_eye = c
            #print(target_eye)
            rules = get_rules(target_eye)
            for d in rules :
                if df.iloc[i,2] in d.get_text():
                    target_rule = d
            match = re.search(r'PCode=([A-Z0-9]+)', target_rule['href'])
            time.sleep(2)
            response_rules2 = requests.get(rules_prefix + match.group(1), headers=headers)
            with open('output.html', 'w', encoding = 'utf-8') as f :
                f.write(response_rules2.text)
            soup_rules = BeautifulSoup(response_rules2.text, "html.parser")
            rules = soup_rules.find_all("div", class_ = "row")
            name = [soup_rules.find("a", id = "hlLawName").text]
            for i,a in enumerate(rules):
                a = a.text.split("\n",2)
                data.append(name + a[:-1])
            if not data[-1][-1]:
                data.pop()

        df_ = pd.DataFrame(data)
        if mode2:
            filename = time.strftime("%Y%m%d_%H%M%S.xlsx", time.localtime())
            df_.to_excel((folder_pth+"/"+filename), index=False, header=False)
        else:
            df_exist_ = pd.read_excel(file_pth, header = None)
            df_ = pd.concat([df_exist_,df_])
            df_.to_excel(file_pth, index=False, header=False)



    
#show
App().mainloop()