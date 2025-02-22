import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import ttk
import threading
import pandas as pd
import re
from concurrent.futures import ThreadPoolExecutor

BASE_DIR = ""
excel_path = ""
template_path = ""

def browse_base_dir():
    global BASE_DIR
    BASE_DIR = filedialog.askdirectory()
    if BASE_DIR:
        base_dir_label.config(text=f"Base Directory: {BASE_DIR}", foreground="#1E90FF")
        base_dir_entry.delete(0, tk.END)
        base_dir_entry.insert(0, BASE_DIR)

def drop_file(event, file_type):
    file_path = event.data.strip().replace('{', '').replace('}', '')
    global excel_path, template_path
    
    if file_type == "excel":
        excel_path = file_path
        excel_label.config(text=f"Selected: {os.path.basename(excel_path)}", foreground="#1E90FF")
    elif file_type == "template":
        template_path = file_path
        template_label.config(text=f"Selected: {os.path.basename(template_path)}", foreground="#1E90FF")

def upload_excel():
    global excel_path
    excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if excel_path:
        excel_label.config(text=f"Selected: {os.path.basename(excel_path)}", foreground="#1E90FF")

def upload_template():
    global template_path
    template_path = filedialog.askopenfilename(filetypes=[("CONF Files", "*.conf")])
    if template_path:
        template_label.config(text=f"Selected: {os.path.basename(template_path)}", foreground="#1E90FF")

def generate_conf_file(device_name, token, template_content):
    if not BASE_DIR:
        return
    
    folder_path = os.path.join(BASE_DIR, f"192.168.0.{token}", "OA", "use", "cpu")
    os.makedirs(folder_path, exist_ok=True)
    output_file = os.path.join(folder_path, "qemgr.conf")
    
    modified_content = re.sub(r"<item node='[^']+' label='DNA Operate [^']+'", 
                              f"<item node='{device_name}' label='DNA Operate {device_name}'", 
                              template_content)
    
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(modified_content)

def process_files():
    if not BASE_DIR:
        messagebox.showerror("Error", "Please select a base directory")
        return
    
    if not excel_path or not template_path:
        messagebox.showerror("Error", "Please select both Excel and template files")
        return
    
    try:
        df = pd.read_excel(excel_path, engine='openpyxl', usecols=["Device", "Token"])
        devices = df["Device"].values
        tokens = df["Token"].values
        
        with open(template_path, "r", encoding="utf-8") as file:
            template_content = file.read()
        
        with ThreadPoolExecutor() as executor:
            executor.map(generate_conf_file, devices, tokens, [template_content] * len(devices))
        
        messagebox.showinfo("Success", "CONF files generated successfully! 💙")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def process_files_threaded():
    threading.Thread(target=process_files, daemon=True).start()

root = TkinterDnD.Tk()
root.title("SysGeneratorFiles")
root.geometry("600x550")
root.configure(bg="#E0F7FA")

style = ttk.Style()
style.theme_use('clam')
style.configure("TFrame", background="#E0F7FA")
style.configure("TLabel", font=("Comic Sans MS", 12), background="#E0F7FA", foreground="#00008B")
style.configure("TButton", font=("Comic Sans MS", 10, "bold"), padding=5, background="#1E90FF", foreground="#FFFFFF")
style.map("TButton", background=[["active", "#4169E1"]])

frame = ttk.Frame(root, padding=20, style="TFrame")
frame.pack(expand=True)

title_label = ttk.Label(frame, text="SysGenerated Files 💾", font=("Comic Sans MS", 18, "bold"), foreground="#1E90FF")
title_label.pack(pady=10)

base_dir_frame = ttk.Frame(frame)
base_dir_frame.pack(pady=5, fill=tk.X)

base_dir_label = ttk.Label(base_dir_frame, text="Select Base Directory 📁", font=("Comic Sans MS", 12))
base_dir_label.pack(side=tk.LEFT)
base_dir_entry = ttk.Entry(base_dir_frame, width=50)
base_dir_entry.pack(side=tk.LEFT, padx=5)
browse_button = ttk.Button(base_dir_frame, text="...", command=browse_base_dir, width=3)
browse_button.pack(side=tk.LEFT)

excel_label = ttk.Label(frame, text="Drag & Drop or Upload Excel File 📂", font=("Comic Sans MS", 12))
excel_label.pack(pady=5)
excel_label.drop_target_register(DND_FILES)
excel_label.dnd_bind('<<Drop>>', lambda e: drop_file(e, "excel"))
ttk.Button(frame, text="Upload Excel 📄", command=upload_excel).pack(pady=5)

template_label = ttk.Label(frame, text="Drag & Drop or Upload Template File 📄", font=("Comic Sans MS", 12))
template_label.pack(pady=5)
template_label.drop_target_register(DND_FILES)
template_label.dnd_bind('<<Drop>>', lambda e: drop_file(e, "template"))
ttk.Button(frame, text="Upload Template 📂", command=upload_template).pack(pady=5)

generate_button = ttk.Button(frame, text="Generate CONF Files ✨", command=process_files_threaded)
generate_button.pack(pady=20)

root.mainloop()
