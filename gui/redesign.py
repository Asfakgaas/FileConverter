#libraries i use
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from pathlib import Path

#GUI
win = tk.Tk()
win.title("Csv to Exel converter")
win.geometry("600x600")
win.configure(bg='magenta3')

#file_path_bar
file_lable = tk.Label(win, text="No file selected", width=40, height=2, bd=5, relief="groove")
file_lable.place(x=50, y=148)

file_lable2 = tk.Label(win, text="No file selected", width=40, height=2, bd=5, relief="groove")
file_lable2.place(x=50, y=248)

def convert():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        file_lable.config(text="Selected File: " + file_path)
    else:
        file_lable.config(text="No file selected")
    input_file = pd.read_csv(file_path)
    newfile_path = Path(file_path)
    new_ext = newfile_path.with_suffix(".xlsx")
    output_file = pd.ExcelWriter(new_ext)
    input_file.to_excel(output_file, sheet_name='Sheet1', index=False)
    output_file.close()
    lab = tk.Label(win, text = "converted")
    lab.place(x=250, y = 350)
    print("converted")

def exel_conver():
    file_path1 = filedialog.askopenfilename(filetypes=[("EXECL FILE", "*.xlsx")])
    if file_path1:
        file_lable2.config(text="Selected File: " + file_path1)
    else:
        file_lable2.config(text="No file selected")
    input_file = pd.read_excel(file_path1)
    newfile_path = Path(file_path1)
    new_ext = newfile_path.with_suffix(".csv")
    output_file = new_ext
    input_file.to_csv(output_file, index=False)
    lab = tk.Label(win, text="converted")
    lab.place(x=250, y=350)
    print("converted")


#file search button
style = ttk.Style()
style.configure("TButton", font=("bolt", 14), background="white")
#button = ttk.Button(win, text="Upload .csv file", style="TButton", command=file_search)
#button.place(x=400, y=150)
button2 = ttk.Button(win, text="convert to execl", style="TButton", command=convert)
button2.place(x= 250, y = 400)
#button3 = ttk.Button (win, text="Upload execl file", style="TButton", command=execl_file)
#button3.place(x=400, y=250)
button4 = ttk.Button(win, text="convert to csv", style="TButton", command=exel_conver)
button4.place(x=100, y=400)





win.mainloop()

