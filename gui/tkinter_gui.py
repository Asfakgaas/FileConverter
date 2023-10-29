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
win.configure(bg='black')


#file_path_bar
file_lable = tk.Label(win, text = "No file selected")
file_lable.place(x=225, y= 150)
file_lable.pack()

#search file .csv
def file_search():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        file_lable.config(text="Selected File: " + file_path)
    else:
         file_lable.config(text="No file selected")

def convert():
    input_file = pd.read_csv(file_path)
    newfile_path = Path(file_path)
    new_ext = newfile_path.with_suffix(".xlsx")
    output_file = pd.ExcelWriter(new_ext)
    input_file.to_excel(output_file, sheet_name='Sheet1', index=False)
    output_file.close()
    lab = tk.Label(win, text = "converted")
    lab.place(x=250, y = 350)
    lab.pack()
    print("converted")

#file search button
style = ttk.Style()
style.configure("TButton", font=("bolt", 14), background="white")
button = ttk.Button(win, text="Upload .csv file", style="TButton", command=file_search)

button.place(x=500, y=400)
button2 = ttk.Button(win, text="convert", style="TButton", command=convert)
button2.place(x= 400, y = 500)
button.pack()




win.mainloop()

