#libraries i use
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from pathlib import Path
import pytesseract
from PIL import Image
import pdf2docx
from pdf2docx import Converter
import os
import docx


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
def img_text():
    file_path = filedialog.askopenfilename(filetypes=())
    image = Image.open(file_path)
    text_img = pytesseract.image_to_string(image)
    new_win = tk.Tk()
    new_win.title("Text")
    new_win.geometry("600x600")
    new_win.configure(bg='white')
    #file_lable3 = tk.Label(new_win,text = text_img, width= 100, height= 100)
    text = tk.Text(new_win)
    text.insert("1.0", text_img)
    text.pack()
    #file_lable3.pack()
    new_win.mainloop()
def img_web():
    img_path = filedialog.askopenfilename(filetypes=())
    if img_web:
        web_path = filedialog.askdirectory()
        image = Image.open(img_path)
        image = image.convert('RGB')
        image.save(web_path+'/new-format.webp', 'webp')
        lab = tk.Label(win, text="converted")
        lab.place(x=250, y=350)
    else:
        pass

def pdf_word():
    pdf_path = filedialog.askopenfilename(filetypes=[("pdf file", "*.pdf")])
    if pdf_path:
        doc_path = filedialog.askdirectory()
        document = docx.Document()
        os.makedirs(doc_path, exist_ok=True)
        doc_new_path = os.path.join(doc_path, "converted.docx")
        cv = Converter(pdf_path)
        cv.convert(doc_new_path)
        lab = tk.Label(win, text="converted")
        lab.place(x=250, y=350)
        cv.close()
    else:
        pass








#file search button
style = ttk.Style()
style.configure("TButton", font=("bolt", 14), background="white")
button = ttk.Button(win, text="pdf to word", style="TButton", command=pdf_word)
button.place(x=400, y=150)
button2 = ttk.Button(win, text="convert to execl", style="TButton", command=convert)
button2.place(x= 250, y = 400)
button3 = ttk.Button (win, text="Upload img", style="TButton", command=img_text)
button3.place(x=400, y=250)
button4 = ttk.Button(win, text="convert to csv", style="TButton", command=exel_conver)
button4.place(x=100, y=400)
button5 = ttk.Button(win, text="convert to web", style="TButton", command=img_web)
button5.place(x=400, y=350)

win.mainloop()

