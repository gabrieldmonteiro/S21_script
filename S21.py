import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Border, Side
import pandas as pd
import os

def select_file1():
    file_path = filedialog.askopenfilename()
    file_path_entry1.delete(0, "end")
    file_path_entry1.insert(0, file_path)

def select_folder():
    folder_path = filedialog.askdirectory()
    folder_path_entry.delete(0, "end")
    folder_path_entry.insert(0, folder_path)

def convert_file():
    file_path = file_path_entry1.get()
    output_path = folder_path_entry.get()      
    format_file(file_path, output_path)
    show_popup()

def show_popup():
    messagebox.showinfo("S21 - Formatter", "O arquivo foi formatado! :)")    

def format_file(original_file_path: str, output_path: str):
    file_name = os.path.splitext(os.path.basename(original_file_path))[0]    
    with open(f'{original_file_path}', 'r') as f:
        with open(f'temp.csv', 'w') as f1:
            for i in range(10):
                next(f)

            for line in f:
                f1.write(line)

    df = pd.read_csv('temp.csv', encoding='cp1252', delimiter=';')
    df = df[['Order No.', 'Qtd', 'Valor', 'Lista Prç', 'Mercadoria', 'Nome Mercad.',
             'Vend A', 'Nm Vend A', 'Data da Ord.', 'Núm OC Clnt', 'Núm OE', 'Status']]
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    
    medium_border = Border(left=Side(style='medium'),
                        right=Side(style='medium'),
                        top=Side(style='medium'),
                        bottom=Side(style='medium'))

    
    for row in ws.iter_rows(min_row=2, min_col=1):
        for cell in row:
            cell.border = medium_border

    
    wb.save(f'{output_path}/{file_name}.xlsx')    
    os.remove("temp.csv")

root = tk.Tk()
root.title("S21 - Formatter")
width = 500
height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (width // 2)
y = (screen_height // 2) - (height // 2)
root.geometry('{}x{}+{}+{}'.format(width, height-180, x, y))
root.resizable(width=False, height=False)

file_path_entry1 = tk.Entry(root, width=50)
file_path_entry1.grid(row=0, column=1, padx=5, pady=5)

select_file_button1 = tk.Button(root, text="Selecionar Arquivo", command=select_file1)
select_file_button1.grid(row=0, column=0, padx=5, pady=5)

folder_path_entry = tk.Entry(root, width=50)
folder_path_entry.grid(row=1, column=1, padx=5, pady=5)

select_folder_button = tk.Button(root, text="Selecionar Pasta de Destino", command=select_folder)
select_folder_button.grid(row=1, column=0, padx=5, pady=5)

convert_button1 = tk.Button(root, text="Converter Arquivo", command=convert_file)
convert_button1.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

file_path_entry1.config(width=50)
folder_path_entry.config(width=50)

root.mainloop()
