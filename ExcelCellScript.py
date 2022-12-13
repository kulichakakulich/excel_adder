import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showwarning, showinfo

import openpyxl


class Excel:
    def __init__(self, fnt, fne, cell, data, jank):
        self.fnt = ""
        self.fne = ""
        self.cell = ""
        self.data = []
        self.jank = 0


def write_adds():
    ent_cell = ent_c.get()
    if Excel.jank == 1:
        write_file_adds(ent_cell)
    else:
        write_oak(ent_cell)


def write_oak(ent_cell):
    path = Excel.fne  # название файла
    wb = openpyxl.load_workbook(path)
    sheets = wb.worksheets
    for sheet, i in zip(sheets, range(150)):  # range - нужное количество записей
        sheet[f"{ent_cell}"] = None
        sheet[f"{ent_cell}"] = f"#{i + 1}/АООК инв. 50232"
    wb.save(path)
    showinfo("Выполнено", "Листы пронумерованы AООК")


def write_file_adds(ent_cell):
    path = Excel.fne  # название файла
    wb = openpyxl.load_workbook(path)
    sheets = wb.worksheets
    data = Excel.data = []
    with open(Excel.fnt, "r+", encoding="utf-8") as f:
        data = [line.rstrip('\n') for line in f]
    f.close()
    if len(data) > 2:
        for sheet, i in zip(sheets, data):  # Для вставки чисел из текстового файла
            sheet[f"{ent_cell}"] = None
            sheet[f"{ent_cell}"] = i
        wb.save(path)
        showinfo("Выполнено", "Данные записаны в файл")
    else:
        showwarning("Ошибка", "Пустой файл")
    Excel.jank = 0
    Excel.fnt = None


def file_read_e():
    Excel.fne = askopenfilename(defaultextension="xlsx")
    if not Excel.fne or not "*.xlsx":
        showwarning("Предупреждение", "Не выбрана таблица")


def file_read_t():
    Excel.fnt = askopenfilename(defaultextension="txt")
    Excel.jank = 1
    if not Excel.fnt or not "*.txt":
        showwarning("Предупреждение", "Не выбран текстовый файл")


window = tk.Tk()
window.title("Excel reader script")
window.geometry("200x200")
window.resizable(width=False, height=False)
canvas = tk.Canvas(window, height=400, width=300)
canvas.pack()
frame = tk.Frame(window)
frame.place(relheight=1, relwidth=1)
Excel.jank = 0
btn_excelFile = tk.Button(frame, text="Выбери excel", command=file_read_e)
btn_excelFile.pack(anchor="n", padx=10, pady=5)
btn_txtFile = tk.Button(frame, text="Выбери txt файл", command=file_read_t)
btn_txtFile.pack(anchor="n", padx=10, pady=10)
lbl_name = tk.Label(frame, text="Введи номер ячейки:")
lbl_name.pack(anchor="n", padx=10, pady=2)
ent_c = tk.Entry(frame)
ent_c.pack(anchor="n", padx=10, pady=2)
btn_confirm = tk.Button(frame, text="Выполнить", command=write_adds)
btn_confirm.pack(anchor="n", padx=10, pady=15)

window.mainloop()
