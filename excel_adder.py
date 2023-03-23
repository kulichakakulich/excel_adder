import openpyxl
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showwarning, showinfo


class Excel:
    def __init__(self):
        self.dataT = ""
        self.dataE = ""
        self.cell_name = ""
        self.count = 0

    def file_read_e(self):
        self.dataE = askopenfilename(
            filetypes=[("Excel files", ".xlsx")])
        if not self.dataE or not "*.xlsx":
            showwarning("Предупреждение", "Не выбрана таблица")

    def file_read_t(self):
        self.dataT = askopenfilename(filetypes=[("Text files", "*.txt")])
        self.count = 1
        if not self.dataT or not "*.txt":
            showwarning("Предупреждение", "Не выбран текстовый файл")

    def write_adds(self, cell_name):
        if self.count == 1:
            self.write_file_adds(cell_name)
        else:
            self.write_oak(cell_name)

    def write_oak(self, cell_name):
        wb = openpyxl.load_workbook(self.dataE)
        sheets = wb.worksheets
        for sheet, i in zip(
                sheets, range(150)):
            sheet[cell_name] = None
            sheet[cell_name] = f"#{i + 1}/АООК инв. 50232"
        wb.save(self.dataE)
        showinfo("Выполнено", "Листы пронумерованы AООК")

    def write_file_adds(self, ent_cell):
        wb = openpyxl.load_workbook(self.dataE)
        sheets = wb.worksheets
        data = []
        with open(self.dataT, "r+", encoding="utf-8") as f:
            data = [line.rstrip('\n') for line in f]
        f.close()
        if len(data) > 2:
            for sheet, i in zip(
                    sheets, data):
                sheet[f"{ent_cell}"] = None
                sheet[f"{ent_cell}"] = i
            wb.save(self.dataE)
            showinfo("Выполнено", "Данные записаны в файл")
        else:
            showwarning("Ошибка", "Пустой файл")
        self.count = 0
        self.dataT = None


def main():
    window = tk.Tk()
    window.title("Excel reader script")
    window.geometry("200x200")
    window.resizable(width=False, height=False)
    canvas = tk.Canvas(window, height=400, width=300)
    canvas.pack()
    frame = tk.Frame(window)
    frame.place(relheight=1, relwidth=1)
    exc = Excel()

    btn_excelFile = tk.Button(
        frame, text="Выбери excel", command=exc.file_read_e)
    btn_excelFile.pack(anchor="n", padx=10, pady=5)
    btn_txtFile = tk.Button(
        frame, text="Выбери txt файл", command=exc.file_read_t)
    btn_txtFile.pack(anchor="n", padx=10, pady=10)

    lbl_name = tk.Label(frame, text="Введи номер ячейки:")
    lbl_name.pack(anchor="n", padx=10, pady=2)
    cell_lbl_name = tk.Entry(frame)
    cell_lbl_name.pack(anchor="n", padx=10, pady=2)
    btn_confirm = tk.Button(frame, text="Выполнить",
                            command=lambda: exc.write_adds(cell_lbl_name.get()))
    btn_confirm.pack(anchor="n", padx=10, pady=15)

    window.mainloop()


if __name__ == "__main__":
    main()
