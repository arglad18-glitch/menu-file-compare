import tkinter as tk
from tkinter import messagebox, ttk
update_date='09.10.2026'
# =====================================================================
# СТАТИЧЕСКИЕ ИМПОРТЫ ВСЕХ РАБОЧИХ МОДУЛЕЙ
# для упаковки в один .exe
# =====================================================================
import p135_excel_worker


# =====================================================================
# СЛОВАРЬ ПРОГРАММ: (Описание, ссылка на функцию запуска)
# =====================================================================
PROGRAMS = {
    "135 Сверка двух файлов по двум первым столбцам": (
        "1. В каталоге с программой должно быть 2 сверяемых файла (.xlsx, .xls, .ods). При ином количестве программа выдаст ошибку и завершит работу.\n"
        "2. В каждом файле форматов .xlsx активным должен быть сверяемый лист (если листов несколько) — программа автоматически считает данные именно с него. Для форматов .xls и и .ods всегда считывается только первый лист книги.\n"
        "3. Сверка ведется по двум первым столбцам каждого файла (последовательность соответствующих столбцов должна совпадать). Если порядок столбцов иной, необходимо предварительно поменять столбцы местами вручную.\n"
        "4. Если нужна сверка только по 1 столбцу — обязательно добавьте второй пустой столбец.\n"
        "5. Между 2 и 3 столбцом исходных таблиц будет добавлен новый столбец с результатами сверки.\n"
        "6. Сверка идет в обе стороны. Изменения в исходные файлы на диске не вносятся — результат сверки генерируется во временных файлах.\n",
        p135_excel_worker.p135_run_excel_comparator
    ),
}

# =====================================================================
# ГРАФИЧЕСКАЯ ИНИЦИАЛИЗАЦИЯ
# =====================================================================
root = tk.Tk()
root.title("Выберите программу для запуска")
root.geometry("775x605")
root.minsize(700, 550)

lbl_info = tk.Label(
    root,
    text="1. Поместите данный файл в один каталог с исходными файлами.\n"
         "2. Запустите данный файл из каталога с исходными файлами.\n"
         "3. Выберите из списка нужную программу.\n"
         "4. Нажмите кнопку 'Запустить'",
    font=("Times New Roman", 14), justify="left", fg="#333333"
)
lbl_info.pack(anchor="w", padx=20, pady=(15, 10))

lbl_choose = tk.Label(root, text="Выберите программу для запуска:", font=("Times New Roman", 10, "bold"))
lbl_choose.pack(anchor="w", padx=20, pady=(5, 5))

combo_programs = ttk.Combobox(root, values=list(PROGRAMS.keys()), state="readonly", font=("Times New Roman", 12))
combo_programs.pack(fill="x", padx=20, pady=5)

lbl_desc = tk.Label(root, text="Описание алгоритма программы:", font=("Times New Roman", 10, "bold"))
lbl_desc.pack(anchor="w", padx=20, pady=(15, 5))

text_desc = tk.Text(root, wrap="word", font=("Times New Roman", 12), bg="#f5f5f5", height=16, bd=1, relief="solid")
text_desc.pack(fill="both", expand=True, padx=20, pady=5)
text_desc.insert("1.0", "Выберите программу из списка выше...")
text_desc.config(state=tk.DISABLED)

btn_start = tk.Button(
    root,
    text="Запустить",
    font=("Times New Roman", 10, "bold"),
    bg="#e1e1e1",
    activebackground="#cccccc",
    state=tk.DISABLED,
    width=20,
    pady=5
)
btn_start.pack(anchor="e", padx=20, pady=15)

lbl_version = tk.Label(root, text="Обновление от " + update_date, font=("Times New Roman", 10), fg="#888888")
lbl_version.place(x=20, y=575)


# =====================================================================
# ФУНКЦИИ ОБРАБОТКИ СОБЫТИЙ
# =====================================================================
def on_program_select(event):
    selected_name = combo_programs.get()
    if selected_name in PROGRAMS:
        description, _run_func = PROGRAMS[selected_name]
        text_desc.config(state=tk.NORMAL)
        text_desc.delete("1.0", tk.END)
        text_desc.insert("1.0", description)
        text_desc.config(state=tk.DISABLED)
        btn_start.config(state=tk.NORMAL)
    else:
        btn_start.config(state=tk.DISABLED)


def on_start_click():
    selected_name = combo_programs.get()
    if selected_name in PROGRAMS:
        _description, run_function = PROGRAMS[selected_name]
        try:
            run_function(root)
        except Exception as e:
            messagebox.showerror("Ошибка выполнения", f"Не удалось запустить программу:\n{str(e)}")


# =====================================================================
# СТАРТ ГЛАВНОГО ЦИКЛА ИНТЕРФЕЙСА
# =====================================================================
if __name__ == '__main__':
    combo_programs.bind("<<ComboboxSelected>>", on_program_select)
    btn_start.config(command=on_start_click)
    root.update_idletasks()
    root.mainloop()