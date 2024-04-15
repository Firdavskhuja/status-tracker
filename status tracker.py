import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Отслеживание товаров")
        self.track_entries = []
        self.columns = [
            "№",
            "Трек номер",
            "Статус",
            "Дата добавления",
            "Дата изменения",
            "Вес (кг)",
            "Куб. м³"
        ]
        self.create_widgets()

    def create_widgets(self):
        # Верхняя часть окна
        top_frame = ttk.Frame(self.root, padding="20")
        top_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        ttk.Label(top_frame, text="Трек номер:").grid(row=0, column=0, padx=10, pady=5)
        self.track_number_entry = ttk.Entry(top_frame, width=50)
        self.track_number_entry.grid(row=0, column=1, padx=10, pady=5)

        add_button = ttk.Button(top_frame, text="Добавить", command=self.add_track)
        add_button.grid(row=0, column=2, padx=10, pady=5)

        change_status_button = ttk.Button(top_frame, text="Изменить статус", command=self.change_status)
        change_status_button.grid(row=0, column=3, padx=10, pady=5)

        edit_button = ttk.Button(top_frame, text="Редактировать", command=self.edit_item)
        edit_button.grid(row=0, column=4, padx=10, pady=5)

        search_button = ttk.Button(top_frame, text="Найти", command=self.search_track)
        search_button.grid(row=0, column=5, padx=10, pady=5)

        # Комбо-бокс для фильтрации статуса
        ttk.Label(top_frame, text="Фильтр по статусу:").grid(row=1, column=0, padx=10, pady=5)
        self.status_filter_combo = ttk.Combobox(top_frame, values=["Все", "Не доставлено", "Доставлено"])
        self.status_filter_combo.set("Все")
        self.status_filter_combo.grid(row=1, column=1, padx=10, pady=5)
        self.status_filter_combo.bind("<<ComboboxSelected>>", self.filter_by_status)

        # Нижняя часть окна (таблица)
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(bottom_frame, columns=self.columns, show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Добавляем горизонтальный и вертикальный скроллбары
        #xscrollbar = ttk.Scrollbar(bottom_frame, orient="horizontal", command=self.tree.xview)
        #xscrollbar.pack(side=tk.BOTTOM, fill=tk.Y)
        #self.tree.configure(xscrollcommand=xscrollbar.set)

        yscrollbar = ttk.Scrollbar(bottom_frame, orient="vertical", command=self.tree.yview)
        yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=yscrollbar.set)

        # Настройка заголовков таблицы
        for col in self.columns:
            self.tree.heading(col, text=col, anchor=tk.CENTER)
            self.tree.column(col, anchor=tk.CENTER)  # Выравнивание содержимого по центру

        # Применяем стиль для заголовков таблицы
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Times New Roman", 14, "bold"))  # Жирный шрифт для заголовков

        # Применяем стиль Times New Roman размером 14 к остальным элементам
        style.configure("TLabel", font=("Times New Roman", 14))  # Метки
        style.configure("TButton", font=("Times New Roman", 14))  # Кнопки

        # Применяем стиль к содержимому таблицы (Treeview)
        style.configure("Treeview", font=("Times New Roman", 14))  # Шрифт для содержимого таблицы

        self.load_data_from_excel()  # Загружаем данные из Excel при запуске приложения

    def add_track(self):
        track_number = self.track_number_entry.get().strip()
        if track_number:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_entry = {
                "№": len(self.track_entries) + 1,
                "Трек номер": track_number,
                "Статус": "Не доставлено",
                "Дата добавления": current_time,
                "Дата изменения": "",
                "Вес (кг)": "",
                "Куб. м³": ""
            }
            self.track_entries.append(new_entry)
            self.update_treeview()
            self.save_to_excel()  # Сохраняем данные в Excel после добавления
            self.track_number_entry.delete(0, tk.END)  # Очищаем поле ввода

    def change_status(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_id = int(self.tree.item(selected_item, "text")) - 1
            current_status = self.track_entries[item_id]["Статус"]
            if current_status == "Не доставлено":
                self.track_entries[item_id]["Статус"] = "Доставлено"
                self.track_entries[item_id]["Дата изменения"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # Запрашиваем вес товара только при статусе "Доставлено"
                weight = simpledialog.askfloat("Вес товара", "Введите вес товара в кг:")
                if weight is not None:
                    self.track_entries[item_id]["Вес (кг)"] = weight
                    # Запрашиваем объем товара в кубических метрах
                    volume = simpledialog.askfloat("Объем товара", "Введите объем товара в куб. м³:")
                    if volume is not None:
                        self.track_entries[item_id]["Куб. м³"] = volume
                        self.update_treeview()
                        self.save_to_excel()  # Сохраняем данные в Excel после добавления объема товара
            elif current_status == "Доставлено":
                messagebox.showinfo("Информация", "Товар уже доставлен.")

    def edit_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_id = int(self.tree.item(selected_item, "text")) - 1
            # Отображаем диалоговое окно для редактирования записи
            new_values = self.edit_dialog(self.track_entries[item_id])
            if new_values is not None:
                self.track_entries[item_id].update(new_values)
                self.update_treeview()
                self.save_to_excel()  # Сохраняем данные в Excel после редактирования

    def edit_dialog(self, entry_values):
        # Создаем диалоговое окно для редактирования
        dialog = tk.Toplevel(self.root)
        dialog.title("Редактирование записи")

        # Создаем и располагаем метки и поля ввода для каждой колонки
        entries = {}
        for i, col in enumerate(self.columns):
            if col == "№":
                continue  # Пропускаем колонку с номером, так как она не редактируется
            ttk.Label(dialog, text=col).grid(row=i, column=0, padx=10, pady=5)
            entry = ttk.Entry(dialog, width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(tk.END, entry_values.get(col, ""))  # Заполняем поле текущим значением
            entries[col] = entry

        # Функция сохранения изменений
        def save_changes():
            new_values = {col: entry.get() for col, entry in entries.items()}
            dialog.destroy()
            return new_values

        # Кнопка сохранения изменений
        save_button = ttk.Button(dialog, text="Сохранить", command=save_changes)
        save_button.grid(row=len(self.columns), column=0, columnspan=2, padx=10, pady=10)

        dialog.transient(self.root)  # Устанавливаем родительское окно
        dialog.grab_set()  # Блокируем взаимодействие с основным окном
        self.root.wait_window(dialog)  # Ожидаем закрытия диалогового окна

    def search_track(self):
        search_query = self.track_number_entry.get().strip()
        if search_query:
            filtered_entries = [
                entry for entry in self.track_entries if str(entry["Трек номер"]).lower() == search_query.lower()
            ]
            if filtered_entries:
                self.tree.delete(*self.tree.get_children())
                for entry in filtered_entries:
                    values = [entry.get(col, "") for col in self.columns]
                    self.tree.insert("", "end", text=entry["№"], values=values)
                self.track_number_entry.delete(0, tk.END)
                self.track_number_entry.insert(0, search_query)
            else:
                messagebox.showinfo("Поиск", f"Трек номер '{search_query}' не найден.")

    def filter_by_status(self, event):
        self.update_treeview()

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for entry in self.track_entries:
            if self.status_filter_combo.get() == "Все" or entry.get("Статус") == self.status_filter_combo.get():
                values = [entry.get(col, "") for col in self.columns]
                self.tree.insert("", "end", text=entry["№"], values=values)

    def save_to_excel(self):
        df = pd.DataFrame(self.track_entries)
        excel_file_path = "inventory_data.xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title = "Inventory Data"

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(name='Palatino Linotype', size=14)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        wb.save(filename=excel_file_path)
        messagebox.showinfo("Сохранение", f"Данные сохранены в Excel файл: {excel_file_path}")

    def load_data_from_excel(self):
        try:
            df = pd.read_excel("inventory_data.xlsx")
            self.track_entries = df.to_dict(orient="records")
            self.update_treeview()
            messagebox.showinfo("Загрузка", "Данные загружены из Excel файла.")
        except FileNotFoundError:
            messagebox.showwarning("Внимание", "Файл Excel не найден. Создан новый список трек-номеров.")

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()
