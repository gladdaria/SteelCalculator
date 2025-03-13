import json
import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

def resource_path(relative_path):
        """Преобразует относительный путь в абсолютный для PyInstaller"""
        try:
            # PyInstaller создает временную папку в _MEIPASS
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

class SteelAlphaCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор длины порезки")
        self.session_data = []
        
        # Инициализация данных
        self.data = self.load_data()
        self.temperatures = self.parse_temperatures()
        self.all_grades = [grade["steel_grade"] for grade in self.data["steel_grades"]]
        
        # Загрузка иконок
        self.icons = self.load_icons()
        
        # Построение интерфейса
        self.create_widgets()
        self.configure_bindings()


    def load_icons(self):
        """Загрузка иконок из папки icons с обработкой ошибок"""
        icons = {
            'calculate': None,
            'reset': None,
            'add': None,
            'export': None,
            'clear': None
        }
        
        icon_mapping = {
            'calculate': 'icons/calculator.ico',
            'reset': 'icons/undo.ico',
            'add': 'icons/add.ico',
            'export': 'icons/excel.ico',
            'clear': 'icons/clean.ico'
        }
        
        try:
            icons_dir = os.path.join(os.path.dirname(__file__), "icons")
            for name, filename in icon_mapping.items():
                path = os.path.join(icons_dir, filename)
                if os.path.exists(path):
                    icons[name] = tk.PhotoImage(file=resource_path(filename))
        except Exception as e:
            messagebox.showwarning("Ошибка иконок", f"Не удалось загрузить иконки: {str(e)}")
        
        return icons

    def load_data(self):
        """Загрузка данных из JSON-файла"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            json_path = resource_path(os.path.join("data", "steel_data.json"))
            
            with open(json_path, "r", encoding="utf-8") as file:
                data = json.load(file)
                
                if "temperature_values" not in data or "steel_grades" not in data:
                    raise ValueError("Некорректный формат JSON")
                
                return data
        except Exception as e:
            messagebox.showerror("Ошибка данных", f"Ошибка загрузки: {str(e)}")
            return {"temperature_values": [], "steel_grades": []}

    def parse_temperatures(self):
        """Преобразование температур в отсортированный список"""
        return sorted([int(temp) for temp in self.data["temperature_values"]])

    def create_widgets(self):
        """Построение графического интерфейса"""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        # Блок ввода параметров
        self.create_input_frame(main_frame)
        
        # Блок результатов
        self.create_results_frame(main_frame)
        
        # Таблица сессии
        self.create_session_frame(main_frame)
        
        # Всплывающие подсказки
        self.create_tooltips()

    def create_input_frame(self, parent):
        """Создание блока ввода параметров"""
        input_frame = ttk.LabelFrame(parent, text=" Параметры ввода ", padding=15)
        input_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Поле поиска марки стали
        ttk.Label(input_frame, text="Марка стали:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.search_var = tk.StringVar()
        self.grade_combo = ttk.Combobox(
            input_frame, 
            textvariable=self.search_var,
            width=30,
            values=self.all_grades
        )
        self.grade_combo.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="ew")

        # Поля температур
        self.create_temperature_fields(input_frame)
        
        # Выбор типа расчета
        self.create_calculation_type(input_frame)
        
        # Динамические поля ввода
        self.create_dynamic_fields(input_frame)
        
        # Кнопки управления
        self.create_control_buttons(input_frame)

    def create_temperature_fields(self, parent):
        """Создание полей для температур"""
        ttk.Label(parent, text="Температура среды (°C):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.temperature_env_entry = ttk.Entry(parent, width=15, validate="key")
        self.temperature_env_entry.config(validatecommand=(self.root.register(self.validate_int), "%P"))
        self.temperature_env_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(parent, text="Температура порезки (°C):").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.temp_cut_entry = ttk.Entry(parent, width=15, validate="key")
        self.temp_cut_entry.config(validatecommand=(self.root.register(self.validate_int), "%P"))
        self.temp_cut_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")

    def create_calculation_type(self, parent):
        """Создание переключателя типа расчета"""
        type_frame = ttk.Frame(parent)
        type_frame.grid(row=2, column=0, columnspan=4, pady=10, sticky="ew")
        
        ttk.Label(type_frame, text="Тип расчета:").pack(side="left", padx=5)
        self.cut_type = tk.StringVar()
        ttk.Radiobutton(type_frame, text="Мера", variable=self.cut_type, 
                       value="мера", command=self.update_input_fields).pack(side="left", padx=10)
        ttk.Radiobutton(type_frame, text="Крата", variable=self.cut_type,
                       value="крата", command=self.update_input_fields).pack(side="left", padx=10)

    def create_dynamic_fields(self, parent):
        """Создание динамических полей ввода"""
        self.measure_frame = ttk.Frame(parent)
        self.krata_frame = ttk.Frame(parent)

        # Поля для меры
        ttk.Label(self.measure_frame, text="Длина проката (мм):").pack(side="left", padx=5)
        self.measure_entry = ttk.Entry(self.measure_frame, width=15, validate="key")
        self.measure_entry.config(validatecommand=(self.root.register(self.validate_int), "%P"))
        self.measure_entry.pack(side="left")

        # Поля для краты
        ttk.Label(self.krata_frame, text="Кратность (мм):").pack(side="left", padx=5)
        self.krata_step_entry = ttk.Entry(self.krata_frame, width=10, validate="key")
        self.krata_step_entry.config(validatecommand=(self.root.register(self.validate_int), "%P"))
        self.krata_step_entry.pack(side="left")
        
        ttk.Label(self.krata_frame, text="Число крат:").pack(side="left", padx=5)
        self.krata_count_entry = ttk.Entry(self.krata_frame, width=10, validate="key")
        self.krata_count_entry.config(validatecommand=(self.root.register(self.validate_int), "%P"))
        self.krata_count_entry.pack(side="left")

    def create_control_buttons(self, parent):
        """Создание кнопок управления"""
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=4, column=0, columnspan=4, pady=15, sticky="e")
        
        # Кнопка расчета
        self.calc_btn = ttk.Button(
            btn_frame,
            image=self.icons['calculate'],
            text=" Рассчитать" if self.icons['calculate'] else "Рассчитать",
            compound='left',
            command=self.calculate
        )
        self.calc_btn.pack(side="left", padx=5)

        # Кнопка сброса
        self.reset_btn = ttk.Button(
            btn_frame,
            image=self.icons['reset'],
            text=" Сброс" if self.icons['reset'] else "Сброс",
            compound='left',
            command=self.reset_all
        )
        self.reset_btn.pack(side="left", padx=5)

    def create_results_frame(self, parent):
        """Создание блока результатов"""
        results_frame = ttk.LabelFrame(parent, text=" Результаты расчетов ", padding=15)
        results_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        columns = ("Марка стали", "Температура среды (°C)", 
                 "Температура порезки (°C)", "Коэффициент α (×10⁻⁶/K)", 
                 "Длина проката (мм)", "Усадка", "Длина порезки")
        
        # Таблица результатов
        self.result_tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=2)
        col_widths = [160, 160, 160, 180, 140, 100, 140]
        for col, width in zip(columns, col_widths):
            self.result_tree.heading(col, text=col, anchor="center")
            self.result_tree.column(col, width=width, anchor="center")
        self.result_tree.pack(fill="x", pady=5)

        # Метка с итоговым результатом
        self.result_length_label = ttk.Label(
            results_frame, 
            text="", 
            font=("Helvetica", 14, 'bold'),
            foreground="darkred"
        )
        #self.result_length_label.pack(pady=5)
        
        # Добавляем метку для отображения ошибок
        self.result_label = ttk.Label(
            results_frame,
            text="",
            font=("Helvetica", 10),
            foreground="red"
        )
        self.result_length_label.pack(pady=5)
        #self.result_label.pack(pady=5)

    def create_session_frame(self, parent):
        """Создание блока сессии"""
        session_frame = ttk.LabelFrame(parent, text=" Текущая сессия ", padding=15)
        session_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)

        # Таблица сессии
        columns = ("Марка стали", "Температура среды (°C)", 
                 "Температура порезки (°C)", "Коэффициент α (×10⁻⁶/K)", 
                 "Длина проката (мм)", "Усадка", "Длина порезки")
        self.session_tree = ttk.Treeview(session_frame, columns=columns, show="headings", height=5)
        col_widths = [160, 160, 160, 180, 140, 100, 140]
        for col, width in zip(columns, col_widths):
            self.session_tree.heading(col, text=col, anchor="center")
            self.session_tree.column(col, width=width, anchor="center")
        self.session_tree.pack(fill="both", expand=True)

        # Панель управления сессией
        self.create_session_buttons(session_frame)

    def create_session_buttons(self, parent):
        """Создание кнопок управления сессией"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(pady=10)
        
        # Кнопка добавления
        self.add_btn = ttk.Button(
            btn_frame,
            image=self.icons['add'],
            text=" Добавить" if self.icons['add'] else "Добавить",
            compound='left',
            command=self.add_to_session
        )
        self.add_btn.pack(side="left", padx=5)

        # Кнопка экспорта
        self.export_btn = ttk.Button(
            btn_frame,
            image=self.icons['export'],
            text=" Экспорт" if self.icons['export'] else "Экспорт",
            compound='left',
            command=self.export_to_excel
        )
        self.export_btn.pack(side="left", padx=5)

        # Кнопка очистки
        self.clear_btn = ttk.Button(
            btn_frame,
            image=self.icons['clear'],
            text=" Очистить" if self.icons['clear'] else "Очистить",
            compound='left',
            command=self.clear_session
        )
        self.clear_btn.pack(side="left", padx=5)

    def create_tooltips(self):
        """Добавление всплывающих подсказок"""
        tooltips = {
            self.calc_btn: "Рассчитать длину порезки",
            self.reset_btn: "Сбросить все параметры",
            self.add_btn: "Добавить в текущую сессию",
            self.export_btn: "Экспорт в Excel",
            self.clear_btn: "Очистить сессию"
        }
        
        for widget, text in tooltips.items():
            self.create_tooltip(widget, text)

    def create_tooltip(self, widget, text):
        """Создание всплывающей подсказки для виджета"""
        def enter(event):
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + 20
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = ttk.Label(self.tooltip, text=text, background="#ffffe0", 
                            relief="solid", borderwidth=1, padding=5)
            label.pack()
            
        def leave(event):
            if hasattr(self, 'tooltip') and self.tooltip:
                self.tooltip.destroy()
                
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)


    def validate_int(self, value):
        """Валидация целых чисел"""
        return value == "" or value.isdigit()

    def update_input_fields(self):
        """Обновление полей ввода"""
        self.measure_frame.grid_forget()
        self.krata_frame.grid_forget()
        
        if self.cut_type.get() == "мера":
            self.measure_frame.grid(row=4, column=0, columnspan=2, pady=5, sticky="w")
        elif self.cut_type.get() == "крата":
            self.krata_frame.grid(row=4, column=0, columnspan=2, pady=5, sticky="w")

    def configure_bindings(self):
        """Настройка обработчиков событий"""
        self.grade_combo.bind("<KeyRelease>", self.on_search)
        self.grade_combo.bind("<Return>", lambda e: self.calculate())

    def on_search(self, event):
        """Динамический поиск"""
        search_term = self.search_var.get().lower()
        filtered = [g for g in self.all_grades if search_term in g.lower()]
        self.grade_combo["values"] = filtered

    def add_to_session(self):
        """Добавление текущего результата в сессию"""
        try:
            items = self.result_tree.get_children()
            if not items:
                raise ValueError("Нет данных для сохранения")
            
            item_values = self.result_tree.item(items[0], 'values')
            self.session_data.append(item_values)
            self.session_tree.insert("", "end", values=item_values)
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def clear_session(self):
        """Очистка текущей сессии"""
        if messagebox.askyesno("Подтверждение", "Вы точно хотите очистить текущую сессию?"):
            self.session_data.clear()
            self.session_tree.delete(*self.session_tree.get_children())

    def reset_all(self):
        """Сброс всех полей ввода и результатов"""
        self.search_var.set("")
        self.temperature_env_entry.delete(0, tk.END)
        self.temp_cut_entry.delete(0, tk.END)
        self.measure_entry.delete(0, tk.END)
        self.krata_step_entry.delete(0, tk.END)
        self.krata_count_entry.delete(0, tk.END)
        self.result_tree.delete(*self.result_tree.get_children())
        self.result_length_label.config(text="")
        self.result_label.config(text="")

    def export_to_excel(self):
        """Экспорт данных сессии в Excel"""
        try:
            if not self.session_data:
                messagebox.showerror("Ошибка", "Нет данных для экспорта")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if not file_path:
                return

            df = pd.DataFrame(self.session_data, columns=[
                "Марка стали", "Температура среды (°C)", 
                "Температура порезки (°C)", "Коэффициент α (×10⁻⁶/K)", 
                "Длина проката (мм)", "Усадка", "Длина порезки"
            ])
            
            df.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("Успех", f"Файл сохранен:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка экспорта: {str(e)}")

    def calculate(self):
        """Основная функция расчета"""
        try:
            self.result_tree.delete(*self.result_tree.get_children())
            self.result_label.config(text="")
        
            # Сбор данных
            selected_grade = self.grade_combo.get()
            env_temp = self.temperature_env_entry.get()
            temp_cut = self.temp_cut_entry.get()
            cut_type = self.cut_type.get()

            # Валидация обязательных полей
            if not all([selected_grade, env_temp, temp_cut, cut_type]):
                raise ValueError("Заполните все обязательные поля")

            # Конвертация в целые числа
            env_temp = int(env_temp)
            temp_cut = int(temp_cut)

            # Поиск данных стали
            grade_data = next(
                (g for g in self.data["steel_grades"] if g["steel_grade"] == selected_grade), 
                None
            )
            if not grade_data:
                raise ValueError("Марка стали не найдена")

            # Расчет коэффициента α
            lower_temp, upper_temp = self.find_nearest_temperatures(temp_cut)
            alpha = self.interpolate_alpha(grade_data["alpha"], temp_cut, lower_temp, upper_temp)

            # Расчет длины проката
            if cut_type == "мера":
                measure = self.measure_entry.get()
                if not measure:
                    raise ValueError("Введите длину проката")
                cut_length = int(measure)
                length_info = f"{cut_length} мм"
            elif cut_type == "крата":
                step = self.krata_step_entry.get()
                count = self.krata_count_entry.get()
                if not all([step, count]):
                    raise ValueError("Заполните все поля для краты")
                cut_length = int(step) * int(count)
                length_info = f"{step} × {count}"
            else:
                raise ValueError("Неизвестный тип расчета")

            # Расчет параметров
            delta_temp = temp_cut - env_temp
            shrinkage = alpha * 1e-6 * delta_temp * cut_length
            total_length = cut_length + shrinkage

            # Формирование данных для таблицы
            values = (
                selected_grade,
                f"{env_temp}°C",
                f"{temp_cut}°C",
                f"{alpha:.2f}×10⁻⁶/K",
                length_info,
                f"{shrinkage:.2f} мм",
                f"{total_length:.0f} мм"
            )

            # Добавление в таблицу
            self.result_tree.insert("", "end", values=values)
            self.result_length_label.config(text=f"Расчетная длина порезки: {total_length:.2f} мм")

        except ValueError as ve:
            self.result_label.config(text=str(ve), foreground="red")
            self.result_length_label.config(text="")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
            self.result_length_label.config(text="")

    def find_nearest_temperatures(self, target_temp):
        """Поиск ближайших температур"""
        temps = self.temperatures
        if not temps:
            return None, None
        
        if target_temp <= temps[0]:
            return None, temps[0]
        if target_temp >= temps[-1]:
            return temps[-1], None
        
        left, right = 0, len(temps)-1
        while left <= right:
            mid = (left + right) // 2
            if temps[mid] < target_temp:
                left = mid + 1
            else:
                right = mid - 1
        return temps[right], temps[left]

    def interpolate_alpha(self, alphas, target_temp, t1, t2):
        """Линейная интерполяция коэффициента α"""
        try:
            if not alphas:
                raise ValueError("Отсутствуют данные коэффициентов расширения")
                
            if t2 is None:
                return float(alphas[-1])
            if t1 is None:
                return float(alphas[0])
            
            idx1 = self.temperatures.index(t1)
            idx2 = self.temperatures.index(t2)
            
            if idx1 >= len(alphas) or idx2 >= len(alphas):
                raise ValueError("Несоответствие данных температур и коэффициентов")
            
            return alphas[idx1] + (target_temp - t1) * (alphas[idx2] - alphas[idx1]) / (t2 - t1)
            
        except (ValueError, ZeroDivisionError, IndexError) as e:
            messagebox.showerror("Ошибка данных", f"Ошибка интерполяции: {str(e)}")
            return 0.0

if __name__ == "__main__":
    root = tk.Tk()
    app = SteelAlphaCalculator(root)
    root.mainloop()