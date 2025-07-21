import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import shutil
from datetime import datetime, date
from calendar import monthrange
import xlrd
import xlwt
from xlutils.copy import copy
import json
import win32print
import win32api
import os
import re
from tkinter import messagebox


class DocumentGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Генератор пака документов Excel v 0.0.2")
        self.root.geometry("1000x700")

        # Основная директория
        self.base_dir = r"D:\ИП Леонтьев говновозка"

        # Типы документов
        self.doc_types = ["Акт выполненных работ", "Счет", "Счет-фактура"]

        # Настройки (сохраняются в файл)
        self.settings_file = "settings.json"
        self.contractors = []
        self.contractor_prices = {}  # Сохраненные цены для контрагентов
        self.templates_dir = "templates"

        # Переменные для чекбоксов
        self.contractor_vars = {}
        self.doc_type_vars = {}

        # Переменные для данных в таблице
        self.table_data = {}  # {contractor: {'quantity': var, 'price': var, 'total': var}}

        # Создаем папку для шаблонов если её нет
        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir)

        # Добавить файл для хранения номеров документов
        self.numbers_file = "document_numbers.json"
        self.document_numbers = {}  # {doc_type: current_number}

        self.load_document_numbers()
        self.load_settings()
        self.create_widgets()

    def load_document_numbers(self):
        """Загружает номера документов из файла"""
        try:
            if os.path.exists(self.numbers_file):
                with open(self.numbers_file, 'r', encoding='utf-8') as f:
                    self.document_numbers = json.load(f)
            else:
                # Инициализация начальными номерами для каждого типа документа
                self.document_numbers = {
                    "Акт выполненных работ": 1,
                    "Счет": 1,
                    "Счет-фактура": 1
                }
                self.save_document_numbers()
        except Exception as e:
            print(f"Ошибка загрузки номеров документов: {e}")
            self.document_numbers = {
                "Акт выполненных работ": 1,
                "Счет": 1,
                "Счет-фактура": 1
            }

    def save_document_numbers(self):
        """Сохраняет номера документов в файл"""
        try:
            with open(self.numbers_file, 'w', encoding='utf-8') as f:
                json.dump(self.document_numbers, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения номеров документов: {e}")

    def get_next_document_number(self, doc_type):
        """Получает следующий номер документа для указанного типа"""
        current_number = self.document_numbers.get(doc_type, 1)
        self.document_numbers[doc_type] = current_number + 1
        self.save_document_numbers()
        return current_number

    def manage_document_numbers(self):
        """Окно управления номерами документов"""
        window = tk.Toplevel(self.root)
        window.title("Управление номерами документов")
        window.geometry("400x250")

        frame = ttk.Frame(window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Текущие номера документов:",
                  font=('TkDefaultFont', 10, 'bold')).pack(pady=(0, 10))

        # Словарь для хранения переменных полей ввода
        number_vars = {}

        for i, (doc_type, current_number) in enumerate(self.document_numbers.items()):
            row_frame = ttk.Frame(frame)
            row_frame.pack(fill=tk.X, pady=5)

            ttk.Label(row_frame, text=f"{doc_type}:").pack(side=tk.LEFT)

            var = tk.StringVar(value=str(current_number))
            number_vars[doc_type] = var

            entry = ttk.Entry(row_frame, textvariable=var, width=10)
            entry.pack(side=tk.RIGHT)

        def save_numbers():
            try:
                for doc_type, var in number_vars.items():
                    self.document_numbers[doc_type] = int(var.get())
                self.save_document_numbers()
                messagebox.showinfo("Успех", "Номера документов обновлены")
                window.destroy()
            except ValueError:
                messagebox.showerror("Ошибка", "Введите корректные числа")

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))

        ttk.Button(button_frame, text="Сохранить", command=save_numbers).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Отмена", command=window.destroy).pack(side=tk.RIGHT)

    def number_to_words(self, number):
        """Преобразует число в текст прописью"""
        try:
            num = float(number)
            rubles = int(num)
            kopecks = int(round((num - rubles) * 100))

            # Массивы для преобразования
            ones = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']
            tens = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят',
                    'девяносто']
            teens = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать',
                     'семнадцать', 'восемнадцать', 'девятнадцать']
            hundreds = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот',
                        'девятьсот']

            def convert_group(n):
                """Преобразует группу из трех цифр"""
                result = []
                h = n // 100
                t = (n % 100) // 10
                o = n % 10

                if h > 0:
                    result.append(hundreds[h])

                if t == 1:
                    result.append(teens[o])
                else:
                    if t > 0:
                        result.append(tens[t])
                    if o > 0:
                        result.append(ones[o])

                return ' '.join(result)

            if rubles == 0:
                rubles_text = "ноль рублей"
            else:
                # Разбиваем на группы по тысячам
                groups = []
                temp = rubles

                # Единицы
                if temp > 0:
                    group = temp % 1000
                    if group > 0:
                        group_text = convert_group(group)
                        # Склонение для рублей
                        if group % 10 == 1 and group % 100 != 11:
                            rubles_word = "рубль"
                        elif group % 10 in [2, 3, 4] and group % 100 not in [12, 13, 14]:
                            rubles_word = "рубля"
                        else:
                            rubles_word = "рублей"
                        groups.append(f"{group_text} {rubles_word}")
                    temp //= 1000

                # Тысячи
                if temp > 0:
                    group = temp % 1000
                    if group > 0:
                        group_text = convert_group(group)
                        # Корректировка для женского рода (тысячи)
                        group_text = group_text.replace('один', 'одна').replace('два', 'две')

                        if group % 10 == 1 and group % 100 != 11:
                            thousands_word = "тысяча"
                        elif group % 10 in [2, 3, 4] and group % 100 not in [12, 13, 14]:
                            thousands_word = "тысячи"
                        else:
                            thousands_word = "тысяч"
                        groups.insert(0, f"{group_text} {thousands_word}")
                    temp //= 1000

                # Миллионы
                if temp > 0:
                    group = temp % 1000
                    if group > 0:
                        group_text = convert_group(group)
                        if group % 10 == 1 and group % 100 != 11:
                            millions_word = "миллион"
                        elif group % 10 in [2, 3, 4] and group % 100 not in [12, 13, 14]:
                            millions_word = "миллиона"
                        else:
                            millions_word = "миллионов"
                        groups.insert(0, f"{group_text} {millions_word}")

                rubles_text = ' '.join(groups)

            # Склонение копеек
            if kopecks % 10 == 1 and kopecks % 100 != 11:
                kopecks_word = "копейка"
            elif kopecks % 10 in [2, 3, 4] and kopecks % 100 not in [12, 13, 14]:
                kopecks_word = "копейки"
            else:
                kopecks_word = "копеек"

            # Первая буква заглавная
            result = f"{rubles_text.capitalize()} {kopecks:02d} {kopecks_word}"
            return result

        except:
            return "ошибка преобразования"

    def format_amount(self, amount):
        """Форматирует сумму без лишних нулей"""
        try:
            num = float(amount)
            if num == int(num):
                return str(int(num))
            else:
                return f"{num:.2f}".rstrip('0').rstrip('.')
        except:
            return amount

    def load_settings(self):
        """Загружает настройки из файла"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.contractors = settings.get('contractors', [])
                    self.contractor_prices = settings.get('contractor_prices', {})
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")
            self.contractors = []
            self.contractor_prices = {}

    def save_settings(self):
        """Сохраняет настройки в файл"""
        try:
            # Обновляем сохраненные цены из таблицы
            for contractor in self.contractors:
                if contractor in self.table_data:
                    price_value = self.table_data[contractor]['price'].get()
                    if price_value:
                        self.contractor_prices[contractor] = price_value

            settings = {
                'contractors': self.contractors,
                'contractor_prices': self.contractor_prices
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: {e}")

    def create_widgets(self):
        """Создает интерфейс приложения"""
        # Главное меню
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Меню настроек
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Настройки", menu=settings_menu)
        settings_menu.add_command(label="Управление контрагентами", command=self.manage_contractors)
        settings_menu.add_command(label="Управление шаблонами", command=self.manage_templates)
        settings_menu.add_command(label="Номера документов", command=self.manage_document_numbers)
        print_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Печать", menu=print_menu)
        print_menu.add_command(label="Печать документов", command=self.open_print_window)

        def add_print_menu_to_widgets(self):
            """В методе create_widgets после settings_menu добавить:"""
            print_menu = tk.Menu(menubar, tearoff=0)
            menubar.add_cascade(label="Печать", menu=print_menu)
            print_menu.add_command(label="Печать документов", command=self.open_print_window)

        # Основной фрейм с прокруткой
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        main_frame = ttk.Frame(scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Верхняя панель с месяцем и кнопками
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 20))

        # Выбор месяца и года
        date_frame = ttk.LabelFrame(top_frame, text="Период", padding="5")
        date_frame.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(date_frame, text="Месяц:").grid(row=0, column=0, padx=5)
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        months = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
        ttk.Combobox(date_frame, textvariable=self.month_var, values=months,
                     state="readonly", width=8).grid(row=0, column=1, padx=5)

        ttk.Label(date_frame, text="Год:").grid(row=0, column=2, padx=5)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Entry(date_frame, textvariable=self.year_var, width=8).grid(row=0, column=3, padx=5)

        # Кнопки действий
        button_frame = ttk.LabelFrame(top_frame, text="Действия", padding="5")
        button_frame.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Button(button_frame, text="Создать выбранные документы",
                   command=self.create_selected_documents).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Сохранить цены",
                   command=self.save_settings).pack(side=tk.LEFT, padx=5)

        # Быстрые действия
        quick_frame = ttk.LabelFrame(top_frame, text="Быстрые действия", padding="5")
        quick_frame.pack(side=tk.LEFT)

        ttk.Button(quick_frame, text="Выбрать всех",
                   command=self.select_all_contractors).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="Снять выбор",
                   command=self.deselect_all_contractors).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="Все документы",
                   command=self.select_all_docs).pack(side=tk.LEFT, padx=2)

        # Выбор типов документов
        doc_frame = ttk.LabelFrame(main_frame, text="Типы документов для создания", padding="10")
        doc_frame.pack(fill=tk.X, pady=(0, 20))

        doc_inner_frame = ttk.Frame(doc_frame)
        doc_inner_frame.pack()

        for i, doc_type in enumerate(self.doc_types):
            var = tk.BooleanVar(value=True)
            self.doc_type_vars[doc_type] = var
            ttk.Checkbutton(doc_inner_frame, text=doc_type, variable=var).grid(row=0, column=i, padx=20)

        # Основная таблица
        table_frame = ttk.LabelFrame(main_frame, text="Контрагенты и данные", padding="10")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        self.create_table(table_frame)

        # Область для логов
        log_frame = ttk.LabelFrame(main_frame, text="Лог операций", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=8)
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Настройка прокрутки основного окна
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Привязка колесика мыши
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def create_table(self, parent):
        """Создает таблицу с контрагентами"""
        if hasattr(self, 'table_frame'):
            self.table_frame.destroy()

        self.table_frame = ttk.Frame(parent)
        self.table_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовки таблицы
        headers = ["Выбрать", "Контрагент", "Количество (куб.м)", "Цена за единицу", "Общая стоимость"]
        for i, header in enumerate(headers):
            label = ttk.Label(self.table_frame, text=header, font=('TkDefaultFont', 9, 'bold'))
            label.grid(row=0, column=i, padx=5, pady=5, sticky='ew')

        # Строки для каждого контрагента
        for row, contractor in enumerate(self.contractors, 1):
            # Чекбокс выбора
            var = tk.BooleanVar(value=True)
            self.contractor_vars[contractor] = var
            ttk.Checkbutton(self.table_frame, variable=var).grid(row=row, column=0, padx=5, pady=2)

            # Название контрагента
            ttk.Label(self.table_frame, text=contractor).grid(row=row, column=1, padx=5, pady=2, sticky='w')

            # Поля ввода данных
            quantity_var = tk.StringVar()
            price_var = tk.StringVar()
            total_var = tk.StringVar()

            # Загружаем сохраненную цену
            if contractor in self.contractor_prices:
                price_var.set(self.contractor_prices[contractor])

            self.table_data[contractor] = {
                'quantity': quantity_var,
                'price': price_var,
                'total': total_var
            }

            # Поле количества
            quantity_entry = ttk.Entry(self.table_frame, textvariable=quantity_var, width=15)
            quantity_entry.grid(row=row, column=2, padx=5, pady=2)

            # Поле цены
            price_entry = ttk.Entry(self.table_frame, textvariable=price_var, width=15)
            price_entry.grid(row=row, column=3, padx=5, pady=2)

            # Поле общей стоимости (только для чтения)
            total_entry = ttk.Entry(self.table_frame, textvariable=total_var, width=15, state='readonly')
            total_entry.grid(row=row, column=4, padx=5, pady=2)

            # Привязываем автоматический расчет
            def make_calculator(contractor_name):
                def calculate(*args):
                    try:
                        data = self.table_data[contractor_name]
                        quantity = float(data['quantity'].get() or 0)
                        price = float(data['price'].get() or 0)
                        total = quantity * price
                        # Форматируем без лишних нулей
                        data['total'].set(self.format_amount(total))
                    except ValueError:
                        self.table_data[contractor_name]['total'].set("0")

                return calculate

            calculator = make_calculator(contractor)
            quantity_var.trace('w', calculator)
            price_var.trace('w', calculator)

        # Настройка растягивания столбцов
        for i in range(5):
            self.table_frame.columnconfigure(i, weight=1)

    def select_all_contractors(self):
        """Выбирает всех контрагентов"""
        for var in self.contractor_vars.values():
            var.set(True)

    def deselect_all_contractors(self):
        """Снимает выбор со всех контрагентов"""
        for var in self.contractor_vars.values():
            var.set(False)

    def select_all_docs(self):
        """Выбирает все типы документов"""
        for var in self.doc_type_vars.values():
            var.set(True)

    def log(self, message):
        """Добавляет сообщение в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def manage_contractors(self):
        """Окно управления контрагентами"""
        window = tk.Toplevel(self.root)
        window.title("Управление контрагентами")
        window.geometry("400x300")

        frame = ttk.Frame(window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # Список контрагентов
        listbox = tk.Listbox(frame, height=10)
        listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        for contractor in self.contractors:
            listbox.insert(tk.END, contractor)

        # Поле для нового контрагента
        entry_frame = ttk.Frame(frame)
        entry_frame.pack(fill=tk.X, pady=(0, 10))

        new_contractor_var = tk.StringVar()
        ttk.Entry(entry_frame, textvariable=new_contractor_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        def add_contractor():
            name = new_contractor_var.get().strip()
            if name and name not in self.contractors:
                self.contractors.append(name)
                listbox.insert(tk.END, name)
                new_contractor_var.set("")
                self.save_settings()
                self.create_table(self.table_frame.master)  # Пересоздаем таблицу

        def remove_contractor():
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                contractor = self.contractors.pop(index)
                listbox.delete(index)
                # Удаляем из сохраненных цен
                if contractor in self.contractor_prices:
                    del self.contractor_prices[contractor]
                self.save_settings()
                self.create_table(self.table_frame.master)  # Пересоздаем таблицу

        ttk.Button(entry_frame, text="Добавить", command=add_contractor).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(frame, text="Удалить выбранного", command=remove_contractor).pack()

    def manage_templates(self):
        """Окно управления шаблонами"""
        window = tk.Toplevel(self.root)
        window.title("Управление шаблонами")
        window.geometry("700x500")

        frame = ttk.Frame(window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Шаблоны должны находиться в папке 'templates'").pack(pady=(0, 10))
        ttk.Label(frame, text="Имена файлов: {Контрагент}_{Тип документа}.xls").pack(pady=(0, 10))
        ttk.Label(frame, text="В шаблонах используйте метки:").pack(pady=(0, 5))
        ttk.Label(frame, text="[НОМЕР], [ДАТА], [КОЛИЧЕСТВО], [ЦЕНА], [СТОИМОСТЬ], [СТОИМОСТЬ_ПРОПИСЬЮ]",
                  font=('TkDefaultFont', 9, 'bold')).pack(pady=(0, 10))

        # Список существующих шаблонов
        listbox = tk.Listbox(frame, height=15)
        listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        def refresh_templates():
            listbox.delete(0, tk.END)
            if os.path.exists(self.templates_dir):
                for file in os.listdir(self.templates_dir):
                    if file.endswith('.xls'):
                        listbox.insert(tk.END, file)

        refresh_templates()

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="Обновить список", command=refresh_templates).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Открыть папку шаблонов",
                   command=lambda: os.startfile(self.templates_dir)).pack(side=tk.RIGHT)

    def get_template_path(self, contractor, doc_type):
        """Возвращает путь к шаблону"""
        filename = f"{contractor}_{doc_type}.xls"
        return os.path.join(self.templates_dir, filename)

    def get_output_path(self, contractor, doc_type, month, year, doc_number):
        """Возвращает путь для сохранения документа с номером"""
        month_names = {
            1: "январь", 2: "февраль", 3: "март", 4: "апрель",
            5: "май", 6: "июнь", 7: "июль", 8: "август",
            9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"
        }

        month_name = month_names[month]
        folder_name = f"{month_name} {year}"

        # Последний день месяца
        last_day = monthrange(year, month)[1]
        date_str = f"{last_day:02d}.{month:02d}.{year}"

        # Имя файла с номером
        filename = f"{doc_type} № {doc_number} от {date_str}.xls"

        dir_path = os.path.join(self.base_dir, contractor, folder_name)
        os.makedirs(dir_path, exist_ok=True)

        return os.path.join(dir_path, filename)

    def create_document(self, contractor, doc_type, month, year, quantity, price, total, doc_number):
        """Создает один документ с номером"""
        try:
            template_path = self.get_template_path(contractor, doc_type)
            output_path = self.get_output_path(contractor, doc_type, month, year, doc_number)

            if not os.path.exists(template_path):
                raise Exception(f"Шаблон не найден: {template_path}")

            rb = xlrd.open_workbook(template_path, formatting_info=True)
            wb = copy(rb)
            sheet = wb.get_sheet(0)

            last_day = monthrange(year, month)[1]
            date_str = f"{last_day:02d}.{month:02d}.{year}"

            formatted_quantity = self.format_amount(quantity)
            formatted_price = self.format_amount(price)
            formatted_total = self.format_amount(total)
            total_in_words = self.number_to_words(total)

            original_sheet = rb.sheet_by_index(0)

            for row_idx in range(original_sheet.nrows):
                for col_idx in range(original_sheet.ncols):
                    try:
                        cell_value = original_sheet.cell_value(row_idx, col_idx)
                        if isinstance(cell_value, str) and cell_value:
                            new_value = cell_value

                            # Добавляем замену номера документа
                            if "[НОМЕР]" in new_value:
                                new_value = new_value.replace("[НОМЕР]", str(doc_number))
                            if "[ДАТА]" in new_value:
                                new_value = new_value.replace("[ДАТА]", date_str)
                            if "[КОЛИЧЕСТВО]" in new_value:
                                new_value = new_value.replace("[КОЛИЧЕСТВО]", formatted_quantity)
                            if "[ЦЕНА]" in new_value:
                                new_value = new_value.replace("[ЦЕНА]", formatted_price)
                            if "[СТОИМОСТЬ]" in new_value:
                                new_value = new_value.replace("[СТОИМОСТЬ]", formatted_total)
                            if "[СТОИМОСТЬ_ПРОПИСЬЮ]" in new_value:
                                new_value = new_value.replace("[СТОИМОСТЬ_ПРОПИСЬЮ]", total_in_words)

                            if new_value != cell_value:
                                sheet.write(row_idx, col_idx, new_value)
                    except Exception:
                        continue

            wb.save(output_path)
            return True, output_path

        except Exception as e:
            return False, str(e)

    def create_selected_documents(self):
        """Создает выбранные документы для выбранных контрагентов"""
        try:
            month = int(self.month_var.get())
            year = int(self.year_var.get())

            # Получаем выбранных контрагентов
            selected_contractors = [contractor for contractor, var in self.contractor_vars.items() if var.get()]

            # Получаем выбранные типы документов
            selected_doc_types = [doc_type for doc_type, var in self.doc_type_vars.items() if var.get()]

            if not selected_contractors:
                messagebox.showwarning("Предупреждение", "Не выбран ни один контрагент")
                return

            if not selected_doc_types:
                messagebox.showwarning("Предупреждение", "Не выбран ни один тип документа")
                return

            # Проверяем заполненность данных
            contractors_with_data = []
            for contractor in selected_contractors:
                data = self.table_data[contractor]
                quantity = data['quantity'].get().strip()
                price = data['price'].get().strip()

                if not quantity or not price:
                    self.log(f"Пропуск {contractor}: не заполнены количество или цена")
                    continue

                try:
                    float(quantity)
                    float(price)
                    contractors_with_data.append(contractor)
                except ValueError:
                    self.log(f"Пропуск {contractor}: некорректные числовые данные")

            if not contractors_with_data:
                messagebox.showerror("Ошибка", "Нет контрагентов с корректно заполненными данными")
                return

            # Создаем документы
            total_created = 0
            total_errors = 0

            self.log(f"Начинаем создание документов для {len(contractors_with_data)} контрагентов")

            for contractor in contractors_with_data:
                data = self.table_data[contractor]
                quantity = data['quantity'].get()
                price = data['price'].get()
                total = data['total'].get()

                contractor_created = 0

                # Получаем номера документов для каждого типа
                contractor_doc_numbers = {}
                for doc_type in selected_doc_types:
                    contractor_doc_numbers[doc_type] = self.get_next_document_number(doc_type)

                # Создаем документы с полученными номерами
                for doc_type in selected_doc_types:
                    doc_number = contractor_doc_numbers[doc_type]
                    # ИСПРАВЛЕНИЕ: добавляем недостающий аргумент doc_number
                    success, result = self.create_document(contractor, doc_type, month, year, quantity, price, total,
                                                           doc_number)

                    if success:
                        self.log(f"✓ Создан: {contractor} - {doc_type} № {doc_number}")
                        total_created += 1
                        contractor_created += 1
                    else:
                        self.log(f"✗ Ошибка: {contractor} - {doc_type}: {result}")
                        total_errors += 1

                if contractor_created > 0:
                    self.log(f"Для {contractor} создано {contractor_created} документов")

            # Сохраняем настройки (цены)
            self.save_settings()

            # Показываем результат
            if total_created > 0:
                message = f"Успешно создано документов: {total_created}"
                if total_errors > 0:
                    message += f"\nОшибок: {total_errors}"
                messagebox.showinfo("Результат", message)
            else:
                messagebox.showerror("Ошибка", f"Не удалось создать ни одного документа. Ошибок: {total_errors}")

        except Exception as e:
            error_msg = f"Общая ошибка: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Ошибка", error_msg)

    def get_available_printers(self):
        """Получает список доступных принтеров"""
        try:
            printers = [printer[2] for printer in
                        win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            return printers
        except:
            return []

    def get_month_folder_name(self, month, year):
        """Возвращает название папки месяца"""
        month_names = {
            1: "январь", 2: "февраль", 3: "март", 4: "апрель",
            5: "май", 6: "июнь", 7: "июль", 8: "август",
            9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"
        }
        return f"{month_names[month]} {year}"

    def find_documents_for_print(self, contractor, month, year, doc_types):
        """Находит документы для печати"""
        import re

        folder_name = self.get_month_folder_name(month, year)
        contractor_path = os.path.join(self.base_dir, contractor, folder_name)

        if not os.path.exists(contractor_path):
            return []

        found_docs = []

        # Получаем все файлы в папке
        all_files = [f for f in os.listdir(contractor_path) if f.endswith('.xls')]

        for doc_type in doc_types:
            # Создаем регулярное выражение для точного поиска
            # Ищем файлы, которые начинаются с типа документа, за которым следует " № " и номер
            pattern = re.compile(rf"^{re.escape(doc_type)} № \d+")

            for file in all_files:
                if pattern.match(file):
                    found_docs.append({
                        'type': doc_type,
                        'file': file,
                        'path': os.path.join(contractor_path, file)
                    })
                    break  # Найден файл для этого типа документа, переходим к следующему типу

        return found_docs

    def print_document(self, file_path, printer_name, copies=1):
        """Печатает документ Excel"""
        try:
            # Открываем Excel файл и печатаем
            win32api.ShellExecute(0, "print", file_path, f'/p:{printer_name}', ".", 0)
            return True
        except Exception as e:
            return False, str(e)

    def open_print_window(self):
        """Открывает окно печати документов"""
        window = tk.Toplevel(self.root)
        window.title("Печать документов")
        window.geometry("800x900")

        main_frame = ttk.Frame(window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Верхняя панель с настройками
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки печати", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))

        # Первая строка: месяц, год, принтер
        row1 = ttk.Frame(settings_frame)
        row1.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(row1, text="Месяц:").pack(side=tk.LEFT, padx=(0, 5))
        month_var = tk.StringVar(value=str(datetime.now().month))
        months = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
        ttk.Combobox(row1, textvariable=month_var, values=months, state="readonly", width=8).pack(side=tk.LEFT,
                                                                                                  padx=(0, 20))

        ttk.Label(row1, text="Год:").pack(side=tk.LEFT, padx=(0, 5))
        year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Entry(row1, textvariable=year_var, width=8).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(row1, text="Принтер:").pack(side=tk.LEFT, padx=(0, 5))
        printer_var = tk.StringVar()
        printers = self.get_available_printers()
        printer_combo = ttk.Combobox(row1, textvariable=printer_var, values=printers, state="readonly", width=30)
        printer_combo.pack(side=tk.LEFT, padx=(0, 20))
        if printers:
            printer_combo.set(printers[0])

        # Вторая строка: типы документов
        row2 = ttk.Frame(settings_frame)
        row2.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(row2, text="Типы документов:").pack(side=tk.LEFT, padx=(0, 10))
        doc_vars = {}
        for doc_type in self.doc_types:
            var = tk.BooleanVar(value=True)
            doc_vars[doc_type] = var
            ttk.Checkbutton(row2, text=doc_type, variable=var).pack(side=tk.LEFT, padx=(0, 15))

        # Третья строка: количество копий для каждого типа документа
        row3 = ttk.Frame(settings_frame)
        row3.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(row3, text="Количество копий:").pack(side=tk.LEFT, padx=(0, 10))
        copies_vars = {}
        for doc_type in self.doc_types:
            ttk.Label(row3, text=f"{doc_type}:").pack(side=tk.LEFT, padx=(0, 5))
            var = tk.StringVar(value="1")
            copies_vars[doc_type] = var
            ttk.Entry(row3, textvariable=var, width=5).pack(side=tk.LEFT, padx=(0, 15))

        # Четвертая строка: кнопки
        row4 = ttk.Frame(settings_frame)
        row4.pack(fill=tk.X)

        ttk.Button(row4, text="Обновить список",
                   command=lambda: self.update_print_table(table_frame, month_var, year_var, doc_vars)).pack(
            side=tk.LEFT, padx=(0, 10))
        ttk.Button(row4, text="Печать выбранных",
                   command=lambda: self.print_selected_documents(table_frame, printer_var, copies_vars)).pack(
            side=tk.LEFT)
        ttk.Button(row4, text="Выбрать все",
                   command=lambda: self.select_all_print_documents(table_frame)).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(row4, text="Снять выбор",
                   command=lambda: self.deselect_all_print_documents(table_frame)).pack(side=tk.LEFT, padx=(10, 0))

        # Основная таблица с документами
        table_frame = ttk.LabelFrame(main_frame, text="Найденные документы", padding="10")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Создаем таблицу
        self.create_print_table(table_frame, month_var, year_var, doc_vars)

        # Область для логов печати
        log_frame = ttk.LabelFrame(main_frame, text="Лог печати", padding="5")
        log_frame.pack(fill=tk.X)

        print_log = tk.Text(log_frame, height=6)
        print_log.pack(fill=tk.X)

        # Сохраняем ссылки для использования в других методах
        window.print_log = print_log
        window.table_frame = table_frame

    def create_print_table(self, parent, month_var, year_var, doc_vars):
        """Создает таблицу документов для печати"""
        # Очищаем предыдущую таблицу
        for widget in parent.winfo_children():
            widget.destroy()

        # Создаем Treeview для отображения документов
        columns = ('select', 'contractor', 'doc_type', 'file_name')
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=15)

        # Настраиваем заголовки
        tree.heading('select', text='Выбрать')
        tree.heading('contractor', text='Контрагент')
        tree.heading('doc_type', text='Тип документа')
        tree.heading('file_name', text='Файл')

        # Настраиваем ширину столбцов
        tree.column('select', width=80)
        tree.column('contractor', width=200)
        tree.column('doc_type', width=150)
        tree.column('file_name', width=350)

        # Скроллбар
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        parent.tree = tree

        # Заполняем таблицу данными
        self.update_print_table(parent, month_var, year_var, doc_vars)

    def update_print_table(self, table_frame, month_var, year_var, doc_vars):
        """Обновляет таблицу документов для печати"""
        try:
            month = int(month_var.get())
            year = int(year_var.get())
            selected_doc_types = [doc_type for doc_type, var in doc_vars.items() if var.get()]

            tree = table_frame.tree
            tree.delete(*tree.get_children())

            for contractor in self.contractors:
                documents = self.find_documents_for_print(contractor, month, year, selected_doc_types)

                for doc in documents:
                    tree.insert('', 'end', values=(
                        '☐',  # checkbox placeholder
                        contractor,
                        doc['type'],
                        doc['file']
                    ), tags=(contractor, doc['path']))

            # Привязываем обработчик клика для чекбоксов
            def on_click(event):
                item = tree.selection()[0]
                values = list(tree.item(item, 'values'))
                if values[0] == '☐':
                    values[0] = '☑'
                else:
                    values[0] = '☐'
                tree.item(item, values=values)

            tree.bind('<Button-1>', on_click)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить список: {e}")

    def print_selected_documents(self, table_frame, printer_var, copies_vars):
        """Печатает выбранные документы в правильном порядке"""
        try:
            tree = table_frame.tree
            printer_name = printer_var.get()

            if not printer_name:
                messagebox.showwarning("Предупреждение", "Выберите принтер")
                return

            # Собираем все выбранные документы
            selected_items = []
            for item in tree.get_children():
                values = tree.item(item, 'values')
                if values[0] == '☑':  # Если выбран
                    tags = tree.item(item, 'tags')
                    file_path = tags[1] if len(tags) > 1 else None
                    selected_items.append({
                        'contractor': values[1],
                        'doc_type': values[2],
                        'file_name': values[3],
                        'file_path': file_path
                    })

            if not selected_items:
                messagebox.showwarning("Предупреждение", "Не выбран ни один документ")
                return

            # Группируем документы по контрагентам
            contractors_docs = {}
            for item in selected_items:
                contractor = item['contractor']
                if contractor not in contractors_docs:
                    contractors_docs[contractor] = {}

                doc_type = item['doc_type']
                contractors_docs[contractor][doc_type] = item

            # Получаем окно для доступа к логу
            window = table_frame.winfo_toplevel()
            print_log = window.print_log

            total_printed = 0
            total_errors = 0

            # Печатаем в правильном порядке: по контрагентам, потом по типам документов
            for contractor in sorted(contractors_docs.keys()):
                print_log.insert(tk.END, f"\n--- Печать документов для {contractor} ---\n")
                print_log.see(tk.END)
                window.update_idletasks()

                # Печатаем документы в заданном порядке типов
                for doc_type in self.doc_types:
                    if doc_type in contractors_docs[contractor]:
                        item = contractors_docs[contractor][doc_type]
                        copies_count = int(copies_vars[doc_type].get()) if copies_vars[doc_type].get().isdigit() else 1

                        print_log.insert(tk.END, f"  {doc_type} ({copies_count} копий):\n")

                        # Печатаем все копии подряд
                        for copy_num in range(copies_count):
                            try:
                                success = self.print_document(item['file_path'], printer_name)
                                if success:
                                    total_printed += 1
                                    print_log.insert(tk.END, f"    ✓ Копия {copy_num + 1}: {item['file_name']}\n")
                                else:
                                    total_errors += 1
                                    print_log.insert(tk.END,
                                                     f"    ✗ Ошибка копии {copy_num + 1}: {item['file_name']}\n")
                            except Exception as e:
                                total_errors += 1
                                print_log.insert(tk.END, f"    ✗ Ошибка копии {copy_num + 1}: {e}\n")

                            print_log.see(tk.END)
                            window.update_idletasks()

            print_log.insert(tk.END, f"\n=== ИТОГО ===\n")
            print_log.insert(tk.END, f"Отправлено на печать: {total_printed} документов\n")
            if total_errors > 0:
                print_log.insert(tk.END, f"Ошибок: {total_errors}\n")

            message = f"Отправлено на печать: {total_printed} документов"
            if total_errors > 0:
                message += f"\nОшибок: {total_errors}"

            messagebox.showinfo("Результат печати", message)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при печати: {e}")

    def select_all_print_documents(self, table_frame):
        """Выбирает все документы в таблице печати"""
        try:
            tree = table_frame.tree
            for item in tree.get_children():
                values = list(tree.item(item, 'values'))
                values[0] = '☑'
                tree.item(item, values=values)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось выбрать все документы: {e}")

    def deselect_all_print_documents(self, table_frame):
        """Снимает выбор со всех документов в таблице печати"""
        try:
            tree = table_frame.tree
            for item in tree.get_children():
                values = list(tree.item(item, 'values'))
                values[0] = '☐'
                tree.item(item, values=values)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось снять выбор: {e}")

    def run(self):
        """Запуск приложения"""
        self.root.mainloop()


if __name__ == "__main__":
    app = DocumentGenerator()
    app.run()