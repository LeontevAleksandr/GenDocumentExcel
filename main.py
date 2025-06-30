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


class DocumentGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Генератор документов Excel")
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

        self.load_settings()
        self.create_widgets()

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
        ttk.Label(frame, text="[ДАТА], [КОЛИЧЕСТВО], [ЦЕНА], [СТОИМОСТЬ], [СТОИМОСТЬ_ПРОПИСЬЮ]",
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

    def get_output_path(self, contractor, doc_type, month, year):
        """Возвращает путь для сохранения документа"""
        # Названия месяцев
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

        filename = f"{doc_type} {date_str}.xls"

        dir_path = os.path.join(self.base_dir, contractor, folder_name)

        # Создаем директорию если её нет
        os.makedirs(dir_path, exist_ok=True)

        return os.path.join(dir_path, filename)

    def create_document(self, contractor, doc_type, month, year, quantity, price, total):
        """Создает один документ"""
        try:
            # Получаем пути
            template_path = self.get_template_path(contractor, doc_type)
            output_path = self.get_output_path(contractor, doc_type, month, year)

            # Проверяем наличие шаблона
            if not os.path.exists(template_path):
                raise Exception(f"Шаблон не найден: {template_path}")

            # Открываем шаблон для чтения
            rb = xlrd.open_workbook(template_path, formatting_info=True)
            wb = copy(rb)
            sheet = wb.get_sheet(0)

            # Заполняем данные
            last_day = monthrange(year, month)[1]
            date_str = f"{last_day:02d}.{month:02d}.{year}"

            # Форматируем суммы
            formatted_quantity = self.format_amount(quantity)
            formatted_price = self.format_amount(price)
            formatted_total = self.format_amount(total)

            # Получаем расшифровку суммы прописью
            total_in_words = self.number_to_words(total)

            # Читаем оригинальный файл для поиска меток
            original_sheet = rb.sheet_by_index(0)

            # Ищем и заменяем ячейки с определенными значениями
            for row_idx in range(original_sheet.nrows):
                for col_idx in range(original_sheet.ncols):
                    try:
                        cell_value = original_sheet.cell_value(row_idx, col_idx)
                        if isinstance(cell_value, str) and cell_value:
                            new_value = cell_value
                            # Замена даты
                            if "[ДАТА]" in new_value:
                                new_value = new_value.replace("[ДАТА]", date_str)
                            # Замена количества
                            if "[КОЛИЧЕСТВО]" in new_value:
                                new_value = new_value.replace("[КОЛИЧЕСТВО]", formatted_quantity)
                            # Замена цены
                            if "[ЦЕНА]" in new_value:
                                new_value = new_value.replace("[ЦЕНА]", formatted_price)
                            # Замена общей стоимости
                            if "[СТОИМОСТЬ]" in new_value:
                                new_value = new_value.replace("[СТОИМОСТЬ]", formatted_total)
                            # Замена расшифровки суммы
                            if "[СТОИМОСТЬ_ПРОПИСЬЮ]" in new_value:
                                new_value = new_value.replace("[СТОИМОСТЬ_ПРОПИСЬЮ]", total_in_words)

                            # Если значение изменилось, записываем его
                            if new_value != cell_value:
                                sheet.write(row_idx, col_idx, new_value)
                    except Exception as e:
                        # Игнорируем ошибки отдельных ячеек
                        continue

            # Сохраняем файл
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

                for doc_type in selected_doc_types:
                    success, result = self.create_document(contractor, doc_type, month, year, quantity, price, total)

                    if success:
                        self.log(f"✓ Создан: {contractor} - {doc_type}")
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

    def run(self):
        """Запуск приложения"""
        self.root.mainloop()


if __name__ == "__main__":
    app = DocumentGenerator()
    app.run()