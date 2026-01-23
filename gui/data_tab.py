import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from config import Config

# ==================== ВКЛАДКА 1: ЗАВАНТАЖЕННЯ ДАНИХ ====================
class DataTabMixin:
    
    def create_data_tab(self):
        load_frame = ttk.LabelFrame(self.tab1, text="Завантаження Excel файлу", padding=10)
        load_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(load_frame, text="Обрати Excel файл", command=self.load_excel_file).pack(side='left', padx=5)
        self.file_label = ttk.Label(load_frame, text="Файл не обрано")
        self.file_label.pack(side='left', padx=20)
        
        # Панель редагування
        edit_frame = ttk.LabelFrame(self.tab1, text="Редагування таблиці", padding=10)
        edit_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(edit_frame, text="➕ Додати стовпець", command=self.add_column).pack(side='left', padx=5)
        ttk.Button(edit_frame, text="🗑️ Видалити обрані стовпці", command=self.delete_selected_columns).pack(side='left', padx=5)
        ttk.Label(edit_frame, text="| Клік для редагування комірки | Ctrl+клік для виділення кількох стовпців", 
                  font=('Segoe UI', 9, 'italic')).pack(side='left', padx=20)
        
        data_frame = ttk.LabelFrame(self.tab1, text="Перегляд даних", padding=10)
        data_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.tree, _ = self._setup_treeview_with_scrollbars(data_frame, ['Параметр'], height=10)
        self.tree.heading('Параметр', text='Параметр')
        self.tree.column('Параметр', width=300, anchor='w')
        
        # Прив'язка подій для редагування
        self.tree.bind('<Double-1>', self.on_click)
        self.tree.bind('<FocusOut>', self.on_focus_out)
        self.tree.bind('<Return>', self.save_edit)
        self.tree.bind('<Escape>', self.cancel_edit)
        
        self.stats_label = ttk.Label(data_frame, text="Статистика: дані не завантажено", font=('Segoe UI', 9, 'italic'))
        self.stats_label.pack(side='bottom', anchor='w', pady=2)
        
        self._initialize_empty_table()

    # ==== МЕТОДИ ДЛЯ РЕДАГУВАННЯ ===
    
    # Створює порожню таблицю з параметрами
    def _initialize_empty_table(self):
        if self.excel_data is None:
            data = {0: Config.PARAMETER_NAMES[:Config.MIN_ROWS]}
            self.excel_data = pd.DataFrame(data)
            self.num_rows = Config.MIN_ROWS
            self.num_columns = 1
            self.data_columns = 0
            self._refresh_table_display()

    # Додає новий стовпець до таблиці
    def add_column(self):
        if self.excel_data is None:
            self._initialize_empty_table()
        
        new_col_idx = self.excel_data.shape[1]
        self.excel_data[new_col_idx] = ['' for _ in range(self.num_rows)]
        
        self.num_columns = self.excel_data.shape[1]
        self.data_columns = self.num_columns - 1
        
        self._refresh_table_display()
        self._update_parameters_after_change()
        self.stats_label.config(
            text=f"Статистика: Рядків: {self.num_rows}, Стовпців: {self.num_columns} (з них з даними: {self.data_columns})"
        )        
        print(f"✅ Додано стовпець {self.data_columns}")
    
    # Видаляє обрані стовпці (підтримка множинного вибору)
    def delete_selected_columns(self):
        if self.excel_data is None or self.num_columns <= 1:
            messagebox.showwarning("Попередження", "Неможливо видалити стовпець з назвами параметрів!")
            return
                
        # Перевіряємо які колонки виділені (крім першої)
        for col_id in self.tree['columns']:
            if col_id == 'Параметр':
                continue
            # Отримуємо тег виділення через selection
            try:
                bbox = self.tree.bbox('', col_id)
            except:
                pass
        
        # Альтернативний підхід - діалог з можливістю вибору кількох стовпців
        columns = [f"Стовпець {i}" for i in range(1, self.num_columns)]
        if not columns:
            messagebox.showinfo("Інформація", "Немає стовпців для видалення")
            return
        
        # Створюємо діалог з Listbox для множинного вибору
        dialog = tk.Toplevel(self.root)
        dialog.title("Видалення стовпців")
        dialog.geometry("350x400")

        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Оберіть стовпці для видалення (Ctrl+клік для кількох):", 
                  font=('Segoe UI', 10, 'bold')).pack(pady=10)
        
        # Frame для Listbox зі скролом
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical')
        listbox = tk.Listbox(list_frame, selectmode='multiple', yscrollcommand=scrollbar.set, 
                             font=('Segoe UI', 10), height=12)
        scrollbar.config(command=listbox.yview)
        
        for col in columns:
            listbox.insert('end', col)
        
        listbox.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        def confirm_delete():
            selected_indices = listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("Попередження", "Оберіть хоча б один стовпець!")
                return
            
            selected_cols = [columns[i] for i in selected_indices]
            col_numbers = [int(col.split()[-1]) for col in selected_cols]
            
            if messagebox.askyesno("Підтвердження", 
                                  f"Видалити {len(selected_cols)} стовпців?\n" + 
                                  ", ".join(selected_cols)):
                # Видаляємо стовпці (від більшого до меншого щоб не збити індекси)
                for col_idx in sorted(col_numbers, reverse=True):
                    if col_idx < len(self.excel_data.columns):
                        self.excel_data = self.excel_data.drop(columns=[col_idx])
                
                # Перенумеровуємо стовпці
                self.excel_data.columns = range(len(self.excel_data.columns))
                
                self.num_columns = self.excel_data.shape[1]
                self.data_columns = self.num_columns - 1
                
                self._refresh_table_display()
                self._update_parameters_after_change()  # ОНОВЛЮЄМО ПАРАМЕТРИ!
                
                self.stats_label.config(
                    text=f"Статистика: Рядків: {self.num_rows}, Стовпців: {self.num_columns} (з них з даними: {self.data_columns})"
                )
                
                print(f"🗑️ Видалено {len(selected_cols)} стовпців: {', '.join(selected_cols)}")
                dialog.destroy()
        
        # Кнопки
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="✓ Видалити", command=confirm_delete, width=15).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="✗ Скасувати", command=dialog.destroy, width=15).pack(side='left', padx=5)
        
        # Підказка
        ttk.Label(dialog, text="💡 Натисніть Ctrl і клацайте на стовпці для вибору кількох", 
                  font=('Segoe UI', 8, 'italic'), foreground='gray').pack(pady=5)
    
    # Обробка кліку для початку редагування
    def on_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        
        if not row_id or not column:
            return
        
        col_idx = int(column.replace('#', '')) - 1
        
        # Не дозволяємо редагувати перший стовпець
        if col_idx == 0:
            return
        
        row_idx = self.tree.index(row_id)
        
        # Зберігаємо поточне редагування
        if self.editing_item:
            self.save_edit()
        
        # спочатку отримуємо bbox
        bbox = self.tree.bbox(row_id, column)
        if not bbox:
            print(f"❌ Не вдалося отримати bbox для {row_id}, {column}")
            return
        
        print(f"📍 bbox: {bbox}")  # Для діагностики
        
        # Отримуємо поточне значення
        current_value = self.tree.item(row_id)['values'][col_idx]
        
        self.editing_item = row_id
        self.editing_column = col_idx
        self.editing_row = row_idx
        
        # використовуємо Frame як батьківський віджет і коректно позиціонуємо Entry
        self.entry_edit = ttk.Entry(self.tree, font=('Segoe UI', 9))
        
        # додаємо відступи і підняття віджета
        x, y, width, height = bbox
        self.entry_edit.place(
            x=x + 1,           # Невеликий відступ
            y=y + 1,
            width=width - 2,   # Трохи вужче щоб вміщалось
            height=height - 2
        )
        
        # Піднімаємо Entry на передній план
        self.entry_edit.lift()
        
        # Вставляємо значення
        value_str = str(current_value).replace(',', '.') if current_value else ''
        self.entry_edit.insert(0, value_str)
        self.entry_edit.select_range(0, 'end')
        self.entry_edit.focus_set()
        
        # Прив'язуємо події
        self.entry_edit.bind('<Return>', self.save_edit)
        self.entry_edit.bind('<Escape>', self.cancel_edit)
        self.entry_edit.bind('<FocusOut>', self.on_focus_out)
    
    # Зберігає відредаговане значення
    def save_edit(self, event=None):
        if self._is_validating:
            return
        
        if not self.editing_item or not hasattr(self, 'entry_edit') or self.entry_edit is None:
            return
        
        new_value = self.entry_edit.get().strip()
        
        # Валідація: перевіряємо чи це число
        if new_value:
            try:
                clean_value = new_value.replace(',', '.')
                float_value = float(clean_value)
                
                # 🔹 ПЕРЕНОСИМО ПРАПОРЕЦЬ СЮДИ
                self._is_validating = True 
                
                # 🔹 ПЕРЕВІРКА ГРАНИЦ ПАРАМЕТРА
                if not self._validate_parameter_bounds(self.editing_row, float_value):
                    # ВІДВ'ЯЗУЄМО подію, щоб при закритті messagebox не було повторного виклику
                    if hasattr(self, 'entry_edit') and self.entry_edit:
                        self.entry_edit.unbind('<FocusOut>')
                    
                    temp_row = self.editing_row
                    temp_col = self.editing_column
                    temp_item = self.editing_item
                    
                    self.cancel_edit()
                    
                    # Скидаємо прапорець лише ПІСЛЯ того, як вікно помилки закрите (через паузу)
                    self.root.after(200, lambda: setattr(self, '_is_validating', False))
                    self.root.after(300, lambda: self._reopen_edit(temp_item, temp_col, temp_row))
                    return
                
                # Якщо валідація пройшла успішно — знімаємо блок
                self._is_validating = False
                
                # Оновлюємо DataFrame
                self.excel_data.iloc[self.editing_row, self.editing_column] = clean_value
                
                # Оновлюємо відображення в таблиці
                values = list(self.tree.item(self.editing_item)['values'])
                values[self.editing_column] = new_value.replace('.', ',')
                self.tree.item(self.editing_item, values=values)
                
                print(f"Оновлено комірку [{self.editing_row}, {self.editing_column}] = {new_value}")
                
                # Оновлюємо параметри для симуляції
                self._update_parameters_after_change()
                
            except ValueError:
                # ВСТАНОВЛЮЄМО ПРАПОРЕЦЬ ПЕРЕД ПОКАЗОМ ПОМИЛКИ
                self._is_validating = True
                
                # ЗНИЩУЄМО Entry ПЕРЕД показом діалогу!
                temp_row = self.editing_row
                temp_col = self.editing_column
                temp_item = self.editing_item
                self.cancel_edit()  # Це знищить Entry і скине editing_item

                messagebox.showerror("Помилка валідації", "Значення має бути числом!\n\nПриклади:\n• 1.5\n• 1,5\n• 0.001\n• 123")
                
                # СКИДАЄМО ПРАПОРЕЦЬ ПІСЛЯ ЗАКРИТТЯ ДІАЛОГУ
                self._is_validating = False
                
                self.root.after(100, lambda: self._reopen_edit(temp_item, temp_col, temp_row))
                return
        else:
            # Порожнє значення - це OK
            self.excel_data.iloc[self.editing_row, self.editing_column] = ''
            values = list(self.tree.item(self.editing_item)['values'])
            values[self.editing_column] = ''
            self.tree.item(self.editing_item, values=values)
            self._update_parameters_after_change()
        
        self.cancel_edit()
    
    # МЕТОД ДЛЯ ПРОВЕРКИ ГРАНИЦ
    def _validate_parameter_bounds(self, row_idx, value):
        """
        Перевіряє чи значення входить в допустимі межі для параметра
        Повертає True якщо OK, False якщо виходить за межі
        """
        bounds = {
            0: (0.001, 2, "λ1 (інтенсівність потоку подій МРЧ)"),
            1: (0.001, 1, "s1 (тривалість первинної обробки)"),
            2: (0.001, 1, "s2 (тривалість вторинної обробки подій МРЧ)"),
            3: (0.001, 10000, "N (кількість подій для моделювання)"),
            4: (0.001, 5, "Deadline 1"),
            5: (0.001, 2, "λ2 (частота подій жорсткого)"),
            6: (0.001, 2, "s2b (тривалість вторинної обробки подій жорсткого РЧ)"),
            7: (0.001, 5, "Deadline 2")
        }
        
        if row_idx in bounds:
            min_val, max_val, param_name = bounds[row_idx]
            if not (min_val <= value <= max_val):
                messagebox.showerror(
                    "Помилка валідації",
                    f"{param_name}\n\n"
                    f"Допустимий діапазон: від {min_val} до {max_val}\n"
                    f"Введене значення: {value}"
                )
                return False
        
        return True

    # Скасовує редагування
    def cancel_edit(self, event=None):
        if hasattr(self, 'entry_edit') and self.entry_edit is not None:
            try:
                self.entry_edit.destroy()
            except:
                pass
            self.entry_edit = None
        
        self.editing_item = None
        self.editing_column = None
        self.editing_row = None
    
    # Обробка втрати фокусу
    def on_focus_out(self, event=None):
        if self._is_validating:
            return
        
        if hasattr(self, 'entry_edit') and self.entry_edit and self.editing_item:
            self.save_edit()
    
    # Оновлює параметри для симуляції після зміни даних
    def _update_parameters_after_change(self):
        self.convert_excel_to_parameters()
        
        # Оновлюємо s2_values та d1_values
        if self.excel_data is not None and self.num_columns > 1:
            self.s2_values = self._get_grouped_values(
                self.excel_data.iloc[2, 1:self.num_columns].values)
            self.d1_values = self._get_grouped_values(
                self.excel_data.iloc[4, 1:self.num_columns].values)
            
            # Оновлюємо варіанти графіків
            if hasattr(self, 'plot_combobox'):
                self.update_plot_options_based_on_s2_and_d1()
    
    # Оновлює відображення таблиці
    def _refresh_table_display(self):
        if self.excel_data is None:
            return
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.update_tree_columns(self.data_columns)
        
        for i in range(min(self.num_rows, len(Config.PARAMETER_NAMES))):
            param_name = Config.PARAMETER_NAMES[i]
            values = [param_name] + list(self.excel_data.iloc[i, 1:self.num_columns].values)
            
            str_values = [str(v).replace('.', ',') if pd.notna(v) and v != '' else '' for v in values]
            self.tree.insert('', 'end', values=str_values)

    # ==================== МЕТОД ЗАВАНТАЖЕННЯ EXCEL ФАЙЛУ ====================
    def load_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Оберіть Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            self._validation_error_shown = False
            self._clear_previous_results()
            self.excel_data = pd.read_excel(file_path, header=None)
            self.current_file_path = file_path
            
            self.num_rows = self.excel_data.shape[0]
            self.num_columns = self.excel_data.shape[1]
            self.data_columns = self.num_columns - 1
            
            if not self._validate_excel_structure():
                if not hasattr(self, '_validation_error_shown'):
                    messagebox.showerror("Помилка", 
                        f"Файл має неправильний формат!\n\n"
                        f"Вимоги:\n"
                        f"• Мінімум {Config.MIN_ROWS} рядків\n"
                        f"• Мінімум {Config.MIN_COLS} стовпці\n"
                        f"• Всі дані (окрім 1-го стовпця) мають бути числами")
                self.excel_data = None
                return
            
            self._process_excel_data()
            self.file_label.config(text=f"Файл: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося завантажити файл: {str(e)}")
            self.excel_data = None

    # ОЧИЩАЄ ПОПЕРЕДНІ РЕЗУЛЬТАТИ ПРИ ЗАВАНТАЖЕННІ НОВОГО ФАЙЛУ
    def _clear_previous_results(self):
        print("\n🧹 Очищення попередніх результатів...")
                
        if hasattr(self, 'results_tree'):
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            self.results_tree['columns'] = ['Метрика']
            self.results_tree.heading('Метрика', text='Метрика')
            self.results_tree.column('Метрика', width=300, anchor='w')
                
        if hasattr(self, 'plot_frame'):
            for widget in self.plot_frame.winfo_children():
                widget.destroy()
                
        self.simulation_results = self.current_plot_frame = None
                
        if hasattr(self, 'status_label'):
            self.status_label.config(text="Статус: Очікування запуску")
        
        if hasattr(self, 'progress'):
            self.progress['value'] = 0
                
        if hasattr(self, 'plot_var'):
            self.plot_var.set("")
        
        if hasattr(self, 'plot_combobox'):
            self.plot_combobox['values'] = []
                
        if hasattr(self, 'plot_description'):
            self.plot_description.config(state='normal')
            self.plot_description.delete('1.0', 'end')
            self.plot_description.insert('1.0', "Завантажте дані та виконайте симуляцію")
            self.plot_description.config(state='disabled')
                                         
    # МЕТОД С ПРОВЕРКОЙ ГРАНИЦ ПРИ ЗАГРУЗКЕ
    def _validate_excel_structure(self):
        if not (self.excel_data.shape[0] >= Config.MIN_ROWS and 
                self.excel_data.shape[1] >= Config.MIN_COLS and 
                self.data_columns >= 1):
            return False
                
        validation_errors, empty_columns = [], []
        
        for col in range(1, self.excel_data.shape[1]):
            col_name = f"Стовпець {col}"
            has_any_data = False
            
            for row in range(Config.MIN_ROWS):
                cell_value = self.excel_data.iloc[row, col]          
                if pd.isna(cell_value) or str(cell_value).strip() == "":
                    continue
                
                has_any_data = True                
                try:
                    value_str = str(cell_value).strip().replace(',', '.')
                    float_value = float(value_str)
                    
                    # 🔹 ПРОВЕРКА ГРАНИЦ ПРИ ЗАГРУЗКЕ ФАЙЛА
                    bounds = {
                        0: (0.001, 2, "λ1"),
                        1: (0.001, 1, "s1"),
                        2: (0.001, 1, "s2"),
                        3: (0.001, 10000, "N"),
                        4: (0.001, 5, "Deadline 1"),
                        5: (0.001, 2, "λ2"),
                        6: (0.001, 2, "s2b"),
                        7: (0.001, 5, "Deadline 2")
                    }

                    if row in bounds:
                        min_val, max_val, param_name = bounds[row]
                        if not (min_val <= float_value <= max_val):
                            validation_errors.append(
                                f"{col_name}, Рядок {row + 1} ({param_name}): '{cell_value}' "
                                f"виходить за межі [{min_val}..{max_val}]"
                            )
                    
                except (ValueError, TypeError):
                    validation_errors.append(f"{col_name}, Рядок {row + 1}: '{cell_value}' - не є числом")
                        
            if not has_any_data:
                empty_columns.append(col_name)
        
        if validation_errors:
            self._validation_error_shown = True
            error_message = "Знайдено некоректні дані в Excel файлі:\n\n"
            
            for error in validation_errors[:10]:
                error_message += f"• {error}\n"
            
            if len(validation_errors) > 10:
                error_message += f"\n... та ще {len(validation_errors) - 10} помилок"
            
            error_message += "\n\nВимоги:"
            error_message += "\n• λ1: 0..2"
            error_message += "\n• s1: 0..1"
            error_message += "\n• s2: 0..1"
            error_message += "\n• N: 0..10000"
            error_message += "\n• Deadline 1: 0..2"
            error_message += "\n• λ2: 0..2"
            error_message += "\n• s2b: 0..2"
            error_message += "\n• Deadline 2: 0..2"
            
            messagebox.showerror("Помилка валідації даних", error_message)
            
            print("\n" + "="*60)
            print("ПОМИЛКИ ВАЛІДАЦІЇ EXCEL ФАЙЛУ")
            print("="*60)
            for error in validation_errors:
                print(f"  ✗ {error}")
            if empty_columns:
                print(f"\nПорожні стовпці (це нормально): {', '.join(empty_columns)}")
            print("="*60 + "\n")
            
            return False
        
        if empty_columns:
            print(f"ℹ Знайдено {len(empty_columns)} порожніх стовпців: {', '.join(empty_columns)}")
        
        return True

    # Обробка та відображення даних з Excel
    def _process_excel_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.update_tree_columns(self.data_columns)
        
        for i in range(min(Config.MIN_ROWS, len(Config.PARAMETER_NAMES))):
            param_name = Config.PARAMETER_NAMES[i]
            values = [param_name] + list(self.excel_data.iloc[i, 1:self.num_columns].values)
            
            if i == 2:
                self.s2_values = self._get_grouped_values(self.excel_data.iloc[i, 1:self.num_columns].values)
            elif i == 4:
                self.d1_values = self._get_grouped_values(self.excel_data.iloc[i, 1:self.num_columns].values)
            
            str_values = [str(v).replace('.', ',') if pd.notna(v) else '' for v in values]
            self.tree.insert('', 'end', values=str_values)
        
        empty_count = sum(1 for x in values if pd.isna(x) or x == "")
        self.data_columns -= empty_count
        
        self.stats_label.config(text=f"Статистика: Рядків: {self.num_rows}, Стовпців: {self.num_columns} "
                                      f"(з них з даними: {self.data_columns}), Порожніх: {empty_count}")
        
        self.convert_excel_to_parameters()
        self.update_plot_options_based_on_s2_and_d1()

    # Групує значення в рядку по непорожніх групах
    def _get_grouped_values(self, row_data):
        all_groups, current_group = [], []
        
        for val in row_data:
            if pd.notna(val) and str(val).strip() != "":
                current_group.append(val)
            else:
                if current_group:
                    all_groups.append(current_group)
                    current_group = []
        
        if current_group:
            all_groups.append(current_group)
        
        return all_groups

    # Оновлює колонки дерева відповідно до кількості стовпців даних
    def update_tree_columns(self, num_data_columns):
        new_columns = ['Параметр'] + [f'Стовпець {i+1}' for i in range(num_data_columns)]
        self.tree['columns'] = new_columns
        
        for col in self.tree['columns']:
            self.tree.heading(col, text='')
            self.tree.column(col, width=0)
        
        for i, col in enumerate(new_columns):
            self.tree.heading(col, text=col)
            if i == 0:
                self.tree.column(col, width=300, anchor='w', minwidth=320)
            else:
                self.tree.column(col, width=100, anchor='center', minwidth=80)
        
        for item in self.tree.get_children():
            self.tree.delete(item)

    # Конвертує дані Excel у набір параметрів для симуляції
    def convert_excel_to_parameters(self):
        if self.excel_data is None:
            return []
        
        parameter_sets = []
        
        for col in range(1, self.excel_data.shape[1]):
            params = self._extract_column_params(col)
            parameter_sets.append(params)
        
        self.excel_parameters = parameter_sets
        self._log_parameter_stats(parameter_sets)
        return parameter_sets

    # Витягує параметри з одного стовпця
    def _extract_column_params(self, col):
        if not self._column_has_data(col):
            return None
        
        params = {}
        try:
            for row, param_name in enumerate(Config.PARAM_NAMES):
                value_str = str(self.excel_data.iloc[row, col]).strip()
                
                if param_name == 'N':
                    params[param_name] = int(float(value_str.replace(',', '.'))) if value_str else 0
                else:
                    params[param_name] = self._convert_to_float(value_str)
            
            if all(v == 0 or v == 0.0 for v in params.values()):
                return None
            
            return params
            
        except Exception as e:
            print(f"Помилка стовпця {col}: {e}")
            return None

    # Перевіряє чи стовпець має хоча б одне непорожнє значення
    def _column_has_data(self, col):
        for row in range(Config.MIN_ROWS):
            cell_value = self.excel_data.iloc[row, col]
            if pd.notna(cell_value) and str(cell_value).strip() != "":
                return True
        return False

    # Конвертує рядок у float
    def _convert_to_float(self, value_str):
        try:
            clean_str = str(value_str).strip().replace(',', '.')
            return float(clean_str) if clean_str else 0.0
        except (ValueError, TypeError):
            return 0.0

    # Логує статистику по наборам параметрів
    def _log_parameter_stats(self, parameter_sets):
        total = len(parameter_sets)
        valid = sum(1 for p in parameter_sets if p is not None)
        empty = sum(1 for p in parameter_sets if p is None)
        
        print(f"Загальна кількість наборів: {total}")
        print(f"Валідних: {valid}, Порожніх: {empty}")
