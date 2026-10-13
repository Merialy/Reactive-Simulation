import os
import pandas as pd
import tkinter as tk
import matplotlib.pyplot as plt
from tkinter import ttk, messagebox
from config import Config
from datetime import datetime
from openpyxl.utils import get_column_letter

# ==================== ВКЛАДКА 4: ЕКСПОРТ ====================
class ExportMixin:
    
    def create_export_tab(self):
        main_frame = ttk.Frame(self.tab4)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        control_frame = ttk.LabelFrame(main_frame, text="Опції експорту", padding=10)
        control_frame.pack(fill='x', pady=(0, 10))
        
        data_frame = ttk.Frame(control_frame)
        data_frame.pack(fill='x', pady=5)
        
        ttk.Label(data_frame, text="Дані для експорту:").pack(side='left', padx=5)
        
        self.export_data_type = tk.StringVar(value="all")
        ttk.Radiobutton(data_frame, text="Вхідні параметри", 
                       variable=self.export_data_type, value="input").pack(side='left', padx=5)
        ttk.Radiobutton(data_frame, text="Результати симуляції", 
                       variable=self.export_data_type, value="results").pack(side='left', padx=5)
        ttk.Radiobutton(data_frame, text="Всі дані", 
                       variable=self.export_data_type, value="all").pack(side='left', padx=5)
        
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill='x', pady=10)
        
        ttk.Button(button_frame, text="Експортувати дані в Excel", 
                  command=self.export_data).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Експортувати графіки", 
                  command=self.export_all_plots).pack(side='left', padx=5)
        
        log_frame = ttk.LabelFrame(main_frame, text="Лог експорту", padding=10)
        log_frame.pack(fill='both', expand=True)
        
        log_cnt = ttk.Frame(log_frame)
        log_cnt.pack(fill='both', expand=True)
        log_cnt.grid_rowconfigure(0, weight=1)
        log_cnt.grid_columnconfigure(0, weight=1)
        
        self.export_log = tk.Text(log_cnt, wrap='word', height=10)
        self.export_log.grid(row=0, column=0, sticky='nsew')
        
        scrollbar = ttk.Scrollbar(log_cnt, orient='vertical', command=self.export_log.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.export_log.configure(yscrollcommand=scrollbar.set)
        
        ttk.Button(log_frame, text="Очистити лог", command=lambda: self.export_log.delete('1.0', 'end')).pack(side='right', pady=5)

    # Експорт даних в Excel файл
    def export_data(self):
        if not self.simulation_results:
            messagebox.showwarning("Попередження", "Спочатку виконайте симуляцію!")
            self._log_export("Помилка: Немає результатів симуляції")
            return
        
        try:
            export_folder = Config.EXPORT_FOLDER
            if not os.path.exists(export_folder):
                os.makedirs(export_folder)
                self._log_export(f"Створено папку: {export_folder}")
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"симуляція_експорт_{timestamp}.xlsx"
            file_path = os.path.join(export_folder, file_name)
            
            self._log_export(f"Початок експорту даних...")
            self._log_export(f"Файл буде збережено: {file_path}")
            
            export_data = {}
            
            if self.export_data_type.get() in ["input", "all"] and hasattr(self, 'excel_data'):
                self._log_export("Додавання вхідних параметрів...")
                export_data["Вхідні_параметри"] = self._prepare_input_data()
            
            if self.export_data_type.get() in ["results", "all"] and self.simulation_results:
                self._log_export("Додавання результатів симуляції...")
                results_df, stats = self._prepare_results_data()
                export_data["Результати_симуляції"] = results_df
                ci_df = self._prepare_ci_data()
                if ci_df is not None:
                    self._log_export("Додавання довірчих інтервалів...")
                    export_data["Довірчі_інтервали_95"] = ci_df
            
            self._log_export("Додавання метаданих...")
            export_data["Метадані"] = self._prepare_metadata(file_name, export_folder)
            
            if not export_data:
                self._log_export("❌ Немає даних для експорту")
                messagebox.showwarning("Попередження", "Немає даних для експорту!")
                return
            
            self._export_to_excel(file_path, export_data)
            
            self._log_export(f"✅ Експорт успішно завершено!")
            self._log_export(f"📁 Файл збережено: {file_path}")
            
            if messagebox.askyesno("Відкрити папку", "Бажаєте відкрити папку з файлом?"):
                os.startfile(os.path.abspath(export_folder))
            
        except Exception as e:
            error_msg = f"Помилка при експорті: {str(e)}"
            self._log_export(f"❌ {error_msg}")
            messagebox.showerror("Помилка експорту", error_msg)

    # Підготовка вхідних даних
    def _prepare_input_data(self):
        input_df = self.excel_data.copy()
        
        if input_df.shape[0] >= Config.MIN_ROWS:
            param_names = Config.PARAMETER_NAMES[:Config.MIN_ROWS] + [""] * (input_df.shape[0] - Config.MIN_ROWS)
            input_df.iloc[:, 0] = param_names
            input_df.rename(columns={input_df.columns[0]: "Параметр"}, inplace=True)
        
        for i in range(1, len(input_df.columns)):
            input_df.rename(columns={input_df.columns[i]: f"Стовпець {i}"}, inplace=True)
        
        return input_df

    # Підготовка результатів симуляції
    def _prepare_results_data(self):
        results_by_metric = []
        valid_scenarios = empty_scenarios = 0 # Лічильники для статистики
        
        for metric_idx, metric_label in enumerate(Config.METRIC_NAMES, 1):
            metric_row = {"Метрика": metric_label}
            
            for scenario_idx, result in enumerate(self.simulation_results, 1):
                if result is None:
                    metric_row[f"Сценарій {scenario_idx}"] = ""
                    if scenario_idx == 1: # Лічимо пусті сценарії тільки один раз
                        empty_scenarios += 1
                elif metric_idx in result:
                    value = result[metric_idx]
                    metric_row[f"Сценарій {scenario_idx}"] = self._format_export_value(value)
                else:
                    metric_row[f"Сценарій {scenario_idx}"] = "(відсутня метрика)"
            
            results_by_metric.append(metric_row)
        
        valid_scenarios = len(self.simulation_results) - empty_scenarios
        results_df = pd.DataFrame(results_by_metric)
        
        stats = {'total': len(self.simulation_results),
                 'valid': valid_scenarios,
                 'empty': empty_scenarios }
        
        return results_df, stats

    # Напівширини 95 % довірчих інтервалів (якщо виконано R > 1 прогонів)
    def _prepare_ci_data(self):
        cis = getattr(self, 'simulation_ci', None)
        if not cis or not any(cis):
            return None
        rows = []
        for metric_idx, metric_label in enumerate(Config.METRIC_NAMES, 1):
            row = {"Метрика": metric_label}
            for scenario_idx, ci in enumerate(cis, 1):
                row[f"Сценарій {scenario_idx}"] = self._format_export_value(ci[metric_idx]) if ci else ""
            rows.append(row)
        return pd.DataFrame(rows)

    # Форматує значення для експорту
    def _format_export_value(self, value):
        if isinstance(value, str) and value == "Помилка":
            return "Помилка"
        
        if value is None:
            return "(пропущено)"
        
        if not isinstance(value, (int, float)):
            return str(value)
        
        try:
            if value == 0:
                return "0"
            elif abs(value) >= 1000:
                return f"{value:.2f}"
            elif abs(value) >= 1:
                return f"{value:.6f}"
            else:
                return f"{value:.8f}"
        except:
            return str(value)

    # Підготовка метаданих
    def _prepare_metadata(self, file_name, export_folder):
        valid_count = len([r for r in self.simulation_results if r is not None])
        empty_count = len([r for r in self.simulation_results if r is None])
        
        metadata = {
            "Дата_експорту": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            "Ім'я_файлу": [file_name],
            "Загальна_кількість_сценаріїв": [valid_count + empty_count],
            "Валідних_сценаріїв": [valid_count],
            "Порожніх_сценаріїв": [empty_count],
            "Кількість_метрик": [len(Config.METRIC_NAMES)],
            "Версія_програми": [Config.APP_VERSION],
            "Кількість_прогонів_R": [getattr(self, 'last_run_settings', {}).get('replications', 1)],
            "Seed": [getattr(self, 'last_run_settings', {}).get('seed', '')],
            "Оригінальний_файл": [os.path.basename(self.current_file_path) if self.current_file_path else "Не вказано"],
            "Папка_збереження": [os.path.abspath(export_folder)]
        }
        
        return pd.DataFrame(metadata)

    # Експорт всіх доступних графіків в окремі файли
    def export_all_plots(self):
        if not self.simulation_results:
            messagebox.showwarning("Попередження", "Спочатку виконайте симуляцію!")
            return
        
        try:
            export_folder = Config.PLOTS_FOLDER
            if not os.path.exists(export_folder):
                os.makedirs(export_folder)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            plot_folder = os.path.join(export_folder, f"графіки_{timestamp}")
            os.makedirs(plot_folder, exist_ok=True)            
            self._log_export(f"Початок експорту графіків у: {plot_folder}")
            plot_types = list(self.plot_combobox['values'])
            saved_plots = []
            
            for plot_type in plot_types:
                try:
                    self._log_export(f"Створення: {plot_type}...")
                    
                    fig, ax = plt.subplots(figsize=(10, 6))
                    
                    plot_methods = {
                        "Час очікування vs тривалість вторинної обробки": self.generate_waiting_time_plot,
                        "Максимальний час очікування vs тривалість вторинної обробки": self.generate_max_waiting_plot,
                        "Довжина черги vs тривалість вторинної обробки": self.generate_queue_length_plot,
                        "Порушення дедлайнів vs часові обмеження": self.generate_deadline_plot
                    }
                    
                    plot_method = plot_methods.get(plot_type)
                    if plot_method:
                        plot_method(ax)
                    
                    filename = plot_type.replace(" ", "_").replace(":", "").replace("/", "_")
                    filename = f"графік_{filename}_{timestamp}.png"
                    filepath = os.path.join(plot_folder, filename)
                    
                    fig.savefig(filepath, dpi=Config.PLOT_DPI, bbox_inches='tight')
                    plt.close(fig)
                    
                    saved_plots.append(filepath)
                    self._log_export(f"  ✓ Збережено: {filename}")
                    
                except Exception as e:
                    self._log_export(f"  ✗ Помилка: {str(e)}")
            
            if saved_plots:
                self._log_export(f"✅ Експорт завершено. Збережено: {len(saved_plots)} графіків")
                
                if messagebox.askyesno("Відкрити папку", "Бажаєте відкрити папку з графіками?"):
                    os.startfile(os.path.abspath(plot_folder))
            else:
                self._log_export("❌ Не вдалося створити графіки")
                
        except Exception as e:
            self._log_export(f"❌ Помилка: {str(e)}")
            messagebox.showerror("Помилка експорту", str(e))

    # Експорт у Excel
    def _export_to_excel(self, file_path, data_dict):
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in data_dict.items():
                sheet_name = str(sheet_name)[:31]
                
                if not isinstance(df, pd.DataFrame):
                    df = pd.DataFrame(df)
                
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                try:
                    worksheet = writer.sheets[sheet_name]
                    for idx, column in enumerate(df.columns, 1):
                        try:
                            if df[column].dtype == 'object':
                                max_length = df[column].astype(str).str.len().max()
                            else:
                                max_length = max(df[column].astype(str).str.len().max(), len(str(df[column].dtype)))
                            
                            column_length = max(max_length, len(str(column))) + 2
                            column_width = min(column_length, 50)
                            
                            column_letter = get_column_letter(idx)
                            worksheet.column_dimensions[column_letter].width = column_width
                        except:
                            continue
                except:
                    pass   

    # Логування експорту
    def _log_export(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.export_log.insert('end', log_entry)
        self.export_log.see('end')
        print(log_entry.strip())
