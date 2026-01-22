from tkinter import ttk, messagebox
from config import Config
from simulation.priority_simulator import PriorityQueueSimulation

# ==================== ВКЛАДКА 2: СИМУЛЯЦІЯ ====================
class SimulationMixin:
    
    def create_simulation_tab(self):
        main_frame = ttk.Frame(self.tab2)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        control_frame = ttk.LabelFrame(main_frame, text="Керування симуляцією", padding=10)
        control_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(control_frame, text="Запустити симуляцію", command=self.run_simulation).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Очистити результати", command=self.clear_results).pack(side='left', padx=5)
        
        self.status_label = ttk.Label(control_frame, text="Статус: Очікування запуску")
        self.status_label.pack(side='left', padx=20)
        
        self.progress = ttk.Progressbar(control_frame, mode='determinate', length=300)
        self.progress.pack(side='left', padx=5)
        
        results_frame = ttk.LabelFrame(main_frame, text="Результати симуляції", padding=10)
        results_frame.pack(fill='both', expand=True)
        
        self.results_tree, _ = self._setup_treeview_with_scrollbars(results_frame, ['Метрика'], height=20)
        self.results_tree.heading('Метрика', text='Метрика')
        self.results_tree.column('Метрика', width=300, anchor='w')

    def run_simulation(self):
        if not hasattr(self, 'excel_parameters') or not self.excel_parameters:
            messagebox.showwarning("Попередження", "Спочатку завантажте дані з Excel файлу!")
            return
        
        try:
            self.status_label.config(text="Статус: Запуск симуляції...")
            self.progress['value'] = 0
            self.root.update_idletasks()
            self.clear_results()
            
            all_results = []
            successful, skipped, errors = 0, 0, 0
            
            # Використовуємо ПОТОЧНУ кількість параметрів (після видалення стовпців)
            total_params = len(self.excel_parameters)
            
            for i, params in enumerate(self.excel_parameters, 1):
                if params is None:
                    all_results.append(None)
                    skipped += 1
                    continue
                
                try:
                    self._validate_params(params)
                    
                    progress_value = (i / total_params) * 100
                    self.progress['value'] = progress_value
                    self.status_label.config(text=f"Статус: Симуляція {i}/{total_params}")
                    self.root.update_idletasks()
                    
                    sim = PriorityQueueSimulation(seed=41 + i)
                    result = sim.run_simulation_priority2_full(**params)
                    all_results.append(result)
                    successful += 1
                    
                except Exception as e:
                    print(f"Симуляція {i}: помилка - {e}")
                    error_result = {j: "Помилка" for j in range(1, 19)}
                    all_results.append(error_result)
                    errors += 1
            
            self.simulation_results = all_results
            self.display_results(all_results)
            
            status_text = f"Статус: Завершено. Успішно: {successful}"
            if skipped > 0:
                status_text += f", Пропущено: {skipped}"
            if errors > 0:
                status_text += f", Помилок: {errors}"
            
            self.status_label.config(text=status_text)
            self.progress['value'] = 100
            
            print(f"\n📊 Симуляція завершена: {total_params} параметрів обробл ено")
            
        except Exception as e:
            self.status_label.config(text="Статус: Помилка при симуляції")
            messagebox.showerror("Помилка", f"Не вдалося виконати симуляцію: {str(e)}")

    def _validate_params(self, params):
        for key, value in params.items():
            if not isinstance(value, (int, float)):
                raise ValueError(f"Параметр {key} не числовий")

    def display_results(self, all_results):
        if not all_results:
            return
        
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        num_columns = len(all_results) + 1
        columns = ['Метрика'] + [f'Сценарій {i}' for i in range(1, num_columns)]
        self.results_tree['columns'] = columns
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            if col == 'Метрика':
                self.results_tree.column(col, width=350, anchor='w', minwidth=410)
            else:
                self.results_tree.column(col, width=100, anchor='center', minwidth=80)
        
        for metric_idx, metric_label in enumerate(Config.METRIC_NAMES, 1):
            row_values = [metric_label]
            
            for result in all_results:
                row_values.append(self._format_result_value(result, metric_idx))
            
            self.results_tree.insert('', 'end', values=row_values)

    def _format_result_value(self, result, metric_idx):
        if result is None:
            return ""
        
        if isinstance(result, dict) and metric_idx in result:
            value = result[metric_idx]
            
            if isinstance(value, str) and value == "Помилка":
                return "Помилка"
            
            if not isinstance(value, (int, float)):
                return str(value)
            
            return self._format_numeric_value(value)
        
        return ""

    def _format_numeric_value(self, value):
        try:
            if value == 0:
                return "0"
            elif abs(value) >= 1000:
                return f"{value:.2f}"
            elif abs(value) >= 1:
                return f"{value:.6f}"
            else:
                return f"{value:.6f}"
        except:
            return str(value)

    def clear_results(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        self.results_tree['columns'] = ['Метрика']
        self.results_tree.heading('Метрика', text='Метрика')
        self.results_tree.column('Метрика', width=300, anchor='w')
        
        self.simulation_results = None
        self.status_label.config(text="Статус: Очікування запуску")
        self.progress['value'] = 0

  
