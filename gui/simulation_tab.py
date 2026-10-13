import queue
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
from config import Config
from simulation.replications import run_batch
from simulation.validation import check_params

# ==================== ВКЛАДКА 2: СИМУЛЯЦІЯ ====================
class SimulationMixin:
    
    def create_simulation_tab(self):
        main_frame = ttk.Frame(self.tab2)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        control_frame = ttk.LabelFrame(main_frame, text="Керування симуляцією", padding=10)
        control_frame.pack(fill='x', pady=(0, 10))
        
        self.run_button = ttk.Button(control_frame, text="Запустити симуляцію", command=self.run_simulation)
        self.run_button.pack(side='left', padx=5)
        ttk.Button(control_frame, text="Очистити результати", command=self.clear_results).pack(side='left', padx=5)

        # Кількість незалежних прогонів R і початкове значення генератора (п. 2.6.5)
        ttk.Label(control_frame, text="Прогонів R:").pack(side='left', padx=(15, 2))
        self.replications_var = tk.IntVar(value=Config.DEFAULT_REPLICATIONS)
        ttk.Spinbox(control_frame, from_=1, to=50, width=4, textvariable=self.replications_var).pack(side='left')
        ttk.Label(control_frame, text="Seed:").pack(side='left', padx=(10, 2))
        self.seed_var = tk.IntVar(value=Config.DEFAULT_SEED)
        ttk.Spinbox(control_frame, from_=0, to=10**9, width=8, textvariable=self.seed_var).pack(side='left')
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill='x', pady=(0, 10))
        self.status_label = ttk.Label(status_frame, text="Статус: Очікування запуску")
        self.status_label.pack(side='left', padx=5)

        self.progress = ttk.Progressbar(status_frame, mode='determinate', length=300)
        self.progress.pack(side='right', padx=5)
        
        # Попередження щодо параметрів – у панелі, а не у спливаючому вікні
        self.warnings_text = tk.Text(main_frame, height=3, wrap='word', background='#fff8e1',
                                     relief='flat', font=('Segoe UI', 9))
        self.warnings_text.pack(fill='x', pady=(0, 8))
        self.warnings_text.insert('1.0', "Попередження щодо параметрів з'являться тут після запуску.")
        self.warnings_text.config(state='disabled')

        results_frame = ttk.LabelFrame(main_frame, text="Результати симуляції", padding=10)
        results_frame.pack(fill='both', expand=True)
        
        self.results_tree, _ = self._setup_treeview_with_scrollbars(results_frame, ['Метрика'], height=20)
        self.results_tree.heading('Метрика', text='Метрика')
        self.results_tree.column('Метрика', width=300, anchor='w')

    def run_simulation(self):
        if not hasattr(self, 'excel_parameters') or not self.excel_parameters:
            messagebox.showwarning("Попередження", "Спочатку завантажте дані з Excel файлу!")
            return
        if getattr(self, '_sim_running', False):
            return

        try:
            replications = max(1, int(self.replications_var.get()))
            base_seed = int(self.seed_var.get())
        except (tk.TclError, ValueError):
            messagebox.showerror("Помилка", "Кількість прогонів і seed мають бути цілими числами.")
            return

        self.clear_results()
        self.last_run_settings = {'replications': replications, 'seed': base_seed}

        # ---- Перевірка параметрів усіх сценаріїв (таблиця 2.3) ----
        run_list, warnings_report, self._sim_errors = [], [], {}
        skipped = 0
        for i, params in enumerate(self.excel_parameters, 1):
            if params is None:
                run_list.append(None)
                skipped += 1
                continue
            try:
                self._validate_params(params)
                param_errors, param_warnings = check_params(params)
            except ValueError as e:
                param_errors, param_warnings = [str(e)], []
            if param_errors:
                self._sim_errors[i - 1] = "; ".join(param_errors)
                warnings_report.append(f"Сценарій {i}: ПОМИЛКА – {'; '.join(param_errors)}")
                run_list.append(None)
                continue
            warnings_report += [f"Сценарій {i}: {w}" for w in param_warnings]
            run_list.append(params)

        self._show_warnings(warnings_report)
        self._sim_run_list = run_list
        self._sim_skipped = skipped
        self._sim_warned = {int(w.split(':')[0].split()[1]) for w in warnings_report}

        # ---- Запуск у фоновому потоці, щоб вікно не «зависало» ----
        self._sim_running = True
        self._sim_started = time.time()
        self.run_button.config(state='disabled')
        self.progress['value'] = 0
        self.status_label.config(text=f"Статус: Симуляція... (R = {replications})")
        self._sim_queue = queue.Queue()

        def worker():
            try:
                out = run_batch(run_list, replications=replications, base_seed=base_seed,
                                workers=Config.WORKERS,
                                progress_cb=lambda d, t: self._sim_queue.put(('progress', d, t)))
                self._sim_queue.put(('done', out))
            except Exception as e:
                self._sim_queue.put(('error', str(e)))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self._poll_simulation)

    def _poll_simulation(self):
        try:
            while True:
                msg = self._sim_queue.get_nowait()
                if msg[0] == 'progress':
                    _, done, total = msg
                    self.progress['value'] = done / total * 100
                    n_sc = len([p for p in self._sim_run_list if p is not None])
                    self.status_label.config(
                        text=f"Статус: Симуляція... прогін {done} з {total} "
                             f"({n_sc} сценаріїв × {self.last_run_settings['replications']} прогонів)")
                elif msg[0] == 'done':
                    self._finish_simulation(msg[1])
                    return
                elif msg[0] == 'error':
                    self._sim_running = False
                    self.run_button.config(state='normal')
                    self.status_label.config(text="Статус: Помилка при симуляції")
                    messagebox.showerror("Помилка", f"Не вдалося виконати симуляцію: {msg[1]}")
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll_simulation)

    def _finish_simulation(self, batch):
        replications = self.last_run_settings['replications']
        all_results, all_ci = [], []
        successful = errors = 0
        for i, rep in enumerate(batch):
            if i in self._sim_errors:
                all_results.append({j: "Помилка" for j in range(1, 19)})
                all_ci.append(None)
                errors += 1
            elif rep is None:
                all_results.append(None)
                all_ci.append(None)
            else:
                all_results.append(rep['mean'])
                all_ci.append(rep['ci'] if replications > 1 else None)
                successful += 1

        self.simulation_ci = all_ci
        self.simulation_results = all_results
        self.display_results(all_results)

        elapsed = time.time() - self._sim_started
        status_text = f"Статус: Завершено за {elapsed:.1f} с. Успішно: {successful}"
        if self._sim_skipped:
            status_text += f", Пропущено: {self._sim_skipped}"
        if errors:
            status_text += f", Помилок: {errors}"
        if self._sim_warned:
            status_text += f", з попередженнями: {len(self._sim_warned)} (позначено ⚠)"
        self.status_label.config(text=status_text)
        self.progress['value'] = 100
        self._sim_running = False
        self.run_button.config(state='normal')
        print(f"📊 Симуляція завершена: {len(batch)} сценаріїв, R = {replications}, {elapsed:.1f} с")

    def _show_warnings(self, lines):
        self.warnings_text.config(state='normal')
        self.warnings_text.delete('1.0', 'end')
        if lines:
            self.warnings_text.insert('1.0', "⚠ " + "\n⚠ ".join(lines))
            self.warnings_text.config(height=min(6, max(2, len(lines))))
        else:
            self.warnings_text.insert('1.0', "✓ Параметри всіх сценаріїв коректні.")
            self.warnings_text.config(height=1)
        self.warnings_text.config(state='disabled')

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
        warned = getattr(self, '_sim_warned', set())

        for col in columns:
            title = col
            if col != 'Метрика' and int(col.split()[1]) in warned:
                title = f"⚠ {col}"
            self.results_tree.heading(col, text=title)
            if col == 'Метрика':
                self.results_tree.column(col, width=430, anchor='w', minwidth=300, stretch=False)
            else:
                self.results_tree.column(col, width=150, anchor='center', minwidth=90, stretch=False)
        
        for metric_idx, metric_label in enumerate(Config.METRIC_NAMES, 1):
            row_values = [metric_label]
            
            for k, result in enumerate(all_results):
                text = self._format_result_value(result, metric_idx)
                ci = self.simulation_ci[k] if getattr(self, 'simulation_ci', None) else None
                if ci and text not in ("", "Помилка") and metric_idx in ci:
                    text = f"{text} ± {self._format_sig(ci[metric_idx], 2)}"
                row_values.append(text)
            
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
        """Середнє значення: 4 значущі цифри (без зайвих нулів)."""
        return self._format_sig(value, 4)

    @staticmethod
    def _format_sig(value, sig):
        try:
            if value == 0:
                return "0"
            return f"{value:.{sig}g}"
        except (TypeError, ValueError):
            return str(value)

    def clear_results(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        self.results_tree['columns'] = ['Метрика']
        self.results_tree.heading('Метрика', text='Метрика')
        self.results_tree.column('Метрика', width=300, anchor='w')
        
        self.simulation_results = None
        self.simulation_ci = None
        self.status_label.config(text="Статус: Очікування запуску")
        self.progress['value'] = 0

  
