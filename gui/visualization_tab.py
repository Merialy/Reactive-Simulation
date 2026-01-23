import tkinter as tk
import numpy as np 
import matplotlib.pyplot as plt
from tkinter import ttk, messagebox
from config import Config
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ==================== ВКЛАДКА 3: ВІЗУАЛІЗАЦІЯ ==================== 
class VisualizationMixin:
    
    def create_visualization_tab(self):
        main_frame = ttk.Frame(self.tab3)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        control_frame = ttk.LabelFrame(main_frame, text="Керування графіками", padding=10)
        control_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(control_frame, text="Оберіть графік:").pack(side='left', padx=5)
        
        self.plot_var = tk.StringVar()
        self.plot_combobox = ttk.Combobox(control_frame, textvariable=self.plot_var, state='readonly', width=50)
        self.plot_combobox['values'] = []
        self.plot_combobox.pack(side='left', padx=5)
        self.plot_combobox.bind('<<ComboboxSelected>>', self.on_plot_selected)
        
        ttk.Button(control_frame, text="Побудувати графік", command=self.generate_plot).pack(side='left', padx=5)
        
        self.plot_description = tk.Text(control_frame, height=3, width=60, wrap='word', state='disabled')
        self.plot_description.pack(side='left', padx=10)
        
        self.plot_frame = ttk.Frame(main_frame)
        self.plot_frame.pack(fill='both', expand=True)

    def on_plot_selected(self, event=None):
        plot_type = self.plot_var.get()
        
        descriptions = {
            "Час очікування vs тривалість вторинної обробки":"Графік залежності середнього часу очікування від тривалості вторинної обробки подій м'якого РЧ",
            "Максимальний час очікування vs тривалість вторинної обробки":"Графік залежності максимального часу очікування від тривалості вторинної обробки",
            "Довжина черги vs тривалість вторинної обробки":"Графік залежності середньої довжини черги від тривалості вторинної обробки",
            "Порушення дедлайнів vs часові обмеження":"Графік залежності ймовірності порушення дедлайнів від часових обмежень"
        }
        
        if hasattr(self, 'plot_description'):
            self.plot_description.config(state='normal')
            self.plot_description.delete('1.0', 'end')
            self.plot_description.insert('1.0', descriptions.get(plot_type, "Опис відсутній"))
            self.plot_description.config(state='disabled')

    def update_plot_options_based_on_s2_and_d1(self):
        has_s2 = hasattr(self, 's2_values') and self.s2_values
        has_d1 = hasattr(self, 'd1_values') and self.d1_values
        
        if not has_s2 and not has_d1:
            self.plot_combobox['values'] = []
            self.plot_var.set("")
            return
        
        can_plot_s2 = self._can_plot_varying_data(self.s2_values) if has_s2 else False
        can_plot_d1 = self._can_plot_varying_data(self.d1_values) if has_d1 else False
        
        options = []
        
        if can_plot_s2:
            options.extend([
                "Час очікування vs тривалість вторинної обробки",
                "Максимальний час очікування vs тривалість вторинної обробки",
                "Довжина черги vs тривалість вторинної обробки"
            ])
        
        if can_plot_d1:
            options.append("Порушення дедлайнів vs часові обмеження")
        
        self.plot_combobox['values'] = options
        
        if options:
            self.plot_var.set(options[0])
            self.on_plot_selected()
        else:
            self.plot_var.set("")

    def _can_plot_varying_data(self, value_groups):
        for sublist in value_groups:
            if isinstance(sublist, (list, tuple, np.ndarray)) and len(sublist) > 1:
                values = [float(x) for x in sublist]
                if len(set(values)) > 1:
                    return True
        return False

    def generate_plot(self):
        plot_type = self.plot_var.get()
        
        if not self.simulation_results:
            messagebox.showwarning("Попередження", "Спочатку виконайте симуляцію!")
            return
        
        if not plot_type:
            messagebox.showwarning("Попередження", "Оберіть тип графіка!")
            return
                
        try:
            for widget in self.plot_frame.winfo_children():
                widget.destroy()
            
            fig, ax = plt.subplots(figsize=(8, 5))
            
            plot_methods = {
                "Час очікування vs тривалість вторинної обробки": self.generate_waiting_time_plot,
                "Максимальний час очікування vs тривалість вторинної обробки": self.generate_max_waiting_plot,
                "Довжина черги vs тривалість вторинної обробки": self.generate_queue_length_plot,
                "Порушення дедлайнів vs часові обмеження": self.generate_deadline_plot
            }
            
            plot_method = plot_methods.get(plot_type)
            if plot_method:
                plot_method(ax)
            else:
                messagebox.showwarning("Попередження", "Невідомий тип графіка!")
                return
            
            canvas = FigureCanvasTkAgg(fig, self.plot_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
            self.current_plot_frame = fig
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося побудувати графік: {str(e)}")
            import traceback
            traceback.print_exc()

    def _generate_s2_based_plot(self, ax, metric_indices, labels, title, ylabel, ylim=None):
        s2_data = self._get_varying_values(self.s2_values)
        if not s2_data:
            messagebox.showwarning("Попередження", "Немає даних s2 з різними значеннями!")
            return
        
        s2_flat = []
        for group in s2_data:
            if isinstance(group, (list, tuple, np.ndarray)):
                s2_flat.extend(group)
            else:
                s2_flat.append(group)
        
        # Сортуємо для коректного відображення на графіку
        s2_flat = sorted(s2_flat, reverse=True)
        num_points = len(s2_flat)        
        
        valid_results = [r for r in self.simulation_results if r is not None][:num_points]
        
        if len(valid_results) < num_points:
            messagebox.showwarning("Попередження", f"Недостатньо результатів: {len(valid_results)} з {num_points}")
            return
        
        x_data = np.array(s2_flat)
        y_data_list = []
        
        for metric_idx in metric_indices:
            y_data = [self._extract_metric(r, metric_idx) for r in valid_results]
            y_data_list.append(np.array(y_data))
        
        valid_mask = np.all([~np.isnan(y) for y in y_data_list], axis=0)
        if not np.any(valid_mask):
            messagebox.showerror("Помилка", "Немає валідних даних!")
            return
        
        x_filtered = x_data[valid_mask]
        y_filtered_list = [y[valid_mask] for y in y_data_list]
        
        ax.set_yscale('log')
        ax.set_xlim(min(x_filtered) - 0.05, max(x_filtered) + 0.05)
        if ylim:
            ax.set_ylim(*ylim)
        
        for i, (y_data, label) in enumerate(zip(y_filtered_list, labels)):
            color = Config.PLOT_COLORS[i % len(Config.PLOT_COLORS)]
            marker = Config.PLOT_MARKERS[i % len(Config.PLOT_MARKERS)]
            ax.plot(x_filtered, y_data, f'{marker}-', 
                label=label, linewidth=2, markersize=8, color=color)
        
        ax.set_xlabel('Тривалість вторинної обробки (s2), мс', fontsize=10)
        ax.set_ylabel(ylabel, fontsize=10)
        ax.set_title(title, fontsize=10, fontweight='bold', pad=20)
        ax.grid(True, which='both', linestyle='--', alpha=0.1)
        ax.legend(loc='best')
        plt.tight_layout()

    def _extract_metric(self, result, metric_idx):
        try:
            value = result.get(metric_idx)
            return float(value) if value is not None else np.nan
        except:
            return np.nan

    def _get_varying_values(self, value_groups):
        result = []
        for sublist in value_groups:
            if isinstance(sublist, (list, tuple, np.ndarray)): # 1. Перевіряємо, що це колекція
                unique_set = {float(x) for x in sublist} # 2. Перетворюємо на float і відразу прибираємо дублікати через set
                if len(unique_set) > 1:# 3. Якщо унікальних значень більше одного — це те, що нам потрібно
                    result.append(list(unique_set))# Зберігаємо як список (можна додати sorted(), якщо важливий порядок)
                    
        return result if result else None # Повертаємо список списків або None, якщо нічого не підійшло

    def generate_waiting_time_plot(self, ax):
        self._generate_s2_based_plot(
            ax,
            metric_indices=[1, 2, 3],
            labels=['Первинна обробка', 'Вторинна обробка (м\'який РЧ)', 'Вторинна обробка (жорсткий РЧ)'],
            title='Залежність часу очікування від тривалості вторинної обробки',
            ylabel='Середній час очікування, мс',
            ylim=(0.00005, 1000)
        )
        ax.set_yticks([0.000001, 0.00001, 0.0001, 0.001, 0.01, 0.1, 1, 10, 100, 1000])
        ax.set_yticklabels(['0,000001', '0,00001', '0,0001', '0,001', '0,01', '0,1', '1', '10', '100', '1000'])

    def generate_max_waiting_plot(self, ax):
        self._generate_s2_based_plot(
            ax,
            metric_indices=[6, 7, 8],
            labels=['Макс. час первинної обробки', 'Макс. час вторинної (м\'який)', 'Макс. час вторинної (жорсткий)'],
            title='Залежність максимального часу очікування від тривалості вторинної обробки',
            ylabel='Максимальний час очікування, мс',
            ylim=(0.00005, 1000)
        )
        ax.set_yticks([0.000001, 0.00001, 0.0001, 0.001, 0.01, 0.1, 1, 10, 100, 1000])
        ax.set_yticklabels(['0,000001', '0,00001', '0,0001', '0,001', '0,01', '0,1', '1', '10', '100', '1000'])

    def generate_queue_length_plot(self, ax):
        self._generate_s2_based_plot(
            ax,
            metric_indices=[11, 12, 13],
            labels=['Довжина черги первинного обробника', 'Довжина черги вторинного (м\'який)', 'Довжина черги вторинного (жорсткий)'],
            title='Залежність довжини черги від тривалості вторинної обробки',
            ylabel='Середня довжина черги',
            ylim=(0.00005, 1000)
        )
        ax.set_yticks([0.000001, 0.00001, 0.0001, 0.001, 0.01, 0.1, 1, 10, 100, 1000])
        ax.set_yticklabels(['0,000001', '0,00001', '0,0001', '0,001', '0,01', '0,1', '1', '10', '100', '1000'])

    def generate_deadline_plot(self, ax):
        if not self.d1_values or not self.simulation_results:
            messagebox.showwarning("Попередження", "Немає даних для побудови графіка!")
            return
        
        s2_labels = self._get_s2_labels_for_groups()
        valid_results = [r for r in self.simulation_results if r is not None]
        if not valid_results:
            messagebox.showwarning("Попередження", "Немає валідних результатів!")
            return
        
        current_result_idx = 0
        lines_plotted = 0
        
        for i, d1_group in enumerate(self.d1_values):
            if self._is_constant_group(d1_group):
                current_result_idx += len(d1_group)
                continue
            
            num_points = len(d1_group)
            available_points = len(valid_results) - current_result_idx
            
            if available_points < num_points:
                num_points = min(num_points, available_points)
                if num_points < 2:
                    current_result_idx += num_points
                    continue
            
            group_results = valid_results[current_result_idx:current_result_idx + num_points]
            current_result_idx += num_points
            
            x_data = np.array(d1_group[:num_points])
            y_data = [self._extract_metric(r, 17) for r in group_results]
            y_data = np.array(y_data)
            
            # Угруповання однакових значень d1
            x_data, y_data = self._aggregate_duplicate_x_values(x_data, y_data)
            
            valid_mask = ~np.isnan(y_data)
            if np.sum(valid_mask) >= 2:
                color = Config.PLOT_COLORS[lines_plotted % len(Config.PLOT_COLORS)]
                marker = Config.PLOT_MARKERS[lines_plotted % len(Config.PLOT_MARKERS)]
                label = self._create_s2_based_label(i, s2_labels)
                
                sorted_indices = np.argsort(x_data[valid_mask])
                x_sorted = x_data[valid_mask][sorted_indices]
                y_sorted = y_data[valid_mask][sorted_indices]
                
                ax.plot(x_sorted, y_sorted, marker=marker, linestyle='-', 
                    linewidth=2, markersize=8, color=color, label=label, alpha=0.8)
                lines_plotted += 1
        
        if lines_plotted > 0:
            ax.set_xlabel('Часові обмеження (d1), мс', fontsize=11, fontweight='bold')
            ax.set_ylabel('Ймовірність порушення дедлайнів', fontsize=11, fontweight='bold')
            ax.set_title('Залежність ймовірності порушення дедлайнів від часових обмежень', 
                        fontsize=12, fontweight='bold', pad=20)
            ax.grid(True, alpha=0.3, linestyle='--')
            ax.set_ylim(-0.05, 1.05)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.legend(loc='best', fontsize='small' if lines_plotted > 4 else 'medium', framealpha=0.9)
            plt.tight_layout()
        else:
            messagebox.showwarning("Попередження", "Не вдалося побудувати графік!")

    # Групує однакові значення x та усереднює відповідні значення y
    def _aggregate_duplicate_x_values(self, x_data, y_data):
        if len(x_data) == 0:
            return x_data, y_data
        
        data_dict = {}
        for x, y in zip(x_data, y_data):
            x_rounded = round(float(x), 3)
            if x_rounded not in data_dict:
                data_dict[x_rounded] = []
            data_dict[x_rounded].append(y)
        
        x_unique = []
        y_averaged = []
        for x_val in sorted(data_dict.keys()):
            x_unique.append(x_val)
            y_values = [y for y in data_dict[x_val] if not np.isnan(y)]
            if y_values:
                y_averaged.append(np.mean(y_values))
            else:
                y_averaged.append(np.nan)
        
        return np.array(x_unique), np.array(y_averaged)

    def _is_constant_group(self, group):
        if not group or len(group) <= 1:
            return True
        
        first_val = float(group[0])
        for val in group[1:]:
            if abs(float(val) - first_val) > 0.001:
                return False
        return True

    def _get_s2_labels_for_groups(self):
        s2_labels = []
        
        if not hasattr(self, 's2_values') or not self.s2_values:
            return s2_labels
        
        for sublist in self.s2_values:
            if isinstance(sublist, (list, tuple, np.ndarray)) and len(sublist) > 0:
                try:
                    values = [float(x) for x in sublist]                    
                    if not values:
                        s2_labels.append(None)
                        continue
                    
                    if self._is_constant_group(values):
                        s2_labels.append(values[0])
                    else:
                        unique_vals, counts = np.unique(np.round(values, 2), return_counts=True)
                        most_common_idx = np.argmax(counts)
                        most_common_val = unique_vals[most_common_idx]
                        
                        if counts[most_common_idx] > len(values) * 0.3:
                            s2_labels.append(most_common_val)
                        else:
                            s2_labels.append(round(np.mean(values), 2))
                except Exception as e:
                    print(f"Error processing s2 group: {e}")
                    s2_labels.append(None)
            else:
                s2_labels.append(None)
        
        return s2_labels

    def _create_s2_based_label(self, group_index, s2_labels):
        if group_index < len(s2_labels) and s2_labels[group_index] is not None:
            s2_value = s2_labels[group_index]
            formatted_value = self._format_s2_value(s2_value)
            return f"Час вторинної обробки - {formatted_value} мс"
        return self._get_label_from_data(group_index)

    def _format_s2_value(self, value):
        if abs(value - round(value)) < 0.001:
            return f"{int(round(value))}"
        
        formatted = f"{value:.2f}"
        if formatted.endswith('.00'):
            return formatted[:-3]
        elif formatted.endswith('0'):
            return formatted.rstrip('0').rstrip('.')
        return formatted

    def _get_label_from_data(self, group_index):
        try:
            if hasattr(self, 's2_values') and group_index < len(self.s2_values):
                group = self.s2_values[group_index]
                if isinstance(group, (list, tuple, np.ndarray)) and len(group) > 0:
                    first_val = float(group[0])
                    return f"Час вторинної обробки - {self._format_s2_value(first_val)} мс"
        except:
            pass
        return f"Група {group_index + 1}"
    
    