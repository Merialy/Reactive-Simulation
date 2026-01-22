import tkinter as tk
from tkinter import ttk
from utils.encoding import setup_utf8_output

class GUIBase:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Симулятор моделі реактивності")
        self.root.geometry("800x500")
        
        setup_utf8_output()
        
        # Ініціалізація змінних
        self._init_variables()
        self.create_widgets()

    # Ініціалізація всіх змінних класу
    def _init_variables(self):
        self.excel_data = self.simulation_results = self.current_plot_frame = self.current_file_path = None # Дані та шлях
        self.num_rows = self.num_columns = self.data_columns = 0 # Статистика (числа)
        self.s2_values, self.d1_values = [], [] # Списки
        self.editing_item = self.editing_column = self.editing_row = self.entry_edit = None 
        self._is_validating = False

    # ==================== СТВОРЕННЯ ІНТЕРФЕЙСУ ====================

    # Створення всіх вкладок
    def create_widgets(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.tab1 = ttk.Frame(notebook) 
        self.tab2 = ttk.Frame(notebook)
        self.tab3 = ttk.Frame(notebook)
        self.tab4 = ttk.Frame(notebook)
        
        notebook.add(self.tab1, text="Завантаження даних")
        notebook.add(self.tab2, text="Симуляція")
        notebook.add(self.tab3, text="Візуалізація")
        notebook.add(self.tab4, text="Експорт даних")
        
        self.create_data_tab()
        self.create_simulation_tab()
        self.create_visualization_tab()
        self.create_export_tab()

    # Створює Treeview з скролбарами та контейнером
    def _setup_treeview_with_scrollbars(self, parent, columns, height=10):
        container = ttk.Frame(parent)
        container.pack(fill='both', expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        tree = ttk.Treeview(container, columns=columns, show='headings', height=height)
        
        sy = ttk.Scrollbar(container, orient='vertical', command=tree.yview)
        sx = ttk.Scrollbar(container, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        sy.grid(row=0, column=1, sticky='ns')
        sx.grid(row=1, column=0, sticky='ew')
        
        return tree, container
