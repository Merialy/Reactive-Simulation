import tkinter as tk
from gui.base import GUIBase
from gui.data_tab import DataTabMixin
from gui.simulation_tab import SimulationMixin
from gui.visualization_tab import VisualizationMixin
from gui.export_tab import ExportMixin

# Собираем класс обратно из кирпичиков
class SimulationGUI(GUIBase, DataTabMixin, SimulationMixin, VisualizationMixin, ExportMixin):
    pass

if __name__ == "__main__":
    root = tk.Tk()
    app = SimulationGUI(root)
    root.mainloop()