# Omikron Program Entry Point

# import library
import os
import tkinter as tk
import omikron.config

from omikron.gui import GUI

# initiate program directory structure
if not os.path.exists(f"{omikron.config.DATA_DIR}/data"):
    os.makedirs(f"{omikron.config.DATA_DIR}/data")
if not os.path.exists(f"{omikron.config.DATA_DIR}/data/backup"):
    os.makedirs(f"{omikron.config.DATA_DIR}/data/backup")
os.environ["WDM_PROGRESS_BAR"] = "0"

# GUI initiation
ui = tk.Tk()
gui = GUI(ui)
ui.after(10, gui.print_log)
ui.after(100, gui.check_files)
ui.after(100, gui.check_thread_end)
ui.mainloop()