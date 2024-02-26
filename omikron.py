# Omikron Program Entry Point

# import library
import os
import tkinter as tk

from omikron.gui import GUI

# initiate program directory structure
if not os.path.exists("./data"):
    os.makedirs("./data")
if not os.path.exists("./data/backup"):
    os.makedirs("./data/backup")
os.environ["WDM_PROGRESS_BAR"] = "0"

# GUI initiation
ui = tk.Tk()
gui = GUI(ui)
ui.after(10, gui.print_log)
ui.after(100, gui.check_files)
ui.after(100, gui.check_thread_end)
ui.mainloop()