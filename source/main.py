import os
import tkinter as tk

from gui import GUI

if not os.path.exists("./data"):
    os.makedirs("./data")
if not os.path.exists("./data/backup"):
    os.makedirs("./data/backup")

os.environ["WDM_PROGRESS_BAR"] = "0"

ui = tk.Tk()
gui = GUI(ui)
ui.after(100, gui.thread_log)
ui.after(100, gui.check_files)
ui.after(100, gui.check_thread_end)
ui.mainloop()