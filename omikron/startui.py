import tkinter as tk

def start_ui():
    root = tk.Tk()
    img = tk.PhotoImage(file="omikron/omikron.png")
    root.configure(bg='white')
    canvas = tk.Canvas(root, width=img.width(), height=img.width(), bg="white")
    canvas.pack()
    width = img.width()
    height = img.height()
    x = int((root.winfo_screenwidth()/2) - (width/2))
    y = int((root.winfo_screenheight()/2) - (height/2))
    root.geometry(f"{width}x{height}+{x}+{y}")

    root.overrideredirect(True)
    # root.wm_attributes('-alpha', 0.3)
    # root.wm_attributes("-transparentcolor", "white")

    canvas.create_image(width//2, height//2, image=img)

    def task():
        import json
        import queue
        import os.path
        import pythoncom # only works in Windows
        import threading
        import webbrowser
        import tkinter as tk
        import tkinter.messagebox
        import openpyxl as xl
        import win32com.client # only works in Windows

        from copy import copy
        from tkinter import ttk, filedialog
        from datetime import date as DATE, datetime, timedelta
        from dateutil.relativedelta import relativedelta
        from openpyxl.cell import Cell
        from openpyxl.utils.cell import get_column_letter as gcl
        from openpyxl.worksheet.formula import ArrayFormula
        from openpyxl.worksheet.worksheet import Worksheet
        from openpyxl.worksheet.datavalidation import DataValidation
        from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Protection, Side
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.chrome.service import Service
        from win32process import CREATE_NO_WINDOW # only works in Windows
        from webdriver_manager.chrome import ChromeDriverManager
        root.destroy()

    # root.after(1, task)
    root.mainloop()

start_ui()