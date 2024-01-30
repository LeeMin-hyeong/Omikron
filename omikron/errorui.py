import sys
import tkinter as tk
from omikron.defs import VERSION

def no_config_file_error():
    ui = tk.Tk()

    width = 300
    height = 120
    x = int((ui.winfo_screenwidth()/2) - (width/2))
    y = int((ui.winfo_screenheight()/2) - (height/2))
    ui.geometry(f"{width}x{height}+{x}+{y}")

    ui.title(VERSION)
    ui.resizable(False, False)

    tk.Label(ui).pack()
    tk.Label(ui, text=r"'config.json' 파일이 없어 실행할 수 없습니다.").pack()
    tk.Label(ui, text=r"학원 데스크에 문의해 주세요.").pack()
    tk.Label(ui).pack()

    button = tk.Button(ui, cursor="hand2", text="확인", width=15, command=sys.exit)
    button.pack()

    ui.mainloop()

def corrupted_config_file_error():
    ui = tk.Tk()

    width = 300
    height = 120
    x = int((ui.winfo_screenwidth()/2) - (width/2))
    y = int((ui.winfo_screenheight()/2) - (height/2))
    ui.geometry(f"{width}x{height}+{x}+{y}")

    ui.title(VERSION)
    ui.resizable(False, False)

    tk.Label(ui).pack()
    tk.Label(ui, text=r"'config.json' 파일이 손상되어 실행할 수 없습니다.").pack()
    tk.Label(ui, text=r"학원 데스크에 문의해 주세요.").pack()
    tk.Label(ui).pack()

    button = tk.Button(ui, cursor="hand2", text="확인", width=15, command=sys.exit)
    button.pack()

    ui.mainloop()

def chrome_driver_version_error():
    ui = tk.Tk()

    width = 300
    height = 140
    x = int((ui.winfo_screenwidth()/2) - (width/2))
    y = int((ui.winfo_screenheight()/2) - (height/2))
    ui.geometry(f"{width}x{height}+{x}+{y}")

    ui.title(VERSION)
    ui.resizable(False, False)

    tk.Label(ui).pack()
    tk.Label(ui, text=r"'ChromeDriver' 업데이트가 필요합니다.").pack()
    tk.Label(ui, text=r"'Omikron_installer.bat'을 실행하여").pack()
    tk.Label(ui, text=r"업데이트를 진행할 수 있습니다.").pack()
    tk.Label(ui).pack()

    button = tk.Button(ui, cursor="hand2", text="확인", width=15, command=sys.exit)
    button.pack()

    ui.mainloop()
