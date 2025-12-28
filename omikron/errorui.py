import sys
import tkinter as tk

def no_config_file_error():
    """
    실행 파일과 같은 위치에 `config.json`파일이 존재하지 않을 경우
    """
    ui = tk.Tk()

    width  = 300
    height = 120
    x = int((ui.winfo_screenwidth()/2) - (width/2))
    y = int((ui.winfo_screenheight()/2) - (height/2))
    ui.geometry(f"{width}x{height}+{x}+{y}")

    ui.title("Omikron")
    ui.resizable(False, False)

    tk.Label(ui).pack()
    tk.Label(ui, text=r"'config.json' 파일이 없어 실행할 수 없습니다.").pack()
    tk.Label(ui, text=r"학원 데스크에 문의해 주세요.").pack()
    tk.Label(ui).pack()

    button = tk.Button(ui, cursor="hand2", text="확인", width=15, command=sys.exit)
    button.pack()

    ui.mainloop()

def corrupted_config_file_error():
    """
    `config.json` 파일의 일부 데이터가 손상되었을 경우
    """
    ui = tk.Tk()

    width  = 300
    height = 120
    x = int((ui.winfo_screenwidth()/2) - (width/2))
    y = int((ui.winfo_screenheight()/2) - (height/2))
    ui.geometry(f"{width}x{height}+{x}+{y}")

    ui.title("Omikron")
    ui.resizable(False, False)

    tk.Label(ui).pack()
    tk.Label(ui, text=r"'config.json' 파일이 손상되어 실행할 수 없습니다.").pack()
    tk.Label(ui, text=r"학원 데스크에 문의해 주세요.").pack()
    tk.Label(ui).pack()

    button = tk.Button(ui, cursor="hand2", text="확인", width=15, command=sys.exit)
    button.pack()

    ui.mainloop()
