import customtkinter
import tkinter as tk
from tkinter import ttk
import sys, os
import threading

header = {
    "report_score_by_room_by_subject":"ລາຍງານເກຣດຕາມວິຊາ",
    "report_score":"ລາຍງານເກຣດຕາມຫ້ອງ",
    "statistic_report_by_subject_by_room":"ລາຍງານສະຖິຕິເກຣດ",
    "report_status":"ລາຍງານສະຖານະການສົ່ງເກຣດ",
    "report_statistic_by_room":"",
    "report_by_major":""
}
message={
    "report_score_by_room_by_subject":"",
    "report_score":"",
    "statistic_report_by_subject_by_room":"",
    "report_status":"",
    "report_statistic_by_room":"",
    "report_by_major":""
}


def decorator_loading_popup(func,**kwargs):
    mainfrm = tk.Toplevel()
    mainfrm.title("ກຳລັງດຳເນີນການ")
    pgb1 = ttk.Progressbar(mainfrm,mode="indeterminate",length=200)
    pgb1.pack(padx=10,pady=10)
    p=threading.Thread(target=lambda:pgb1.start())
    p.start()
    try:
        q=threading.Thread(target=lambda:func(**kwargs))
        q.start()
        q.join()
        success_form()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(e)
        error_form()
    mainfrm.destroy()

def decorator_popup(func,**kwargs):
    try:
        func(**kwargs)
        success_form()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(e)
        error_form()



def success_form():
    customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
    customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

    app = customtkinter.CTkToplevel() 

    app.geometry("210x90")
    app.title('ສຳເລັດ')
    app.resizable(0, 0)

    frame_1 = customtkinter.CTkFrame(master=app, width=150, height=200, corner_radius=5)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    label1 = tk.Label(frame_1,text='ດຳເນີນການສຳເລັດ',width=20)
    label1.grid(column=1, row=0, padx=5, pady=5)

    open_button = customtkinter.CTkButton(frame_1, text="ກັບຄືນ",command=app.destroy)
    open_button.grid(column=1, row=1, padx=5, pady=5)

def error_form():
    customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
    customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

    app = customtkinter.CTkToplevel() 

    app.geometry("210x90")
    app.title("ຜິດພາດ")
    app.resizable(0, 0)

    frame_1 = customtkinter.CTkFrame(master=app, width=150, height=200, corner_radius=5)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    label1 = tk.Label(frame_1,text="ດຳເນີນການຜິດພາດ",width=20)
    label1.grid(column=1, row=0, padx=5, pady=5)

    open_button = customtkinter.CTkButton(frame_1, text="ກັບຄືນ",command=app.destroy)
    open_button.grid(column=1, row=1, padx=5, pady=5)
