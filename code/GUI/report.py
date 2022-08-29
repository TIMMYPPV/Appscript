import tkinter as tk
from tkinter import *
import customtkinter
from PIL import Image, ImageTk  
import os

from numpy import mat
from Utility.excel_function import *
from Utility.report_function import *
from Utility.decorator_code import *

def report_form():
    PATH = os.path.dirname(os.path.realpath(__file__))

    customtkinter.set_appearance_mode("System")  
    customtkinter.set_default_color_theme("blue")  
    app = customtkinter.CTkToplevel() 
    app.geometry("250x480")
    app.title("ລາຍງານ")
    app.resizable(0,0)

    app.grid_rowconfigure(0, weight=1)
    app.grid_columnconfigure(0, weight=1, minsize=200)
    app.grab_set()

    frame_1 = customtkinter.CTkFrame(master=app, width=250, height=240, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

    frame_1.grid_columnconfigure(0, weight=1)
    frame_1.grid_columnconfigure(1, weight=1)
    frame_1.grid_rowconfigure(0, minsize=10)  # add empty row for spacing

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ລາຍງານເກຣດຕາມຫ້ອງ", width=190, height=40,
                                    compound="right", command=report_menu_grade_by_room)
    button_1.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ລາຍງານເກຣດຕາມວິຊາ", width=190, height=40,
                                    compound="right", command=report_menu_grade_by_subject)
    button_1.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ລາຍງານສະຖິຕິເກຣດນັກສຶກສາ", width=190, height=40,
                                    compound="right", command=report_menu_grade_statistic)
    button_1.grid(row=3, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ລາຍງານສະຖານະການສົ່ງເກຣດ", width=190, height=40,
                                    compound="right", command=report_menu_status_grade)
    button_1.grid(row=4, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ລາຍງານສະຖິຕິເກຣດຕາມຫ້ອງ", width=190, height=40,
                                    compound="right", command=report_menu_statistic_grade_by_room)
    button_1.grid(row=5, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ລາຍງານສະຖິຕິເກຣດຕາມຊັ້ນຮຽນ", width=190, height=40,
                                    compound="right", command=report_menu_by_major)
    button_1.grid(row=6, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    button_2 = customtkinter.CTkButton(master=frame_1, text="ກັບຄືນ", width=190, height=40,
                                    compound="right", fg_color="#D35B58", hover_color="#C77C78",
                                    command=lambda:app.destroy())
    button_2.grid(row=7, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

def report_menu_grade_by_room():
    app = customtkinter.CTkToplevel() 

    app.geometry("320x100")
    app.title('ລາຍງານເກຣດຕາມຫ້ອງ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  
    sheetname_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງຮຽນ:")
    sheetname_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)

    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    variable = StringVar(frame_1)
    variable.set(matched_room[0])
    drop = customtkinter.CTkOptionMenu(frame_1,variable=variable,values=matched_room)
    drop.grid(column=1, row=0, sticky="ew", padx=5, pady=5)

    ok_button = customtkinter.CTkButton(frame_1, text="ສ້າງລາຍງານ",command=lambda:decorator_popup(report_score,room=drop.get()))
    ok_button.grid(column=1,row=1,padx=5, pady=5)

def report_menu_grade_by_subject():
    app = customtkinter.CTkToplevel() 

    app.geometry("320x150")
    app.title('ລາຍງານເກຣດຕາມວິຊາ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    file_label = customtkinter.CTkLabel(frame_1, text="ວິຊາ:")
    file_label.grid(column=0, row=1, sticky="w", padx=5, pady=10)

    sheetname_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງຮຽນ:")
    sheetname_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)

    subject_list=[]
    variable_subject = StringVar(frame_1)
    variable_subject.set("")
    subject = customtkinter.CTkOptionMenu(frame_1,variable=variable_subject,values=subject_list)
    subject.grid(column=1, row=1, sticky="e", padx=5, pady=5)

    def display_selected(choice):
        print(choice)
        data = pd.read_excel("Subject.xlsx",choice)
        subject.configure(values=data.to_dict(orient="list")['Subject'])

    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    variable = StringVar(frame_1)
    variable.set("")
    drop = customtkinter.CTkOptionMenu(frame_1,variable=variable,values=matched_room,command=display_selected)
    drop.grid(column=1, row=0, sticky="e", padx=5, pady=5)

    ok_button = customtkinter.CTkButton(frame_1, text="ສ້າງລາຍງານ",command=lambda:decorator_popup(report_score_by_room_by_subject,sheetname=drop.get(),subject=subject.get()))
    ok_button.grid(column=1,row=2,padx=5, pady=5)

def report_menu_grade_statistic():
    app = customtkinter.CTkToplevel() 

    app.geometry("320x150")
    app.title('ລາຍງານສະຖິຕິເກຣດ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=50)  

    sheetname_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງຮຽນ:")
    sheetname_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)

    file_label = customtkinter.CTkLabel(frame_1, text="ວິຊາ:")
    file_label.grid(column=0, row=1, sticky="we", padx=5, pady=10)
    
    subject_list=[]
    variable_subject = StringVar(frame_1)
    variable_subject.set("")
    subject = customtkinter.CTkOptionMenu(frame_1,variable=variable_subject,values=subject_list)
    subject.grid(column=1, row=1,  padx=5, pady=5 ,sticky="we")


    def display_selected(choice):
        print(choice)
        data = pd.read_excel("Subject.xlsx",choice)
        subject.configure(values=data.to_dict(orient="list")['Subject'])
        # print(data.to_dict(orient="list")['Subject'])

    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    variable_room = StringVar(frame_1)
    variable_room.set("")
    sheetname = customtkinter.CTkOptionMenu(frame_1,variable=variable_room,values=matched_room,command=display_selected)
    sheetname.grid(column=1, row=0, sticky="w", padx=5, pady=5)

    ok_button = customtkinter.CTkButton(frame_1, text="ສ້າງລາຍງານ",command=lambda:decorator_popup(statistic_report_by_subject_by_room,room=sheetname.get(),subject=subject.get()))
    ok_button.grid(column=1,row=2,padx=5, pady=5,sticky="w")

def report_menu_status_grade():
    app = customtkinter.CTkToplevel() 

    app.geometry("320x100")
    app.title('ລາຍງານສະຖານະການສົ່ງເກຣດ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    sheetname_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງຮຽນ:")
    sheetname_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)

    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    variable = StringVar(frame_1)
    variable.set(matched_room[0])
    sheetname = customtkinter.CTkOptionMenu(frame_1,variable=variable,values=matched_room)
    sheetname.grid(column=1, row=0, sticky="e", padx=5, pady=5)

    ok_button = customtkinter.CTkButton(frame_1, text="ສ້າງລາຍງານ",command=lambda:decorator_popup(report_status,room=sheetname.get()))
    ok_button.grid(column=1,row=1,padx=5, pady=5)

def report_menu_statistic_grade_by_room():
    app = customtkinter.CTkToplevel() 

    app.geometry("320x150")
    app.title('ລາຍງານສະຖິຕິເກຣດຕາມຫ້ອງ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    file_label = customtkinter.CTkLabel(frame_1, text="ເທີມທີ່:")
    file_label.grid(column=0, row=1, sticky="w", padx=5, pady=10)

    sheetname_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງ:")
    sheetname_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)

    subject_list=[]
    variable_subject = StringVar(frame_1)
    variable_subject.set("")
    subject = customtkinter.CTkOptionMenu(frame_1,variable=variable_subject,values=subject_list)
    subject.grid(column=1, row=1, sticky="e", padx=5, pady=5)

    def display_selected(choice):
        data = pd.read_excel("Subject.xlsx",choice)
        term_list = []
        [term_list.append(str(x)) for x in data.to_dict(orient="list")["term"] if str(x) not in term_list]
        subject.configure(values=term_list)

    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    variable = StringVar(frame_1)
    variable.set("")
    drop = customtkinter.CTkOptionMenu(frame_1,variable=variable,values=matched_room,command=display_selected)
    drop.grid(column=1, row=0, sticky="e", padx=5, pady=5)

    ok_button = customtkinter.CTkButton(frame_1, text="ສ້າງລາຍງານ",command=lambda:decorator_popup(report_statistic_by_room,room=drop.get(),term=subject.get()))
    ok_button.grid(column=1,row=2,padx=5, pady=5)
def report_menu_by_major():
    app = customtkinter.CTkToplevel() 

    app.geometry("320x150")
    app.title('ລາຍງານສະຖິຕິເກຣດຕາມສາຍຮຽນ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    file_label = customtkinter.CTkLabel(frame_1, text="ປີ:")
    file_label.grid(column=0, row=1, sticky="w", padx=5, pady=10)

    sheetname_label = customtkinter.CTkLabel(frame_1, text="ສາຍຮຽນ:")
    sheetname_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)

    subject_list=[]
    variable_subject = StringVar(frame_1)
    variable_subject.set("")
    year = customtkinter.CTkOptionMenu(frame_1,variable=variable_subject,values=subject_list)
    year.grid(column=1, row=1, sticky="e", padx=5, pady=5)

    def display_selected(choice):
        namefile = excel_info("Score.xlsx").sheet_names
        subjectfile = excel_info("Subject.xlsx").sheet_names
        matched_room = list(set(namefile)&set(subjectfile))
        term_list =[x[:1] for x in matched_room if choice in x]
        year.configure(values=term_list)

    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    major=[]
    [major.append(str(x)[1:-1]) for x in matched_room if str(x)[1:-1] not in major]
    variable = StringVar(frame_1)
    variable.set("")
    major = customtkinter.CTkOptionMenu(frame_1,variable=variable,values=major,command=display_selected)
    major.grid(column=1, row=0, sticky="e", padx=5, pady=5)

    ok_button = customtkinter.CTkButton(frame_1, text="ສ້າງລາຍງານ",command=lambda:decorator_popup(report_by_major,major=major.get(),year=year.get()))
    ok_button.grid(column=1,row=2,padx=5, pady=5)
