import tkinter as tk
from tkinter import filedialog
from tkinter import *
import customtkinter
from PIL import Image, ImageTk  # <- import PIL for the images
import os
from Utility.excel_function import *

def add_data_menu():

    PATH = os.path.dirname(os.path.realpath(__file__))

    customtkinter.set_appearance_mode("System") 
    customtkinter.set_default_color_theme("blue")  
    app = customtkinter.CTkToplevel()
    app.geometry("220x220")
    app.title("ເພີ່ມຂ້ໍມູນ")
    app.resizable(0,0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10) 

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ເພີ່ມຂໍ້ມູນນັກສຶກສາ", width=150, height=40,
                                    compound="right", command=add_data_student)
    button_1.grid(row=1, column=0, columnspan=2, padx=20, pady=10)

    button_2 = customtkinter.CTkButton(master=frame_1, text="ເພີ່ມຂໍ້ມູນອາຈານແລະລາຍວິຊາ", width=150, height=40,
                                    compound="right", fg_color="#D35B58", hover_color="#C77C78",
                                    command=add_data_teacher)
    button_2.grid(row=2, column=0, columnspan=2, padx=20, pady=10)

    button_4 = customtkinter.CTkButton(master=frame_1, text="ກັບຄືນ", width=50, height=50,
                                    corner_radius=10,  command=app.destroy)
    button_4.grid(row=3, column=0, columnspan=2, padx=20, pady=10)

def add_data_student():
    app = customtkinter.CTkToplevel() 

    app.geometry("470x150")
    app.title('ເພີ່ມຂໍ້ມູນນັກຮຽນ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    def UploadAction(event=None):
        global file_path
        filename = filedialog.askopenfilename()
        file_path = filename
        label1['text']=filename

    file_label = customtkinter.CTkLabel(frame_1, text="Open:")
    file_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)


    sheetname_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງຮຽນ:")
    sheetname_label.grid(column=0, row=1, sticky="w", padx=5, pady=10)

    sheetname_entry = customtkinter.CTkEntry(frame_1,)
    sheetname_entry.grid(column=1, row=1, sticky="e", padx=5, pady=5)

    label1 = tk.Label(frame_1,text='Please choose a file',width=20)
    label1.grid(column=1, row=0, padx=5, pady=5)

    open_button = customtkinter.CTkButton(frame_1, text="ເປີດຟາຍ",command=UploadAction)
    open_button.grid(column=2, row=0, sticky="e", padx=5, pady=5)
    ok_button = customtkinter.CTkButton(frame_1, text="ເພີ່ມ",command=lambda:add_student_data(file_path,sheetname_entry.get().upper(),app))
    ok_button.grid(column=1,row=2,padx=5, pady=5)

def add_data_teacher():
    app = customtkinter.CTkToplevel() 

    app.geometry("470x150")
    app.title('ເພີ່ມຂໍ້ມູນລາຍວິຊາແລະອາຈານ')
    app.resizable(0, 0)
    app.grab_set()
    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    def UploadAction(event=None):
        global file_path
        filename = filedialog.askopenfilename()
        file_path = filename
        label1['text']=filename

    file_label = customtkinter.CTkLabel(frame_1, text="Open:")
    file_label.grid(column=0, row=0, sticky="w", padx=5, pady=10)


    sheetname_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງຮຽນ:")
    sheetname_label.grid(column=0, row=1, sticky="w", padx=5, pady=10)

    # sheetname_entry = customtkinter.CTkEntry(frame_1)
    # sheetname_entry.grid(column=1, row=1, sticky="e", padx=5, pady=5)
    namefile = excel_info("Score.xlsx").sheet_names
    matched_room = list(set(namefile))
    variable = StringVar(frame_1)
    variable.set("")
    sheetname_entry = customtkinter.CTkOptionMenu(frame_1,variable=variable,values=matched_room)
    sheetname_entry.grid(column=1, row=1, sticky="e", padx=5, pady=5)

    label1 = tk.Label(frame_1,text='Please choose a file',width=20)
    label1.grid(column=1, row=0, padx=5, pady=5)

    open_button = customtkinter.CTkButton(frame_1, text="ເປີດຟາຍ",command=UploadAction)
    open_button.grid(column=2, row=0, sticky="e", padx=5, pady=5)

    ok_button = customtkinter.CTkButton(frame_1, text="ເພີ່ມ",command=lambda:add_teacher_data(file_path,variable.get(),app))
    ok_button.grid(column=1,row=2,padx=5, pady=5)

def add_student_data(filename:str,sheetname:str,app):
    if sheetname != "":
        if os.name == 'posix':
            filename = filename.replace('\\', ' ').strip()
        filename = filename.replace("'","")
        filename = filename.replace('"','')
        df = read_name_file(filename)
        write_file("Score.xlsx",sheetname,df)
        # print("ສຳເລັດແລ້ວ")
        app.destroy()
        success_form()
        
def add_teacher_data(filename:str,sheetname:str,app):
    if sheetname != "":
        if os.name == 'posix':
            filename = filename.replace('\\', ' ').strip()
        filename = filename.replace("'","")
        filename = filename.replace('"','')
        df = read_file(filename)
        df_name = read_file("Score.xlsx",sheetname)
        subject_name = df["Subject"].tolist()
        add_subject_to_name_file("Score.xlsx",sheetname,df_name,subject_name)
        add_subject_to_teacher_file("Subject.xlsx",sheetname,df)
    # print("ສຳເລັດແລ້ວ")
    app.destroy()
    success_form()
def success_form():
    customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
    customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

    app = customtkinter.CTkToplevel() 

    app.geometry("170x90")
    app.title('ສຳເລັດ')
    app.resizable(0, 0)

    frame_1 = customtkinter.CTkFrame(master=app, width=150, height=200, corner_radius=5)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    label1 = tk.Label(frame_1,text='ເພິ່ມຂໍ້ມູນ',width=20)
    label1.grid(column=1, row=0, padx=5, pady=5)

    open_button = customtkinter.CTkButton(frame_1, text="ກັບຄືນ",command=app.destroy)
    open_button.grid(column=1, row=1, padx=5, pady=5)