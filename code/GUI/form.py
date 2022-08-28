import tkinter as tk
from tkinter import filedialog
from tkinter import *
import customtkinter
import os
from Utility.excel_function import *
import os
from Utility.google_sheet_function import *
import threading 
from Utility.decorator_code import *
PATH = os.path.dirname(os.path.realpath(__file__))

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

def button_function():
    print("button pressed")

def form_page():
    app = customtkinter.CTkToplevel()  # create CTk window like you do with the Tk window (you can also use normal tkinter.Tk window)
    app.geometry("210x220")
    app.title("ສ້າງຟອມ")
    app.resizable(0,0)

    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  # add empty row for spacing

    button_1 = customtkinter.CTkButton(master=frame_1,  text="ສ້າງຟອມ", width=150, height=40,
                                    compound="right", command=lambda:threading.Thread(target=lambda:decorator_loading_popup(create_form)).start())
    button_1.grid(row=1, column=0, columnspan=2, padx=20, pady=10)

    button_2 = customtkinter.CTkButton(master=frame_1, text="ເພີ່ມ Permission", width=150, height=40,
                                    compound="right", fg_color="#D35B58", hover_color="#C77C78",
                                    command=add_permission)
    button_2.grid(row=2, column=0, columnspan=2, padx=20, pady=10)

    button_4 = customtkinter.CTkButton(master=frame_1, text="ກັບຄືນ", width=50, height=50,
                                    corner_radius=10,  command=app.destroy)
    button_4.grid(row=3, column=0, columnspan=2, padx=20, pady=10)

def add_permission():
    
    customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
    customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

    app = customtkinter.CTkToplevel() 

    app.geometry("340x200")
    app.title('ເພີ່ມ Permission')
    app.resizable(0, 0)

    frame_1 = customtkinter.CTkFrame(master=app, width=120, height=170, corner_radius=15)
    frame_1.grid(row=0, column=0, padx=10, pady=10)

    frame_1.grid_columnconfigure(0, weight=0)
    frame_1.grid_columnconfigure(1, weight=0)
    frame_1.grid_rowconfigure(0, minsize=10)  

    # username
    email_label = customtkinter.CTkLabel(frame_1, text="Email:")
    email_label.grid(column=0, row=0, sticky="w", padx=10, pady=10)

    email_entry = customtkinter.CTkEntry(frame_1,placeholder_text="")
    email_entry.grid(column=1, row=0, sticky="e", padx=10, pady=10)

    room_label = customtkinter.CTkLabel(frame_1, text="ຫ້ອງຮຽນ:")
    room_label.grid(column=0, row=1, sticky="w", padx=10, pady=10)

    def display_selected(choice):
        print(choice)
        data = pd.read_excel("Subject.xlsx",choice)
        subject.configure(values=data.to_dict(orient="list")['Subject'])

    # room_entry = customtkinter.CTkEntry(frame_1,placeholder_text="")
    # room_entry.grid(column=1, row=1, sticky="e", padx=10, pady=10)
    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    variable = StringVar(frame_1)
    variable.set(matched_room[0])
    drop = customtkinter.CTkOptionMenu(frame_1,variable=variable,values=matched_room,command=display_selected)
    drop.grid(column=1, row=1, sticky="e", padx=10, pady=10)

    subject_label = customtkinter.CTkLabel(frame_1, text="ວິຊາ:")
    subject_label.grid(column=0, row=2, sticky="w", padx=10, pady=10)

    subject_list=[]
    variable_subject = StringVar(frame_1)
    variable_subject.set("")
    subject = customtkinter.CTkOptionMenu(frame_1,variable=variable_subject,values=subject_list)
    subject.grid(column=1, row=2, sticky="e", padx=10, pady=10)
    # subject.configure(width = 100)

    # subject_entry = customtkinter.CTkEntry(frame_1,placeholder_text="")
    # subject_entry.grid(column=1, row=2, sticky="e", padx=10, pady=10)

    ok_button = customtkinter.CTkButton(frame_1, text="ເພີ່ມ",command=lambda:threading.Thread(target=lambda:decorator_loading_popup(add_permission_function,email=email_entry.get(),subject=subject.get(),room=drop.get())).start())
    ok_button.grid(column=1,row=3,padx=5, pady=5)

def UploadAction(event=None):
        global file_path
        filename = filedialog.askopenfilename()
        file_path = filename

def create_form():
    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    # print("ຫ້ອງ")
    # print_menu_list(matched_room)
    for i in matched_room:
            data = read_file("Subject.xlsx",i)
            for j in range(0,len(data)):
                if str(data.loc[j,"Google sheet link"])=="nan":
                    spreadsheet_id,link,sheet_id = create_sheet_from_teacher(i,data["Lecture by"][j])
                    data.loc[j,"Google sheet link"]=link
                    data.loc[j,"Spreadsheet ID"]=spreadsheet_id
                    data.loc[j,"Sheet ID"]=sheet_id
                    grant_permission(data["email"][j],data.loc[j,"Spreadsheet ID"])
                    create_sheet_form(i,data.loc[j,"Spreadsheet ID"],data["Subject"][j],data["Lecture by"][j])
                    print(int(data.loc[j,"Sheet ID"]))
                    print(data.loc[j,"Spreadsheet ID"])
                    auto_adjust(int(data.loc[j,"Sheet ID"]),data.loc[j,"Spreadsheet ID"])
            write_file("Subject.xlsx",i,data)
def add_permission_function(email:str,subject:str,room:str):
    df_subject = read_file("Subject.xlsx",room)
    for subject_seq in range(0,len(df_subject)):
        if(df_subject.loc[subject_seq,"Subject"]==subject):
            grant_permission(email,df_subject.loc[subject_seq,"Spreadsheet ID"])
