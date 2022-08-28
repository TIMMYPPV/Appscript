import tkinter
import customtkinter
import threading
from PIL import Image, ImageTk  # <- import PIL for the images
import os
from GUI.add_data import *
from GUI.form import *
from GUI.report import *
from GUI.import_key import *
from GUI.pull import *
from os.path import abspath, dirname

os.chdir(dirname(abspath(__file__)))
# print(os.getcwd())
# print(os.path.realpath(__file__))

PATH = os.path.dirname(os.path.realpath(__file__))

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

app = customtkinter.CTk()  # create CTk window like you do with the Tk window (you can also use normal tkinter.Tk window)
app.geometry("400x320")
app.title("CEIT Appscript")
app.resizable(0,0)



image_size = 20

settings_image = ImageTk.PhotoImage(Image.open(PATH + "/images/settings.png").resize((image_size, image_size)))
bell_image = ImageTk.PhotoImage(Image.open(PATH + "/images/bell.png").resize((image_size, image_size)))
folder_image = ImageTk.PhotoImage(Image.open(PATH + "/images/add-folder.png").resize((image_size, image_size)))
list_image = ImageTk.PhotoImage(Image.open(PATH + "/images/add-list.png").resize((image_size, image_size)))
user_image = ImageTk.PhotoImage(Image.open(PATH + "/images/add-user.png").resize((image_size, image_size)))
chat_image = ImageTk.PhotoImage(Image.open(PATH + "/images/chat.png").resize((image_size, image_size)))
home_image = ImageTk.PhotoImage(Image.open(PATH + "/images/home.png").resize((image_size, image_size)))

app.grid_rowconfigure(0, weight=1)
app.grid_columnconfigure(0, weight=1)

frame_1 = customtkinter.CTkFrame(master=app, width=120, height=200, corner_radius=15)
frame_1.grid(row=0, column=0, padx=10, pady=10)

frame_1.grid_columnconfigure(0, weight=0)
frame_1.grid_columnconfigure(1, weight=0)
frame_1.grid_rowconfigure(0, minsize=10)  # add empty row for spacing

button_1 = customtkinter.CTkButton(master=frame_1, image=folder_image, text="ເພີ່ມຂໍ້ມູນ", width=150, height=40,
                                   compound="right", command=add_data_menu)
button_1.grid(row=1, column=0, columnspan=2, padx=20, pady=10)

button_2 = customtkinter.CTkButton(master=frame_1, image=list_image, text="ສ້າງຟອມ", width=150, height=40,
                                   compound="right", 
                                   command=form_page)
button_2.grid(row=2, column=0, columnspan=2, padx=20, pady=10)

button_4 = customtkinter.CTkButton(master=frame_1, image=home_image, text="ສ້າງReport", width=150, height=40,compound="right", 
                                   corner_radius=10,  command=report_form)
button_4.grid(row=3, column=0, columnspan=2, padx=20, pady=10)
button_6 = customtkinter.CTkButton(master=frame_1,  text="Import Key", width=150, height=40,fg_color="#D35B58", hover_color="#C77C78",
                                   corner_radius=10,  command=import_form)
button_6.grid(row=4, column=0, columnspan=2, padx=20, pady=10)

button_5 = customtkinter.CTkButton(master=app,  text="ດຶງຂໍ້ມູນ", width=100, height=70,
                                   corner_radius=10, compound="bottom",  fg_color=("gray84", "gray25"), hover_color="#C77C78",
                                   command=lambda:threading.Thread(target=lambda:decorator_loading_popup(do_pull)).start())
button_5.grid(row=0, column=1, padx=20, pady=20)

app.mainloop()
