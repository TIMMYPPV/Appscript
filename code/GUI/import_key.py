from tkinter import filedialog 
import tkinter as tk
import customtkinter
from PIL import Image, ImageTk  # <- import PIL for the images
import os
from tkinter.messagebox import showinfo
import shutil,os

def import_form():

    PATH = os.path.dirname(os.path.realpath(__file__))

    customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
    customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
    app = customtkinter.CTkToplevel() 

    app.geometry("470x110")
    app.title('ເພີ່ມkey')
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


    label1 = tk.Label(frame_1,text='Please choose a file',width=10)
    label1.grid(column=1, row=0, padx=5, pady=5)

    open_button = customtkinter.CTkButton(frame_1, text="ເປີດຟາຍ",command=UploadAction)
    open_button.grid(column=2, row=0, sticky="e", padx=5, pady=5)
    ok_button = customtkinter.CTkButton(frame_1, text="ເພີ່ມ",command=lambda:import_key(label1['text'],app))
    ok_button.grid(column=1,row=2,padx=5, pady=5)
def import_key(filepath,app):
    if(filepath!=""):
        shutil.copy2(filepath,os.getcwd())
        app.destroy()
