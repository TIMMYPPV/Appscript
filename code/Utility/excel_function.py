import pandas as pd
import os 
from os.path import exists
# import openpyxl



def read_name_file(filename:str):
    data = pd.read_excel(filename,skiprows=15,usecols="A:C").dropna()
    return data

def read_subject_file(filename:str,sheetname:str):
    if(sheetname==""):
        data = pd.read_excel(filename)
        return data
    else:
        data = pd.read_excel(filename,sheetname)
        return data

def read_file(filename:str,sheetname=""):
    if(sheetname==""):
        data = pd.read_excel(filename)
        return data
    else:
        data = pd.read_excel(filename,sheetname)
        return data

def add_subject_to_teacher_file(filename:str,sheetname:str,df):
    if(exists('Subject.xlsx')):
        sheetnames = show_sheetnames(filename)
        # print(sheetname in sheetnames)
        if(sheetname in sheetnames):
            # print(True)
            data = read_file("Subject.xlsx",sheetname)
            df = pd.concat([data,df],ignore_index=True).drop_duplicates(subset=['Subject'])
        else:
            df["Google sheet link"]=""
            df["Spreadsheet ID"]=""
            df["Status"]=""
            df["Sheet ID"]=""
            df["Status"]=""
    else:
        df["Google sheet link"]=""
        df["Spreadsheet ID"]=""
        df["Status"]=""
        df["Sheet ID"]=""
        df["Status"]=""
    write_file(filename,sheetname,df)

def write_file(filename:str,sheetname:str,df):
    if os.path.exists(filename):
        writer = pd.ExcelWriter(filename, mode="a", engine="openpyxl",if_sheet_exists="replace")
        df.to_excel(writer, sheet_name=sheetname,index=False)
        writer.save()
    else:
        df.to_excel(filename,sheet_name=sheetname,index=False)

def add_subject_to_name_file(filename:str,sheetname:str,df_name,subject_name):
    # print(df_name)
    for i in subject_name:
        if (i not in df_name.columns):
            df_name[i]=""
    # print(df_name)
    write_file(filename,sheetname,df_name)

def adjust_cell(filename:str,sheetname:str):
    df = pd.read_excel(filename,sheetname)
    writer = pd.ExcelWriter(filename) 
    df.to_excel(writer, sheet_name=sheetname, index=False, na_rep='')

    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheetname].set_column(col_idx, col_idx, column_length)

    writer.close()

def excel_info(filename:str):
    data = pd.ExcelFile(filename)
    return data

def print_menu_list(menu):
    for i in range(0,len(menu)):
        print(str(i+1)+"."+menu[i])
        
def check_colname(df,colname:str):
    return(True) if(colname in df) else False

def show_sheetnames(filename:str):
    return(pd.ExcelFile(filename).sheet_names)

# def add_score_to_subject(subject_name:str,df_score,df_subject):

