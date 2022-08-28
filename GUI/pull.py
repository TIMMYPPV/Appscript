from Utility.excel_function import *
from Utility.google_sheet_function import *
import pandas as pd
import numpy

def do_pull():
    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    for room in matched_room:
        df_score = read_file("Score.xlsx",room)
        df_subject = read_file("Subject.xlsx",room)
        for subject_seq in range(0,len(df_subject)):
            # print(subject_seq)
            # print(df_subject.loc[subject_seq,"Spreadsheet ID"])
            # print(len(df_subject.loc[subject_seq,"Spreadsheet ID"]))
            if type(df_subject.loc[subject_seq,"Spreadsheet ID"])==numpy.float64:
                continue
            if df_subject.loc[subject_seq,"Status"]=="Done":
                continue
            score_data = pull_score(df_subject.loc[subject_seq,"Spreadsheet ID"],room)
            room_num,subject_name = pull_info(df_subject.loc[subject_seq,"Spreadsheet ID"],room)
            # print(subject_name)
            df_subject.loc[subject_seq,'Status']="Done"
            for person in range (0,len(score_data)):
                if(score_data.loc[person,"ລະຫັດນັກສຶກສາ"]==df_score.loc[person,"ລະຫັດນັກສຶກສາ"]):
                    df_score.loc[person,subject_name]=score_data.loc[person,"Grade"]
                else:
                    df_subject.loc[subject_seq,'Status']=""
            # print(score_data)
        write_file("Subject.xlsx",room,df_subject)
        write_file("Score.xlsx",room,df_score)