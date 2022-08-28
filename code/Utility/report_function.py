from re import T
from fpdf import FPDF
from matplotlib import pyplot as plt
import numpy as np
from pathlib import Path
import pandas as pd
from Utility.excel_function import *

def report_score(room):
    score = pd.read_excel("Score.xlsx",room)
    subject = pd.read_excel("Subject.xlsx",room)
    subject = subject[["Subject","credits"]].to_records()
    data = {"ລ/ດ":"","ລະຫັດນັກສຶກສາ":"","ຊື່ ແລະ ນາມສະກຸນ":""}
    for i in subject:
        data[i[1]]=i[2]
    row = pd.DataFrame(data,index=[0])
    score = pd.concat([row,score]).reset_index(drop=True).fillna("")
    score.to_excel(str(Path.home())+'/Desktop/report_score.xlsx',room,index=False)

def statistic_report_by_subject_by_room(room:str,subject:str):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margin(20)
    pdf.add_font('Saysettha', fname='saysettha_ot.ttf')
    pdf.add_font('saysetthaB',fname='Saysettha-Bold.ttf')
    pdf.set_font('Saysettha', size=12)
    pdf.image('logo.png',90,8,25)

    pdf.ln(15)
    pdf.cell(0,0,"ສາທາລະນະລັດ ປະຊາທິປະໄຕ ປະຊາຊົນລາວ",align="C")
    pdf.ln(5)
    pdf.cell(0,0,"ສັນຕິພາບ ເອກະລາດ ປະຊາທິປະໄຕ ເອກະພາບ ວັດທະນາຖາວອນ",align="C")
    pdf.ln(15)
    pdf.cell(0,0,"ມະຫາວິທະຍາໄລແຫ່ງຊາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ຄະນະວິສະວະກຳສາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ພາກວິຊາວິສະວະກຳຄອມພິວເຕີ ແລະ ເຕັກໂນໂລຊີຂໍ້ມູນຂ່າວສານ")
    pdf.ln(10)
    pdf.set_font('SaysetthaB', size=14)
    pdf.cell(0,0,"ໃບລາຍງານສະຖິຕິເກຣດນັກສຶກສາ",align="C")
    pdf.set_font('Saysettha', size=12)
    pdf.ln(10)
    data_name = pd.read_excel("Subject.xlsx",room)
    teacher_name = ""
    for i in range(0,len(data_name)):
        if (data_name.loc[i,"Subject"]==subject):
            teacher_name=data_name.loc[i,"Lecture by"]
    pdf.cell(20,0,"ຊື່ອາຈານ : ",align="L")
    pdf.cell(0,0,teacher_name,align="L")
    pdf.ln(5)
    pdf.cell(20,0,"ຫ້ອງຮຽນ : ",align="L")
    pdf.cell(0,0,room,align="L")
    pdf.ln(5)
    pdf.cell(20,0,"ວິຊາ : ",align="L")
    pdf.cell(0,0,subject,align="L")
    pdf.ln(10)

    grades_index = ['A', 'B+', 'B','C+', 'C', 'D+','D','F','W',"I"]

    data = [0,0,0,0,0,0,0,0,0,0]
    col=["ເກຣດ","ຈຳນວນ"]
    student_data = pd.read_excel("Score.xlsx",room).reset_index(drop=True)
    student_data = student_data.fillna("")
    student_data = student_data[["ລ/ດ","ລະຫັດນັກສຶກສາ","ຊື່ ແລະ ນາມສະກຸນ",subject]].to_records(index=False)
    for i in student_data:
        try:
            data[grades_index.index(i[3])]=data[grades_index.index(i[3])]+1
        except:
            continue

    data2 =[]
    for i in range(0,len(grades_index)):
        data2.append([grades_index[i],str(data[i])])
    pdf.set_font("Saysettha", size=10)
    line_height = pdf.font_size * 2
    col_width = pdf.epw / 2 

    pdf.set_font('SaysetthaB', size=10)
    for col_name in col:
        pdf.cell(col_width, line_height, col_name, border=1)
    pdf.ln(line_height)
    pdf.set_font('Saysettha', size=10)

    for row in data2:
        if pdf.will_page_break(line_height):
            pdf.set_font('SaysetthaB', size=10)
            for col_name in col:
                pdf.cell(col_width, line_height, col_name, border=1)
            pdf.ln(line_height)
            pdf.set_font('Saysettha', size=10)
        for datum in row:
            pdf.cell(col_width, line_height, str(datum), border=1)
        pdf.ln(line_height)


    pdf.set_font('SaysetthaB', size=14)

    pdf.ln(line_height)

    fig = plt.figure()
    plt.bar(grades_index,data)

    fig.savefig("bar.png", dpi=300, bbox_inches='tight', pad_inches=0)
    plt.close()
    pdf.ln(line_height)
    pdf.image("bar.png",w=100,x=60)

    pdf.output(str(Path.home())+'/Desktop/report_statistic.pdf', 'F')

def report_score_by_room_by_subject(sheetname:str,subject:str):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margin(20)
    pdf.add_font('Saysettha', fname='saysettha_ot.ttf')
    pdf.add_font('saysetthaB',fname='Saysettha-Bold.ttf')
    pdf.set_font('Saysettha', size=12)
    pdf.image('logo.png',90,8,25)

    pdf.ln(15)
    pdf.cell(0,0,"ສາທາລະນະລັດ ປະຊາທິປະໄຕ ປະຊາຊົນລາວ",align="C")
    pdf.ln(5)
    pdf.cell(0,0,"ສັນຕິພາບ ເອກະລາດ ປະຊາທິປະໄຕ ເອກະພາບ ວັດທະນາຖາວອນ",align="C")
    pdf.ln(15)
    pdf.cell(0,0,"ມະຫາວິທະຍາໄລແຫ່ງຊາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ຄະນະວິສະວະກຳສາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ພາກວິຊາວິສະວະກຳຄອມພິວເຕີ ແລະ ເຕັກໂນໂລຊີຂໍ້ມູນຂ່າວສານ")
    pdf.ln(10)
    pdf.set_font('SaysetthaB', size=14)
    pdf.cell(0,0,"ໃບລາຍງານເກຣດນັກສຶກສາ",align="C")
    pdf.set_font('Saysettha', size=12)
    pdf.ln(10)
    data_name = pd.read_excel("Subject.xlsx",sheetname)
    teacher_name = ""
    for i in range(0,len(data_name)):
        if (data_name.loc[i,"Subject"]==subject):
            teacher_name=data_name.loc[i,"Lecture by"]
    pdf.cell(20,0,"ຊື່ອາຈານ : ",align="L")
    pdf.cell(0,0,teacher_name,align="L")
    pdf.ln(5)
    pdf.cell(20,0,"ຫ້ອງຮຽນ : ",align="L")
    pdf.cell(0,0,"4COM1",align="L")
    pdf.ln(5)
    pdf.cell(20,0,"ວິຊາ : ",align="L")
    pdf.cell(0,0,"Advance Animation",align="L")
    pdf.ln(10)

    TABLE_COL_NAMES = ("ລຳດັບ", "ລະຫັດນັກສຶກສາ", "ຊື່ແລະນາມສະກຸນ", "ເກຣດ")
    data = pd.read_excel("Score.xlsx",sheetname).reset_index(drop=True)
    data = data.fillna("")
    data=data[["ລ/ດ","ລະຫັດນັກສຶກສາ","ຊື່ ແລະ ນາມສະກຸນ",subject]].to_records(index=False)
    pdf.set_font("Saysettha", size=10)
    line_height = pdf.font_size * 1.5
    col_width = pdf.epw / 4 
    col_width=[21.5,49.5,49.5,49.5]
    # print(col_width)
    def render_table_header():
        pdf.set_font('SaysetthaB', size=10)
        for i in range(0,len(TABLE_COL_NAMES)):
            pdf.cell(col_width[i], line_height, TABLE_COL_NAMES[i], border=1)
        pdf.ln(line_height)
        pdf.set_font('Saysettha', size=10)

    render_table_header()
    for row in data:
        if pdf.will_page_break(line_height):
            render_table_header()
        for i in range(0,len(row)):
            pdf.cell(col_width[i], line_height, str(row[i]), border=1)
        pdf.ln(line_height)

    pdf.output(str(Path.home())+'/Desktop/report_score_by_subject.pdf', 'F')

def report_status(room:str):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margin(20)
    pdf.add_font('Saysettha', fname='saysettha_ot.ttf')
    pdf.add_font('saysetthaB',fname='Saysettha-Bold.ttf')
    pdf.set_font('Saysettha', size=12)
    pdf.image('logo.png',90,8,25)

    pdf.ln(15)
    pdf.cell(0,0,"ສາທາລະນະລັດ ປະຊາທິປະໄຕ ປະຊາຊົນລາວ",align="C")
    pdf.ln(5)
    pdf.cell(0,0,"ສັນຕິພາບ ເອກະລາດ ປະຊາທິປະໄຕ ເອກະພາບ ວັດທະນາຖາວອນ",align="C")
    pdf.ln(15)
    pdf.cell(0,0,"ມະຫາວິທະຍາໄລແຫ່ງຊາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ຄະນະວິສະວະກຳສາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ພາກວິຊາວິສະວະກຳຄອມພິວເຕີ ແລະ ເຕັກໂນໂລຊີຂໍ້ມູນຂ່າວສານ")
    pdf.ln(10)
    pdf.set_font('SaysetthaB', size=14)
    pdf.cell(0,0,"ໃບລາຍງານສະຖານະການສົ່ງເກຣດ",align="C")
    pdf.set_font('Saysettha', size=12)
    pdf.ln(10)
    pdf.cell(20,0,"ຫ້ອງຮຽນ : ",align="L")
    pdf.cell(0,0,"4COM1",align="L")
    pdf.ln(5)
    pdf.ln(10)

    TABLE_COL_NAMES = ("ຊື່ວິຊາ", "ສະຖານະ")

    TABLE_DATA = pd.read_excel("Subject.xlsx",room).fillna("ຍັງບໍ່ໄດ້ສົ່ງ")
    TABLE_DATA =TABLE_DATA.reset_index(drop=True)
    TABLE_DATA=TABLE_DATA[["Subject","Status"]].to_records(index=False)
    pdf.set_font("Saysettha", size=10)
    line_height = pdf.font_size*2
    col_width = pdf.epw / 2

    def render_table_header():
        pdf.set_font('SaysetthaB', size=10)
        for col_name in TABLE_COL_NAMES:
            pdf.cell(col_width, line_height, col_name, border=1)
        pdf.ln(line_height)
        pdf.set_font('Saysettha', size=10)

    render_table_header()
    for row in TABLE_DATA:
        if pdf.will_page_break(line_height):
            render_table_header()
        for datum in row:
            pdf.cell(col_width, line_height, str(datum), border=1)
        pdf.ln(line_height)

    pdf.output(str(Path.home())+'/Desktop/report_status.pdf', 'F')

def report_statistic_by_room(room:str,term:str):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margin(20)
    pdf.add_font('Saysettha', fname='saysettha_ot.ttf')
    pdf.add_font('saysetthaB',fname='Saysettha-Bold.ttf')
    pdf.set_font('Saysettha', size=12)
    pdf.image('logo.png',90,8,25)

    pdf.ln(15)
    pdf.cell(0,0,"ສາທາລະນະລັດ ປະຊາທິປະໄຕ ປະຊາຊົນລາວ",align="C")
    pdf.ln(5)
    pdf.cell(0,0,"ສັນຕິພາບ ເອກະລາດ ປະຊາທິປະໄຕ ເອກະພາບ ວັດທະນາຖາວອນ",align="C")
    pdf.ln(15)
    pdf.cell(0,0,"ມະຫາວິທະຍາໄລແຫ່ງຊາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ຄະນະວິສະວະກຳສາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ພາກວິຊາວິສະວະກຳຄອມພິວເຕີ ແລະ ເຕັກໂນໂລຊີຂໍ້ມູນຂ່າວສານ")
    pdf.ln(10)
    pdf.set_font('SaysetthaB', size=14)
    pdf.cell(0,0,"ໃບລາຍງານສະຖິຕິເກຣດຫ້ອງ",align="C")
    pdf.set_font('Saysettha', size=12)
    pdf.ln(10)
    pdf.cell(20,0,"ຫ້ອງຮຽນ : ",align="L")
    pdf.cell(0,0,room,align="L")
    pdf.ln(10)
    pdf.cell(20,0,"ເທິມ : ",align="L")
    pdf.cell(0,0,term,align="L")
    pdf.ln(5)
    pdf.ln(10)
    subject_file = pd.read_excel("Subject.xlsx",room)
    subjects = []
    for i in range(0,len(subject_file)):
        if(str(subject_file.loc[i,"term"])==term):
            subjects.append(subject_file.loc[i,"Subject"])
    data_total = [0,0,0,0,0,0,0,0,0,0]
    # print(subjects)
    for subject in subjects:
        pdf.set_font('SaysetthaB', size=14)
        pdf.cell(0,0,"ວິຊາ : "+subject,align="C")
        pdf.set_font('SaysetthaB', size=12)
        pdf.ln(10)
        grades_index = ['A', 'B+', 'B','C+', 'C', 'D+','D','F','W',"I"]

        data = [0,0,0,0,0,0,0,0,0,0]
        col=["ເກຣດ","ຈຳນວນ"]
        student_data = pd.read_excel("Score.xlsx",room).reset_index(drop=True)
        student_data = student_data.fillna("")
        student_data = student_data[["ລ/ດ","ລະຫັດນັກສຶກສາ","ຊື່ ແລະ ນາມສະກຸນ",subject]].to_records(index=False)
        # print(student_data)
        for i in student_data:
            try:
                data[grades_index.index(i[3])]=data[grades_index.index(i[3])]+1
                data_total[grades_index.index(i[3])]=data_total[grades_index.index(i[3])]+1
            except:
                continue
        data2 =[]
        for i in range(0,len(grades_index)):
            data2.append([grades_index[i],str(data[i])])
        pdf.set_font("Saysettha", size=10)
        line_height = pdf.font_size * 2
        col_width = pdf.epw / 2 

        pdf.set_font('SaysetthaB', size=10)
        for col_name in col:
            pdf.cell(col_width, line_height, col_name, border=1)
        pdf.ln(line_height)
        pdf.set_font('Saysettha', size=10)

        for row in data2:
            if pdf.will_page_break(line_height):
                pdf.set_font('SaysetthaB', size=10)
                for col_name in col:
                    pdf.cell(col_width, line_height, col_name, border=1)
                pdf.ln(line_height)
                pdf.set_font('Saysettha', size=10)
            for datum in row:
                pdf.cell(col_width, line_height, str(datum), border=1)
            pdf.ln(line_height)


        pdf.set_font('SaysetthaB', size=14)

        pdf.ln(line_height)

        fig = plt.figure()
        plt.bar(grades_index,data)

        fig.savefig(subject+"bar.png", dpi=300, bbox_inches='tight', pad_inches=0)
        plt.close()
        pdf.ln(line_height)
        pdf.image(subject+"bar.png",w=83.33,x=50)
        pdf.add_page()

    pdf.set_font('SaysetthaB', size=14)
    pdf.cell(0,0,"ເກຣດລວມ",align="C")
    pdf.set_font('SaysetthaB', size=12)
    pdf.ln(10)
    grades_index = ['A', 'B+', 'B','C+', 'C', 'D+','D','F','W',"I"]

    col=["ເກຣດ","ຈຳນວນ"]
    # student_data = pd.read_excel("Score.xlsx",room).reset_index(drop=True)
    # student_data = student_data.fillna("")
    # student_data = student_data[["ລ/ດ","ລະຫັດນັກສຶກສາ","ຊື່ ແລະ ນາມສະກຸນ",subject]].to_records(index=False)
    # for i in student_data:
    #     try:
    #         data[grades_index.index(i[4])]=data[grades_index.index(i[4])]+1
    #     except:
    #         continue

    data2 =[]
    for i in range(0,len(grades_index)):
        data2.append([grades_index[i],str(data_total[i])])
    pdf.set_font("Saysettha", size=10)
    line_height = pdf.font_size * 2
    col_width = pdf.epw / 2 

    pdf.set_font('SaysetthaB', size=10)
    for col_name in col:
        pdf.cell(col_width, line_height, col_name, border=1)
    pdf.ln(line_height)
    pdf.set_font('Saysettha', size=10)

    for row in data2:
        if pdf.will_page_break(line_height):
            pdf.set_font('SaysetthaB', size=10)
            for col_name in col:
                pdf.cell(col_width, line_height, col_name, border=1)
            pdf.ln(line_height)
            pdf.set_font('Saysettha', size=10)
        for datum in row:
            pdf.cell(col_width, line_height, str(datum), border=1)
        pdf.ln(line_height)


    pdf.set_font('SaysetthaB', size=14)

    pdf.ln(line_height)

    fig = plt.figure()
    plt.bar(grades_index,data)

    fig.savefig("All"+"bar.png", dpi=300, bbox_inches='tight', pad_inches=0)
    plt.close()
    pdf.ln(line_height)
    pdf.image("bar.png",w=83.33,x=50)



    pdf.output(str(Path.home())+'/Desktop/report_statistic.pdf', 'F')

def report_by_major(major:str,year:str):

    pdf = FPDF()
    pdf.add_page()
    pdf.set_margin(20)
    pdf.add_font('Saysettha', fname='saysettha_ot.ttf')
    pdf.add_font('saysetthaB',fname='Saysettha-Bold.ttf')
    pdf.set_font('Saysettha', size=12)
    pdf.image('logo.png',90,8,25)

    pdf.ln(15)
    pdf.cell(0,0,"ສາທາລະນະລັດ ປະຊາທິປະໄຕ ປະຊາຊົນລາວ",align="C")
    pdf.ln(5)
    pdf.cell(0,0,"ສັນຕິພາບ ເອກະລາດ ປະຊາທິປະໄຕ ເອກະພາບ ວັດທະນາຖາວອນ",align="C")
    pdf.ln(15)
    pdf.cell(0,0,"ມະຫາວິທະຍາໄລແຫ່ງຊາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ຄະນະວິສະວະກຳສາດ")
    pdf.ln(5)
    pdf.cell(0,0,"ພາກວິຊາວິສະວະກຳຄອມພິວເຕີ ແລະ ເຕັກໂນໂລຊີຂໍ້ມູນຂ່າວສານ")
    pdf.ln(10)
    pdf.set_font('SaysetthaB', size=14)
    pdf.cell(0,0,"ໃບລາຍງານສະຖິຕິເກຣດຕາມສາຍຮຽນ",align="C")
    pdf.set_font('Saysettha', size=12)
    pdf.ln(10)
    pdf.cell(20,0,"ສາຍຮຽນ : ",align="L")
    pdf.cell(0,0,major,align="L")
    pdf.ln(10)
    pdf.cell(20,0,"ປີ : ",align="L")
    pdf.cell(0,0,year,align="L")
    pdf.ln(5)
    pdf.ln(10)

    grades_index = ['A', 'B+', 'B','C+', 'C', 'D+','D','F','W',"I"]
    data = [0,0,0,0,0,0,0,0,0,0]
    col=["ເກຣດ","ຈຳນວນ"]

    namefile = excel_info("Score.xlsx").sheet_names
    subjectfile = excel_info("Subject.xlsx").sheet_names
    matched_room = list(set(namefile)&set(subjectfile))
    room_list = []
    [room_list.append(x) for x in matched_room if x[1:-1]==major]
    for room in room_list:
        subject_file = pd.read_excel("Subject.xlsx",room)
        subjects = []
        for i in range(0,len(subject_file)):
            if(str(subject_file.loc[i,"year"])==year):
                subjects.append(subject_file.loc[i,"Subject"])
        for subject in subjects:
            student_data = pd.read_excel("Score.xlsx",room).reset_index(drop=True)
            student_data = student_data.fillna("")
            student_data = student_data[["ລ/ດ","ລະຫັດນັກສຶກສາ","ຊື່ ແລະ ນາມສະກຸນ",subject]].to_records(index=False)
            # print(student_data)
            for i in student_data:
                try:
                    data[grades_index.index(i[3])]=data[grades_index.index(i[3])]+1
                except:
                    continue
    data2 =[]
    for i in range(0,len(grades_index)):
        data2.append([grades_index[i],str(data[i])])
    pdf.set_font("Saysettha", size=10)
    line_height = pdf.font_size * 2
    col_width = pdf.epw / 2 

    pdf.set_font('SaysetthaB', size=10)
    for col_name in col:
        pdf.cell(col_width, line_height, col_name, border=1)
    pdf.ln(line_height)
    pdf.set_font('Saysettha', size=10)

    for row in data2:
        if pdf.will_page_break(line_height):
            pdf.set_font('SaysetthaB', size=10)
            for col_name in col:
                pdf.cell(col_width, line_height, col_name, border=1)
            pdf.ln(line_height)
            pdf.set_font('Saysettha', size=10)
        for datum in row:
            pdf.cell(col_width, line_height, str(datum), border=1)
        pdf.ln(line_height)


    pdf.set_font('SaysetthaB', size=14)

    pdf.ln(line_height)

    fig = plt.figure()
    plt.bar(grades_index,data)

    fig.savefig("bar.png", dpi=300, bbox_inches='tight', pad_inches=0)
    plt.close()
    pdf.ln(line_height)
    pdf.image("bar.png",w=83.33,x=50)
    pdf.output(str(Path.home())+'/Desktop/report_statistic_by_major.pdf', 'F')
