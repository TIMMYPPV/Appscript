from pprint import pprint
from googleapiclient import discovery
from google.oauth2 import service_account
import pandas as pd
from os.path import exists,abspath, dirname
import os

SCOPES = ['https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/spreadsheets']
file_exists = exists(dirname(abspath(__file__))[:-8]+'/token.json')
# # os.chdir('../')
# print(dirname(abspath(__file__))[:-8])
os.chdir(dirname(abspath(__file__))[:-8])
# print(os.path.realpath(__file__))
# print(file_exists)
if file_exists:
    credentials = service_account.Credentials.from_service_account_file('token.json', scopes=SCOPES)
    sheetservice = discovery.build('sheets', 'v4', credentials=credentials)
    drive_service = discovery.build('drive', 'v3', credentials=credentials)

def callback(request_id, response, exception):
    if exception:
        # Handle error
        print(exception)
    # else:
    #     print("Permission Id: %s" % response.get('id'))

def create_sheet_from_teacher(classroom:str,teacher_name:str):
    spreadsheet_body = {
    'properties': {
        'title': teacher_name,
    },
    'sheets':{
            "properties":{
                "title": classroom,
            }
        }
    }
    request = sheetservice.spreadsheets().create(body=spreadsheet_body)
    response = request.execute()
    # print(response["sheets"][0]["properties"]["sheetId"])
    return (response["spreadsheetId"],response["spreadsheetUrl"],response["sheets"][0]["properties"]["sheetId"])

def create_sheet_form(classroom:str,spreadsheetId:str,subject_name:str,teacher_name:str):
    sheet = sheetservice.spreadsheets()

    xl = pd.read_excel("Score.xlsx", classroom)
    data = xl
    # print(data)
    # temp_read = data.drop([0,1])
    # new_header = temp_read.iloc[0]
    # temp_read = temp_read[1:]
    # temp_read.columns = new_header
    temp_read = data[["ລ/ດ","ລະຫັດນັກສຶກສາ","ຊື່ ແລະ ນາມສະກຸນ"]]
    temp_read =temp_read.values.tolist()

    value = []
    value.append(["ລາຍງານຄະແນນວິຊາ"])
    value.append([subject_name])
    value.append(['ຫ້ອງ',classroom])
    value.append(['ສອນໂດຍ',teacher_name])
    value.append(["ລ/ດ","ລະຫັດນັກສຶກສາ","ຊື່ ແລະ ນາມສະກຸນ","Grade"])
    for i in temp_read:
        value.append(i)
    # print(type(value))
    result = sheet.values().update(spreadsheetId=spreadsheetId,
                                    valueInputOption="USER_ENTERED",
                                    range=classroom,
                                    body={"values":value}).execute()

def grant_permission(email:str,spreadsheetId:str):
    batch = drive_service.new_batch_http_request(callback=callback)
    user_permission = {
    'type': 'user',
    'role': 'writer',
    'emailAddress': email
    }
    batch.add(drive_service.permissions().create(
            fileId=spreadsheetId,
            body=user_permission,
            fields='id',
    ))
    batch.execute()

def auto_adjust(sheet_id:int,spreadsheet_id:str):
    body = {"requests": [{"autoResizeDimensions": { "dimensions": {"sheetId":str(sheet_id),"dimension": "COLUMNS","startIndex":0,"endIndex": 12},}},{"updateBorders": {"range": {"sheetId": str(sheet_id),"startRowIndex": 4,"startColumnIndex": 0,"endColumnIndex": 4},"top": {"style": "SOLID","width": 1,"color": {"blue": 0,"red":0,"green":0},},"bottom": {"style": "SOLID","width": 1,"color": {"blue": 0,"red":0,"green":0},},"left":{"style": "SOLID","width": 1,"color": {"blue": 0,"red":0,"green":0},},"right":{"style": "SOLID","width": 1,"color": {"blue": 0,"red":0,"green":0},},"innerHorizontal": {"style": "SOLID","width": 1,"color": {"blue": 0,"red":0,"green":0},},"innerVertical": {"style": "SOLID","width": 1,"color": {"blue": 0,"red":0,"green":0},}}}]}
    response = sheetservice.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,body=body).execute()

def pull_score(spreadsheet_id:str,room_num:str):
    response = sheetservice.spreadsheets().values().get(spreadsheetId=spreadsheet_id,range=room_num+"!A5:D",majorDimension='ROWS').execute()
    values = response.get('values', [])
    # print(values)
    data = values[1:]
    v=[]
    grades_index = ['A', 'B+', 'B','C+', 'C', 'D+','D','F','W',"I",'a', 'b+', 'b','c+', 'c', 'd+','d','f','w',"i"]
    for i in data:
        try:
            if i[3] in grades_index:
                v.append({ "ລ/ດ":i[0], "ລະຫັດນັກສຶກສາ":i[1], "ຊື່ ແລະ ນາມສະກຸນ":i[2], "Grade": i[3].upper()})
            else:
                v.append({ "ລ/ດ":i[0], "ລະຫັດນັກສຶກສາ":i[1], "ຊື່ ແລະ ນາມສະກຸນ":i[2], "Grade": "?"})
        except:
            v.append({ "ລ/ດ":i[0], "ລະຫັດນັກສຶກສາ":i[1], "ຊື່ ແລະ ນາມສະກຸນ":i[2], "Grade": ""})
    df = pd.DataFrame(v,columns=values[0])
    return df

def pull_info(spreadsheet_id:str,room_num:str):
    subject_name = sheetservice.spreadsheets().values().get(spreadsheetId=spreadsheet_id,range=room_num+"!A2",majorDimension='ROWS').execute().get('values')[0]
    roomnum = sheetservice.spreadsheets().values().get(spreadsheetId=spreadsheet_id,range=room_num+"!B3",majorDimension='ROWS').execute().get('values')[0]
    return roomnum,subject_name
