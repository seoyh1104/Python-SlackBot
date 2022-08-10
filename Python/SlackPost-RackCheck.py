#--------------------------------------------------------------------------------------------------#
# SlackPost-RackCheck.py: Rack Check Data Collection and Send a message to Slack                   #
#--------------------------------------------------------------------------------------------------#
#  AUTHOR: Yuhui.Seo        2022/08/08                                                             #
#--< CHANGE HISTORY >------------------------------------------------------------------------------#
#          Yuhui.Seo        2022/XX/XX #001(Change XXX)                                            #
#--------------------------------------------------------------------------------------------------#

#--------------------------------------------------------------------------------------------------#
# Main process                                                                                     #
# Python version 3.10.0                                                                            #
#--------------------------------------------------------------------------------------------------#
import pymssql                     # pip install pymssql
import pandas as pd                # pip install pandas
import requests                    # pip install requests
from openpyxl import load_workbook # pip install openpyxl
from openpyxl.styles import Border, Side, Protection
from openpyxl.styles.colors import Color
from datetime import datetime
import os
import warnings

warnings.filterwarnings('ignore')

# Connect Python to SQL Server
def select_sqldata():
    server = '1.121.XX.XX'
    database = 'XXXX'
    username = 'XX'
    password = 'XXXX' # TODO: PW 암호화 필요

    # DB Connection
    conn = pymssql.connect(server, username, password, database, charset='utf8')

    # FIXME: .sql file로 실행하도록 수정
    sql = '''
    
    '''

    df = pd.read_sql(sql = sql, con = conn)
    conn.close()
    return df


def get_current_datetime(format):
    match format:
        case 1:
            format = '%Y%m%d%H%M%S' # yyyyMMddHHmmss
        case 2:
            format = '%Y년%m월%d일%H시%M분%S초' # yyyy년MM월dd일HH시mm분ss초
        case _:
            format = '%Y%m%d%H%M%S'
    current_datetime = now.strftime(format)
    return current_datetime


def exists_dir():
    current_datetime = get_current_datetime(1)
    directory = './폴더명' # NOTE: 폴더명 설정
    filename = '/파일명'   # NOTE: 폴더명 설정
    
    if not os.path.exists(directory):
        os.mkdir(directory)
    
    return current_datetime, directory, filename


def save_filepath():
    current_datetime, directory, filename = exists_dir()
    filepath = directory + filename + '_' + current_datetime
    return filepath


def save_csv(df):
    csv_file = save_filepath() + '.csv'
    df.to_csv(csv_file, header = True, index = False, encoding = 'utf-8')
    print("Csv file saved successfully")
    return csv_file


def csv_to_excel():
    r_csv = pd.read_csv(save_csv(df))
    xlsx_file = save_filepath() + '.xlsx'
    save_xlsx = pd.ExcelWriter(xlsx_file)
    r_csv.to_excel(save_xlsx, index = False) # xlsx 파일로 변환
    save_xlsx.save() #xlsx 파일로 저장
    print("Excel file saved successfully")
    return save_xlsx


def save_excel():
    wb = load_workbook(csv_to_excel())
    ws = wb.active
    
    data_row = str(ws.max_row - 1)
    
    set_sheetdata(ws)
    set_column_size(ws)
    set_color_border(ws)

    xlsx_file = save_filepath() + '.xlsx'
    wb.save(xlsx_file)
    wb.close()
    
    return data_row


def set_sheetdata(ws):
    current_datetime = get_current_datetime(2)
    ws.title = '시트명_' + current_datetime # NOTE: 시트명 설정 '랙정상화대상_'
    ws.sheet_properties.tabColor = 'F4566E'
    
    # 숫자 형식으로 표시
    column = 'P' # P열:재고관리코드
    for row in range(2, ws.max_row + 1):
        ws[column + str(row)].number_format = '0'
    
    ws.delete_cols(21) # 21번째 열:RANK 삭제
    
    return ws


def set_column_size(ws):
    
    AutoFitColumnSize(ws)
    ws.column_dimensions['G'].width = 9 # 상태
    ws.column_dimensions['H'].width = 9 # 입고자명
    ws.column_dimensions['J'].width = 11 # 마지막작업
    ws.column_dimensions['L'].width = 13 # 마지막작업자
    ws.column_dimensions['O'].width = ws.column_dimensions['O'].width + 5 # 상품명
    ws.column_dimensions['S'].width = 9 # 상태코드
        
    return ws


def set_color_border(ws):
    border_color = '000000'

    for col in range(ws.max_column):
        for row in range(ws.max_row):
            cell = ws.cell(row = row + 1, column = col + 1)
            cell.border = make_color_border(border_color)
    
    return ws


def make_color_border(color):
    border = Border(left = Side(style='thin', color=color),
                    right = Side(style='thin', color=color),
                    top = Side(style='thin', color=color),
                    bottom = Side(style='thin', color=color))
    return border


# culumns is passed by list and element of columns means column index in worksheet.
# if culumns = [1, 3, 4] then, 1st, 3th, 4th columns are applied autofit culumn.
# margin is additional space of autofit column. 
def AutoFitColumnSize(worksheet, columns = None, margin = 2):
    for i, column_cells in enumerate(worksheet.columns):
        is_ok = False
        if columns == None:
            is_ok = True
        elif isinstance(columns, list) and i in columns:
            is_ok = True
            
        if is_ok:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + margin

    return worksheet


def set_slack_payload(data_row):
    blocks = ''
    if data_row == '': # TODO: 랙정상화 대상 없을 때, 테스트 필요!
        text = 'text = :memo: 랙정상화 대상은 없습니다.'
    else:
        text = ':memo: SlackPost-RackCheck'
        blocks = [
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": ":memo: 랙정상화 대상은 " + data_row + "건 입니다."
                    }
                }
            ]
    
    return text, blocks


def send_message_to_slack(data_row):
    
    text, blocks = set_slack_payload(data_row)
    
    url = 'https://hooks.slack.com/services/TXXXXXXX' # WebHook Url

    headers = {
        'Content-Type': 'application/json'
        }
    
    payload = {
        'text': text,
        'blocks' : blocks
        }
    
    # requests.post(url, headers = headers, json = payload)
    res = requests.post(url, headers = headers, json = payload)
    print(res.status_code)

#--------------------------------------------------------------------------------------------------#
# Code Entry                                                                                       #
#--------------------------------------------------------------------------------------------------#

df = select_sqldata()
now = datetime.now()

if df.empty:                     # 1.SELECT Rack Check #TODO : Slack에서 /커맨드로 실행
    print('DataFrame is empty!') # 2.결과가 없다면 Slack 전송
    # send_message_to_slack(True)
else:
    data_row = save_excel()         # 3.결과가 있다면 csv파일, excel파일 생성
    send_message_to_slack(data_row) # 4.결과와 Excel파일 Slack 업로드