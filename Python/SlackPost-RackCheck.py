#--------------------------------------------------------------------------------------------------#
# SlackPost-RackCheck.py: Rack Check Data Collection and Send a message to Slack                   #
#--------------------------------------------------------------------------------------------------#
#  AUTHOR: Yuhui.Seo        2022/08/16                                                             #
#--< CHANGE HISTORY >------------------------------------------------------------------------------#
#          Yuhui.Seo        2022/XX/XX #001(Change XXX)                                            #
#--------------------------------------------------------------------------------------------------#

#--------------------------------------------------------------------------------------------------#
# Main process                                                                                     #
# Python version 3.10.0                                                                            #
#--------------------------------------------------------------------------------------------------#
from email import header
from openpyxl import load_workbook # requires: pip install openpyxl
from slack_sdk import WebClient    # requires: pip install slack_sdk
import pymssql                     # requires: pip install pymssql
import pandas as pd                # requires: pip install pandas
from datetime import datetime
from slack_sdk.errors import SlackApiError
from openpyxl.styles import Border, Side, Protection
import os
import socket
import warnings
warnings.filterwarnings('ignore')

# TODO: 1. 엑셀 인쇄 페이지 자동설정

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


def system_info():
    hostname = socket.gethostname() # PC명
    # ip = socket.gethostbyname(hostname) #IP주소
    return hostname


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


# RackCheck Bot
#
# Rack Check List
#
# PCNAME, YYYY-MM-DD HH:mm:ss
# 랙정상화 대상 : XX건
#
# ---------------------------------------------------------
#
# 연번  랙ID  S   Z   X   Y   상태
# 랙정상리스트 엑셀파일
# ---------------------------------------------------------
class SlackAPI:
    """
    슬랙 API 핸들러
    """
    def __init__(self, token):
        # 슬랙 클라이언트 인스턴스 생성
        self.client = WebClient(token)

    def post_Message(self, channel_id, msg, index):
        hostname = system_info()
        try:
            response= self.client.chat_postMessage(
            channel= channel_id,
            blocks=[
                {
                    "type": "header",
                    "text": {
                        "type": "plain_text",
                        "text": "Rack Check List"
                    }
                },
                {
                    "type": "context",
                    "elements": [
                        {
                            "type": "plain_text",
                            "text": hostname + ", " + get_current_datetime(2)
                        }
                    ]
                },
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": msg
                    }
                }
            ]
        )
        except SlackApiError as e:
            print(e.response['error'])

    def post_files_upload(self, channel_id, msg):
        hostname = system_info()
        try:
            response_msg = self.client.chat_postMessage(
                channel= channel_id,
                blocks=[
                    {
                        "type": "header",
                        "text": {
                            "type": "plain_text",
                            "text": "Rack Check List"
                        }
                    },
                    {
                        "type": "context",
                        "elements": [
                            {
                                "type": "plain_text",
                                "text": hostname + ", " + get_current_datetime(2)
                            }
                        ]
                    },
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": msg
                        }
                    },
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": "index  랙ID  S   Z   X   Y   상태 \n" + 
                                    "1   랙ID  S   Z   X   Y   상태 \n" +
                                    "2   랙ID  S   Z   X   Y   상태 \n" +
                                    "3   랙ID  S   Z   X   Y   상태 \n"
                        }
                    }
                ]
            )
            response_xlsx= self.client.files_upload(
                channels= channel_id,
                file= './Download-SVR354/AW_RACK_CHECK_20220812154027.xlsx',
                filename= 'AW_RACK_CHECK_20220812154027.xlsx', # 다운로드 시 파일명(확장자까지 설정 필요)
                filetype= 'xlsx',
                title= 'AW_RACK_CHECK_20220812154027.xlsx', # Slack 파일 첨부의 파일명
                initial_comment= 'initial_comment입니다.',
            )
        except SlackApiError as e:
            print(e.response['error'])

def set_slack_token():
    # Bot OAuth Token:'xoxb-XXXXXXXXXXXXXXXXXX'
    token= 'xoxb-XXXXXXXXXXXXXXXXXX'
    return token


def set_slack_channel():
    channel_id= 'CXXXXXXXXXX'
    return channel_id
    
    
def set_result(result):
    if result == False:
        msg = '> 현재 랙정상화 대상은 없습니다! :tada:'
    else:
        msg = '> 현재 랙정상화 대상은 XX건 입니다.'
    return msg


def set_slack_msg():
    print(str(df))
    index = '결과값'
    return index


#--------------------------------------------------------------------------------------------------#
# Code Entry                                                                                       #
#--------------------------------------------------------------------------------------------------#

df = select_sqldata()
now = datetime.now()
slack = SlackAPI(set_slack_token())
channel_id = set_slack_channel()


if df.empty:                                            # 1.SELECT RackCheck # TODO : Slack에서 /커맨드로 실행
    slack.post_Message(channel_id, set_result(False)) # 2.랙정상화 대상이 없다면 Slack 전송
else:
    data_row = save_excel()                             # 3.랙정상화 대상이 있다면 csv파일, excel파일 생성
    slack.post_files_upload(channel_id, set_result(True), set_slack_msg()) # 4.excel파일 Slack 전송
