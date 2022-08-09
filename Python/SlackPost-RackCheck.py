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
from openpyxl import load_workbook # pip install openpyxl
from openpyxl.styles import Protection
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
    current_datetime = get_current_datetime(2)
    ws.title = '시트명_' + current_datetime # NOTE: 시트명 설정
    ws.sheet_properties.tabColor = 'F4566E'

    ws.protection = Protection(locked = True, hidden = False) # Protection
    
    xlsx_file = save_filepath() + '.xlsx'
    wb.save(xlsx_file)
    wb.close()


#--------------------------------------------------------------------------------------------------#
# Code Entry                                                                                       #
#--------------------------------------------------------------------------------------------------#

df = select_sqldata()
now = datetime.now()

if df.empty:                     # 1.SELECT Rack Check #TODO : Slack에서 /커맨드로 실행
    print('DataFrame is empty!') # 2.결과가 없다면 Slack 전송
else:        
    save_excel()                 # 3.결과가 있다면 csv파일, excel파일 생성
                                 # 4.결과와 Excel파일 Slack 업로드



####################

# from slack_sdk.webhook import WebhookClient
# # Slack Test - seoyuhui.slack.com
# url = 'https://hooks.slack.com/services/TXXXXXX'

# webhook = WebhookClient(url)
# response = webhook.send(
#     text=":memo: SlackPost-RackCheck",
#     blocks=[
#         {
#             "type": "section",
#             "text": {
#                 "type": "mrkdwn",
#                 "text": ":memo: SlackPost-RackCheck / 테스트 / テスト"
#             }
#         }
#     ]
# )