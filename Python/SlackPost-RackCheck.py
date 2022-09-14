#--------------------------------------------------------------------------------------------------#
# SlackPost-RackCheck.py: Rack Check Data Collection and Send a message to Slack                   #
#--------------------------------------------------------------------------------------------------#
#  AUTHOR: Yuhui.Seo        2022/08/16                                                             #
#--< CHANGE HISTORY >------------------------------------------------------------------------------#
#          Yuhui.Seo        2022/09/08 #001(Add config file)                                       #
#--< Version >-------------------------------------------------------------------------------------#
#          Python version 3.10.0 (requires Python version 3.10 or higher.)                         #
#--------------------------------------------------------------------------------------------------#

#--------------------------------------------------------------------------------------------------#
# Main process                                                                                     #
#--------------------------------------------------------------------------------------------------#
import pymssql                             # requires: pip install pymssql
import pandas as pd                        # requires: pip install pandas
from slack_sdk import WebClient            # requires: pip install slack_sdk
from slack_sdk.errors import SlackApiError
from openpyxl import load_workbook         # requires: pip install openpyxl
from openpyxl.styles import Border, Side
from datetime import datetime
import os
import socket
import configparser
#--------------------------------------------------------------------------------------------------#
class SystemInfo:
    def system_info():
        hostname = socket.gethostname() # PC명
        # ip = socket.gethostbyname(hostname) #IP주소
        return hostname
    
    def get_current_datetime(now, format):
        match format:
            case 1:
                format = '%Y%m%d%H%M%S' # yyyyMMddHHmmss
            case 2:
                format = '%Y-%m-%d %H:%M:%S' # yyyy-MM-dd HH:mm:ss
            case 3:
                format = '%Y년%m월%d일 %H시%M분%S초' # yyyy년MM월dd일HH시mm분ss초
            case _:
                format = '%Y%m%d%H%M%S'
        current_datetime = now.strftime(format)
        return current_datetime


class ReadConfig:
    def __init__(self):
        self.conf_file = 'config.ini'
    
    def load_config(self, section):
        if os.path.exists(self.conf_file) == False:
            raise Exception('%s file does not exist. \n' % self.conf_file)
        else: 
            config = configparser.ConfigParser()
            config.read(self.conf_file, encoding = 'utf-8')
            print("Load Config : %s" % self.conf_file)
            
            self.db_config = config[section]
            return self.db_config
        
        
class MssqlController:
    def __init__(self):
        self.db_config = ReadConfig.load_config(ReadConfig(), 'DB_CONFIG')
        
    def __connect__(self):
        try:
            self.conn = pymssql.connect(server = self.db_config['server'],
                                        database = self.db_config['database'],
                                        user = self.db_config['username'], 
                                        password = self.db_config['password'],
                                        charset = self.db_config['charset'])
            self.cur = self.conn.cursor(as_dict = True)
        except Exception as e:
            print('DB not connected: ', e)
            raise
            
    def execute(self, sql):
        self.__connect__()
        self.cur.execute(sql)
        df = pd.DataFrame(self.fetch())
        self.__disconnect__()
        return df
    
    def fetch(self):
        return self.cur.fetchall()
    
    def __disconnect__(self):
        self.conn.close()

# FIXME: .sql file로 실행할 수 있도록 수정
sql = '''
SELECT * FROM
WHERE
ORDER BY
'''

# TODO: 1. Config화
# TODO: 2. EXCEL 인쇄페이지 자동설정
class CreateFile:
    def __init__(self, now):
        self.paths_config = ReadConfig.load_config(ReadConfig(), 'PATHS_CONFIG')
        self.excel_config = ReadConfig.load_config(ReadConfig(), 'EXCEL_CONFIG')
        self.CSV = self.paths_config['CSV']
        self.XLSX = self.paths_config['XLSX']
        self.file_directory = self.paths_config['file_directory']
        self.file_name = self.paths_config['file_name']
        self.now = now
    
    def exists_dir(self):
        current_datetime = SystemInfo.get_current_datetime(self.now, 1)
        if not os.path.exists(self.file_directory):
            os.mkdir(self.file_directory)
        return current_datetime, self.file_directory, self.file_name

    def save_filepath(self):
        current_datetime, directory, filename = self.exists_dir()
        filepath = directory + filename + '_' + current_datetime
        return filepath

    def save_csv(self, df):
        csv_file = self.save_filepath() + self.CSV
        df.to_csv(csv_file, header = True, index = False, encoding = 'utf-8')
        print('Csv file saved successfully')
        return csv_file

    def csv_to_excel(self, df):
        r_csv = pd.read_csv(self.save_csv(df))
        xlsx_file = self.save_filepath() + self.XLSX
        save_xlsx = pd.ExcelWriter(xlsx_file)
        r_csv.to_excel(save_xlsx, index = False) # xlsx 파일로 변환
        save_xlsx.save() #xlsx 파일로 저장
        print('Excel file saved successfully')
        return save_xlsx

    def save_excel(self, df):
        wb = load_workbook(self.csv_to_excel(df))
        ws = wb.active
        
        self.set_sheetdata(ws)
        self.set_column_size(ws)
        self.set_color_border(ws)
        
        xlsx_file = self.save_filepath() + '.xlsx'
        wb.save(xlsx_file)
        wb.close()
        return xlsx_file

    def set_sheetdata(self, ws):
        current_datetime = SystemInfo.get_current_datetime(self.now, 3)
        # 시트명 변경
        ws.title = self.excel_config['sheet_title'] + '_' + current_datetime
        # 시트탭 색변경
        ws.sheet_properties.tabColor = self.excel_config['TAB_COLOR_BLUE']
        # ws.sheet_properties.tabColor = TAB_COLOR_PINK
        # 숫자 형식으로 표시
        column = 'P' # P열:재고관리코드
        for row in range(2, ws.max_row + 1):
            ws[column + str(row)].number_format = '0'
        # 열 삭제
        ws.delete_cols(21) # 21번째 열:RANK 삭제
        return ws

    def set_column_size(self, ws):
        ws.column_dimensions['A'].width = 4  # ✔
        ws.column_dimensions['B'].width = 10 # 랙ID
        ws.column_dimensions['G'].width = 9  # 상태
        ws.column_dimensions['H'].width = 9  # 입고자명
        ws.column_dimensions['J'].width = 11 # 마지막작업
        ws.column_dimensions['K'].width = 11 # 마지막작업일
        ws.column_dimensions['L'].width = 14 # 마지막작업자
        ws.column_dimensions['O'].width = ws.column_dimensions['O'].width + 5 # 상품명
        ws.column_dimensions['S'].width = 9  # 상태코드
        return ws

    def set_color_border(self, ws):
        for col in range(ws.max_column):
            for row in range(ws.max_row):
                cell = ws.cell(row = row + 1, column = col + 1)
                cell.border = self.make_color_border()
        return ws

    def make_color_border(self):
        color = self.excel_config['BORDER_COLOR']
        border = Border(left = Side(style='thin', color=color),
                        right = Side(style='thin', color=color),
                        top = Side(style='thin', color=color),
                        bottom = Side(style='thin', color=color))
        return border

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

    def __init__(self, token, now):
        # 슬랙 클라이언트 인스턴스 생성
        self.client = WebClient(token)
        slack_config = ReadConfig.load_config(ReadConfig(), 'SLACK_CONFIG')
        paths_config = ReadConfig.load_config(ReadConfig(), 'PATHS_CONFIG')
        self.channel_id = slack_config['channel_id']
        self.XLSX = paths_config['XLSX']
        self.now = now
        
    def post_Message(self, msg):
        hostname = SystemInfo.system_info()
        try:
            response= self.client.chat_postMessage(
            channel= self.channel_id,
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
                            "text": hostname + ", " + SystemInfo.get_current_datetime(self.now, 2)
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

    def post_files_upload(self, msg, result, file_name):
        hostname = SystemInfo.system_info()
        try:
            response_msg = self.client.chat_postMessage(
                channel= self.channel_id,
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
                                "text": hostname + ", " + SystemInfo.get_current_datetime(self.now, 2)
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
                            "text": result
                        }
                    }
                ]
            )
            response_xlsx= self.client.files_upload(
                channels= self.channel_id,
                file= file_name,
                filename= 'AW_RACK_CHECK_다운로드시파일명.xlsx', # 다운로드시 파일명(확장자까지 설정 필요)
                filetype= self.XLSX,
                title= 'AW_RACK_CHECK_첨부파일의파일명.xlsx', # 첨부파일의 파일명
                initial_comment= 'initial_comment입니다.'
            )
        except SlackApiError as e:
            print(e.response['error'])


def set_msg(result, df):
    if result == False:
        msg = '현재 랙정상화 대상은 없습니다! :tada:'
    else:
        df_cnt = len(df)
        msg = '> 현재 랙정상화 대상은 ' + str(df_cnt) + '건 입니다.'
    return msg

def set_result(df):
    index = df[['랙ID', 's', 'z', 'x', 'y', '상태']]
    return index

def process1():
    df = MssqlController.execute(MssqlController(), sql)
    return df
    
#--------------------------------------------------------------------------------------------------#
# Code Entry                                                                                       #
#--------------------------------------------------------------------------------------------------#
def main():
    slack_config = ReadConfig.load_config(ReadConfig(), 'SLACK_CONFIG')
    SLACK_TOKEN = slack_config['SLACK_TOKEN']
    now = datetime.now()
    slack = SlackAPI(SLACK_TOKEN, now)
    df = process1()                        # 1.SELECT RackCheck # TODO : Slack에서 /커맨드로 실행
    
    if df.empty:                                           
        slack.post_Message(set_msg(False)) # 2.랙정상화 대상이 없다면 Slack 전송
    else:
        create_file = CreateFile(now)
        file_name = CreateFile.save_excel(create_file, df)
        # CreateFile(df, now)              # 3.랙정상화 대상이 있다면 csv파일, excel파일 생성
        slack.post_files_upload(set_msg(True, df), set_result(df), file_name) # 4. Slack 전송 (excel파일 및 요약데이터)
    
if __name__ == "__main__":
    main()