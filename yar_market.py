# 名稱:     yar_market.py
# 目的:     提供關於近兩年每日開市狀態和相關方法的功能,以及 建立資料庫連接、創建表格插入和查詢數據等操作
# 類別:     StockMarket 
# 函式:     create_table, insert_data, execute_query  
# 版本:     20230523-A01
import sqlite3
import os
import csv
import requests
import datetime
import time
trading_hours = {
    0: [(datetime.time(5, 0, 0),   datetime.time(8, 45, 0))],    # 日盤前
    1: [(datetime.time(8, 45, 0),  datetime.time(13, 45, 0))],   # 日盤
    2: [(datetime.time(13, 45, 0), datetime.time(15, 0, 0))],    # 夜盤前
    3: [(datetime.time(15, 0, 0),  datetime.time(23, 59, 59)),   # 夜盤
        (datetime.time(0, 0, 0),   datetime.time(5, 0, 0))]}     # 夜盤
class StockMarket:
    # 類別 近兩年每日開市狀態與方法
    def __init__(self, path_db, path_tfex):
        self.path_db=path_db
        self.path_tfex=path_tfex
        self.T01_NAME = "t01_market"
        self.status = {}
        now = datetime.datetime.now()
        start_date = datetime.datetime(now.year - 1, 1, 1)
        end_date = datetime.datetime(now.year, 12, 31)
        delta = end_date - start_date 
        days = delta.days + 1        
        start_date0 = int(start_date.strftime('%Y%m%d'))
        end_date0 = int(end_date.strftime('%Y%m%d'))
        create_table(self.path_db,self.T01_NAME,"(date01 CHAR(8) PRIMARY KEY, open01 INTEGER)")
        sql_str=f"SELECT date01, open01 FROM {self.T01_NAME} WHERE date01 BETWEEN {start_date0} AND {end_date0}"
        rows = execute_query(self.path_db,sql_str)
        if  len(rows) < days :
            self._t01_download_insert(now.year - 1)
            self._t01_download_insert(now.year)
            rows = execute_query(self.path_db,sql_str)             
        for row in rows:
                self.status[row[0]] = (row[1] == 1)

    def _t01_download_insert(self,year1: int):
        # 下載指定年的每日開市資料,並新增入t01數據表
        url = f"https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear={year1}"
        fname_csv = os.path.join(self.path_tfex, f"holidaySchedule_{year1}.csv")
        if  not os.path.isfile(fname_csv):
            for i in range(3):
                response = requests.get(url)
                if response.status_code == 200 and len(response.content) >= 1024:
                    with open(fname_csv, "wb") as f:
                        f.write(response.content)
                    break            
        # 解析下載的CSV檔案，並取得休市日的日期清單
        sx = []
        with open(fname_csv, "r") as file:
            reader = csv.reader(file)
            for i, row in enumerate(reader):
                if i < 2:
                    continue
                if row[1] == "" and row[3] == "":
                    break
                if not ("開始交易" in row[3]) and not ("結束交易" in row[3]):
                    dt = datetime.datetime.strptime(row[1], '%m月%d日')
                    sx.append(f"{year1}{dt.month:02d}{dt.day:02d}")
        s1 = str(year1) + "0101"
        # 計算指定年份的所有日期，並新增到t01數據表中
        start_date = datetime.date.fromisoformat(f"{year1}-01-01")
        end_date = datetime.date.fromisoformat(f"{year1}-12-31")
        delta = datetime.timedelta(days=1)
        rows = []
        while start_date <= end_date:
            day_str = start_date.strftime("%Y%m%d")
            k=1
            if day_str in sx :
                k=0
            date0 = datetime.datetime.strptime(day_str, "%Y%m%d")
            if date0.weekday()==5 or  date0.weekday()==6 :
                k=0
            rows.append([day_str, k])
            start_date += delta
        insert_data(self.path_db, self.T01_NAME, rows, sql1="(date01, open01) VALUES (?, ?)") 

    def get_status(self, day_str):
        # 指定日期 得到是否為開市日
        return self.status.get(day_str, False)        

    def get_Dayopen(self, day_str, direct):
        # 指定日期 和方向:1,-1 
        # 1:向後找到第一個開市日,輸出找到的日期
        #-1:向前找到第一個開市日,輸出找到的日期 
        # 找不到 輸出 '19000101'
        day00_str = '19000101'
        if direct not in (1,-1): 
            return day00_str
        sorted_dates = sorted(self.status.keys())
        if direct == -1:
            sorted_dates.reverse()
        for d in sorted_dates:
            if (direct == -1 and d < day_str) or (direct == 1 and d > day_str):
                if self.status.get(d, False):
                    return d
        return day00_str                                         
    
    def get_code(self, day_str): 
        # 以日期求得小台指(當月/近月)商品代碼, (day_str:日期)
        dd_str = day_str
        day0 = datetime.datetime.strptime(dd_str, '%Y%m%d')                 # 將 day_str 轉為日期格式 day0         
        day = day0.replace(day=1)                                           # 設為當月第1天                                       
        while day.weekday() != 2:                                           # 計算當月月結日
            day += datetime.timedelta(days=1)                               # 當月第一個週三
        day += datetime.timedelta(weeks=2)                                  # 當月第三個週三
        dd_str = day.strftime('%Y%m%d')                                     # 透過函數 判斷如果月結日是休市日，求得新的當月月結日 day
        if not self.get_status(day_str) :                                   # 判別 day是休市日
            day_str = self.get_Dayopen(day_str, 1)                          # 取得 下一個開市日
            day = datetime.datetime.strptime(dd_str, '%Y%m%d')
        if day0 > day:                                                      # 如果 day0 >月結日  則 day 設定為下一個月
            year, month = (day.year + 1, 1) if day.month == 12 else (day.year, day.month + 1)       # 處理跨年問題
        else:
            year, month =(day.year,day.month)     
        cod0 = f"MXF{chr(month + 64)}{year%10}"                             # 商品代號 (小台指-當月)
        return cod0                                                         

    def download_fexdata(self):
        # 下載期交所 40天內 每日壓縮數據檔 (30天開市日)  
        s_date = datetime.date.today() - datetime.timedelta(days=40)        # 起始日期為為40天前
        e_date = datetime.date.today() - datetime.timedelta(days=1)         # 結束日期為昨天                                                           # 生成日期序列
        cur_date = s_date
        while cur_date <= e_date:
            d0_str=cur_date.strftime("%Y%m%d")
            if self.get_status(d0_str):                                     # 開市日
                fname=os.path.join(self.path_tfex, f"{d0_str}Fex.zip")
                if not os.path.isfile(fname):
                    d = datetime.datetime.strptime(d0_str, '%Y%m%d').date()                      
                    if (datetime.date.today() - d) <= datetime.timedelta(days=40):  # 期交所(上架30天內開市日)數據檔 
                        url0 = "Daily_" + d0_str[0:4] + "_" + d0_str[4:6] + "_" + d0_str[6:] + ".zip"
                        url = f"https://www.taifex.com.tw/file/taifex/Dailydownload/Dailydownload/{url0}" 
                        download_file(url, fname, retries=3, wait_time=1, min_size=1024)                    
            cur_date += datetime.timedelta(days=1)

# ------------------------------------------------------------------------------------------------------
def get_time_period(t_str=''):
    # 取得 時間段編號  -1:無法辨識; 0:日盤前; 1:日盤; 2:夜盤前; 3:夜盤 
    try:
        if t_str=='':
            time_obj =datetime.datetime.now().time()
        else:
            t=('000000' + t_str[:-6])
            time_obj = datetime.datetime.strptime(t[-6:], "%H%M%S").time()
    except ValueError:
        return -1
    for period, time_ranges in trading_hours.items():
        for start_time, end_time in time_ranges:
            if start_time <= time_obj < end_time:
                return period
    return -1

def download_file(url: str, fname: str, retries: int = 3, wait_time: int = 1, min_size: int = 1024) -> bool:
    # 下載檔案，如果下載失敗會進行重試，直到下載成功或達到重試次數為止。
    for i in range(retries):
        response = requests.get(url)
        if response.status_code == 200 and len(response.content) >= min_size:
            with open(fname, "wb") as f:
                f.write(response.content)
            return True
        time.sleep(wait_time)
    return False

def create_table(db1, tab1, sql1="(date01 CHAR(8) PRIMARY KEY, open01 INTEGER)"):
    with sqlite3.connect(db1) as conn:
        c = conn.cursor()
        sql_str =f"CREATE TABLE IF NOT EXISTS {tab1} {sql1} " 
        c.execute(sql_str)

def insert_data(db1, tab1, data0, sql1="(date01, open01) VALUES (?, ?)"):
    with sqlite3.connect(db1) as conn:
        c = conn.cursor()
        sql_str =f"INSERT INTO {tab1} {sql1}"
        c.executemany(sql_str, data0)

def execute_query(db1, query):
    with sqlite3.connect(db1) as conn:
        c = conn.cursor()
        c.execute(query)
        rows = c.fetchall()
    return rows