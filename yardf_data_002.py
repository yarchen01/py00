import pandas as pd
import numpy as np
from datetime import datetime, timedelta    
class FDataFrame:
    def __init__(self):
        today = datetime.now().date()     
        start_time = datetime.strptime('08:45', '%H:%M')
        end_time = datetime.strptime('13:45', '%H:%M')
        start_datetime = pd.Timestamp.combine(today, start_time.time())
        end_datetime = pd.Timestamp.combine(today, end_time.time())        
        timestamps = pd.date_range(start_datetime, end_datetime, freq='T')
        columns = ['Open', 'High', 'Low', 'Close', 'Vol', 'm20', 'm60', 'std', 'ub', 'lb']
        self.df = pd.DataFrame({'Time': pd.to_datetime(timestamps)}).assign(**{col: [np.nan] * len(timestamps) for col in columns})
        self.df.set_index('Time', inplace=True)
        self.xO = self.df.columns.get_loc('Open')
        self.xH = self.df.columns.get_loc('High')
        self.xL = self.df.columns.get_loc('Low')
        self.xC = self.df.columns.get_loc('Close')
        self.xV = self.df.columns.get_loc('Vol')
        self.x2 = self.df.columns.get_loc('m20')
        self.x6 = self.df.columns.get_loc('m60')
        self.xS = self.df.columns.get_loc('std')
        self.xU = self.df.columns.get_loc('ub')
        self.xD = self.df.columns.get_loc('lb')
        self.sec0 = '084400'                    # 前一筆時標 str(hhmmss)
        self.ix = -1 

    def update_df(self, t:str, pri0:int, qty0:int):
        # 數據更新處理  t:時標(6位字串 hhmmss) pri0:成交價(9999.9)  qty0:成交單量(9999)   
        df = self.df
        ix = self.ix
        tm=t[2:4]
        if tm == self.sec0[2:4]:                # 同一分鐘
            if pri0 > df.iat[ix,self.xH]:
                df.iat[ix,self.xH] = pri0
            if pri0 <df.iat[ix,self.xL]:
                df.iat[ix,self.xL] = pri0
            df.iat[ix,self.xC] = pri0
            df.iat[ix,self.xV] += qty0           
            if t != self.sec0:                  # 同一分不同秒
                self._update_indicators()
        else:                                   # 新一分鐘
            self.ix += 1
            df.iloc[self.ix, [self.xO, self.xH, self.xL, self.xC, self.xV]] = [pri0, pri0, pri0, pri0, qty0]
            self._update_indicators()
        self.sec0 = t                           # 更新 sec0 為當前的時標

    def _update_indicators(self):
        ix=self.ix
        if ix >= 20:
            df = self.df
            cs = df['Close'].dropna().values
            if ix >= 60: 
                rolling_m60 = pd.Series(cs).rolling(window=60).mean().round(2)
                df.iat[ix, self.x6] = rolling_m60.iloc[-1] 
            rolling_m20 = pd.Series(cs).rolling(window=20).mean().round(2)
            rolling_std  = pd.Series(cs).rolling(window=20).std().round(2)
            xm20 = rolling_m20.iloc[-1]
            xstd = rolling_std.iloc[-1] 
            df.iat[ix, self.x2] = xm20                                # 移動平均線數據                                     # 移動平均線數據
            df.iat[ix, self.xS] = xstd                                # 標準差數據
            df.iat[ix, self.xU] = xm20 + 2 * xstd                      # 上軌數據
            df.iat[ix, self.xD] = xm20 - 2 * xstd                      # 下軌數據    
