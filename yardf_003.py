# 名稱:         yardf_003.py 
# 目的:         測試 Frame_input    
# 版本:         20230609-A03
# ------------------------------------------------------------------------------------------------------ 
import os
import pandas as pd
import numpy as np
import wx
import wx.grid
import matplotlib.pyplot as plt
from   matplotlib.backends.backend_wxagg import FigureCanvasWxAgg
import mplfinance as mpf
import matplotlib.gridspec as gridspec
import threading
import struct
import base64
import datetime
import time
import yar_market
import yar_api
market_colors = mpf.make_marketcolors(
    up='r',  # 涨的颜色
    down='g',  # 跌的颜色
    edge='inherit',
    wick='inherit',
    volume=(135/255, 206/255, 250/255)
)
style = mpf.make_mpf_style(
    marketcolors=market_colors
)
api_flag = False
# ------------------------------------------------------------------------------------------------------ 
def encrypt_string(string):
    encoded_bytes = base64.b64encode(string.encode('utf-8'))
    encrypted_string = encoded_bytes.decode('utf-8')
    return encrypted_string

def decrypt_string(encrypted_string):
    decoded_bytes = base64.b64decode(encrypted_string.encode('utf-8'))
    decrypted_string = decoded_bytes.decode('utf-8')
    return decrypted_string

def set_path1(fname0:str) -> tuple[str, str, str, str]:
    # 從檔案取得路徑資料 設定數據存放全路徑名稱
    if not os.path.isfile(fname0):
        return '','','','',''
    with open(fname0, "r") as file:
        data1 = os.path.abspath(file.read())
    path0=data1.split('\n')[0]
    if not os.path.exists(path0[:2]):                           # 檢查磁碟機是否存在 
        path0='C:\\'              
    path_root = os.path.join(path0, 'data00')                   # 數據主目錄  
    path_yuan = os.path.join(path_root, 'yuanta')               # 元大數據下載路徑
    path_tfex = os.path.join(path_root, 'taifex')               # 期交所數據下載路徑
    db0 = os.path.join(path_root, "fex00.db")                   # 數據庫 (存放 期交所-逐筆數據)
    db1 = os.path.join(path_root, "fex01.db")                   # 數據庫 (存放 期交所-每秒/每分/日夜/彙整數據)
    os.makedirs(path_root, exist_ok=True)
    os.makedirs(path_yuan, exist_ok=True)
    os.makedirs(path_tfex, exist_ok=True)
    return path_yuan, path_tfex, db0, db1

def RGB(c1:int,c2:int,c3:int) -> tuple:
    # 轉換 RBG
    return (c1/255,c2/255,c3/255)

# ------------------------------------------------------------------------------------------------------ 
class FDataFrame:
    def __init__(self):
        today = datetime.datetime.now().date()     
        start_time = datetime.datetime.strptime('08:45', '%H:%M')
        end_time = datetime.datetime.strptime('13:45', '%H:%M')
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
        self.w20 = []
        self.w60 = []
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
                if ix >=20:
                    self.w20[-1] = pri0
                if ix >=60:
                    self.w60[-1] = pri0
                self._update_indicators()
        else:                                   # 新一分鐘
            self.ix += 1
            df.iloc[self.ix, [self.xO, self.xH, self.xL, self.xC, self.xV]] = [pri0, pri0, pri0, pri0, qty0]
            self._update_indicators()
            self.w20.append(pri0)
            if len(self.w20) > 20:
                self.w20.pop(0)
            self.w60.append(pri0)
            if len(self.w60) > 60:
                self.w60.pop(0)
        self.sec0 = t                           # 更新 sec0 為當前的時標

    def _update_indicators(self):
        ix=self.ix
        if ix >= 20:
            df = self.df
            cs = df['Close'].dropna().values
            if ix >= 60:
                mean60 = sum(self.w60) / len(self.w60)
                df.iat[ix, self.x6] = round(mean60, 2)    
                #rolling_m60 = pd.Series(cs).rolling(window=60).mean().round(2)
                #df.iat[ix, self.x6] = rolling_m60.iloc[-1]
            mean20 = sum(self.w20) / len(self.w20)
            std20 = (sum((xi - mean20) ** 2 for xi in self.w20) / 20) ** 0.5
            #rolling_m20 = pd.Series(cs).rolling(window=20).mean().round(2)
            #rolling_std  = pd.Series(cs).rolling(window=20).std().round(2)
            df.iat[ix, self.x2] = round(mean20, 2)                      # 移動平均線數據 
            df.iat[ix, self.xS] = round(std20, 2)                       # 標準差數據
            df.iat[ix, self.xU] = round(mean20 + 2 * std20, 2)          # BBand 上軌數據
            df.iat[ix, self.xD] = round(mean20 - 2 * std20, 2)          # BBand 下軌數據
  
# ------------------------------------------------------------------------------------------------------ 
class Dialog_input(wx.Dialog):
    def __init__(self, parent, size, pos, fname,  *args, **kw):
        # 設定畫面 (數據路徑 與 帳號登入) 
        style = wx.DEFAULT_FRAME_STYLE & ~(wx.CLOSE_BOX)                         # 移除關閉按鈕
        super(Dialog_input, self).__init__(parent, *args, **kw, style=style)                                                                                            
        self.SetTitle('帳號/密碼 登入表單   (F001)')                                                          
        self.fname=fname
        self.SetSize(size)                               
        self.SetPosition(pos)                           
        #--------------------------------------------------------------------------設定 Box1
        pnl = wx.Panel(self)
        w1=370; w2=120;       w3=45;      h0=20;     
        y0=15;  y1=40;        y2=70;     
        x0=31;  x1=x0+w2+80;  x2=x1+180;  x4=x0+430;             
        wx.StaticBox( pnl, pos=(5, 1), size=(size[0]-25, size[1]-50))
        wx.StaticText(pnl, pos=(x0,  y0), label='系統訊息:')   
        self.msg2 = wx.TextCtrl(pnl, pos=(x0+55, y0), size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)
        wx.StaticText(pnl, pos=(x0,  y1), label='數據路徑:')
        self.path = wx.TextCtrl(pnl, pos=(x0+55, y1), size=(w1, h0))
        self.path.SetValue('C:\\')             
        wx.StaticText(pnl, pos=(x0,  y2), label='期貨帳號:')
        self.acc = wx.TextCtrl(pnl, pos=(x0+55, y2), size=(w2, h0))        
        wx.StaticText(pnl, pos=(x1,  y2), label='密碼:')
        self.pwd = wx.TextCtrl(pnl, pos=(x1+30, y2), size=(w2, h0),style=wx.TE_PASSWORD )       
        self.btn = wx.Button(pnl, wx.ID_ANY, label='登入', pos=(x2, y2), size=(w3, h0))
        self.ext = wx.Button(pnl, wx.ID_ANY, label='結束', pos=(x4, y0), size=(w3, h0))  
        self.msg2.SetValue(' 請輸入帳號/密碼,登入系統... ')
        self.btn.Bind(wx.EVT_BUTTON, self.logon)
        self.ext.Bind(wx.EVT_BUTTON, self.on_exit)
        self.ShowModal()

    def logon(self,event=None):
        # 帳號密碼登入   
        acc0 = self.acc.GetValue() if self.acc != None else ''
        pwd0 = self.pwd.GetValue() if self.pwd != None else ''
        path0 = self.path.GetValue() if self.path != None else 'C:\\'
        path0 = os.path.abspath(path0)
        if acc0  != '' and  pwd0  != '' :
            data0 = path0 + '\n' + acc0 + '\n' + encrypt_string(pwd0)
            with open(self.fname, 'w') as file:
                file.write(data0)
            self.Destroy()
        else:
            self.msg2.SetValue(' 帳號/密碼不可為空, 請再次輸入帳號/密碼,登入系統... ')    
        
    def on_exit(self, event):
        # 結束退出 
        self.Destroy()      
# ------------------------------------------------------------------------------------------------------ 
class Frame_display(wx.Frame):
    def __init__(self, parent, size, pos, fname, market,  *args, **kw):
        # 設定畫面 (數據顯示) 
        style = wx.DEFAULT_FRAME_STYLE & ~(wx.CLOSE_BOX)                         # 移除關閉按鈕
        super(Frame_display, self).__init__(parent, *args, **kw, style=style)                                        
        self.parent = parent                                                     # 父層元件必須是頂層視窗，才能正常接收事件        
        self.fname=fname
        self.market=market
        self.SetTitle('台指期-小台指 整合系統   (F000)')
        self.SetSize(size)                               
        self.SetPosition(pos)                      
        #------------------------------------------------------------------------------設定 Box1
        pnl = wx.Panel(self)
        self._box1(pnl)                                             # 設定畫面第一部分(box-訊息顯示) 
        self._box2(pnl)                                             # 設定畫面第二部分(grid-即時數據)
        self.pre_clean_init()
        self._box3(pnl)                                             # 設定畫面第三部分(box/plot 繪圖)
        self.Show()
        if not api_flag:
            thread = threading.Thread(target=self.df_from_file)
            thread.start()

    def df_from_file(self):
        fname='.\\20230602-MXFF3.bin0'
        with open(fname, 'rb') as file:
            data0 = file.read()
        record_size = struct.calcsize('iifiififi')
        records = [struct.unpack('iifiififi', data0[i:i+record_size]) for i in range(0, len(data0), record_size)]
        xrecords = [record for record in records if 84500000 <= record[1] <= 134500000]
        self.Symbol0=fname.split('-')[1][:5]
        self.grid.SetCellValue(0, 0, self.Symbol0)
        pri0=xrecords[0][2]
        self.prih=pri0;   self.pril=pri0 
        counter = 0;      counters = len(xrecords)       
        for record in xrecords:
            day2, tim2, pri2, qty2, qty9, prib, qtyb, pris, qtys = record
            self.prih = pri2 if self.prih < pri2 else self.prih
            self.pril = pri2 if self.pril > pri2 else self.pril
            t=('000000000'+str(tim2))[-9:][:6]                       # 時標
            DF.update_df(t,pri2,qty2)                                # 數據處理
            counter += 1
            if (counter % 5000 == 0) or (counter==counters):
                self.msg2.SetValue(f'數據處理中: {counter}/{counters}')
                self.UpdateSymbol(str(pri0),str(pri0),str(self.prih),str(self.pril),t,str(pri2),str(qty2),str(qty9))  # 數據顯示 
                self._update_chart() 

    def _box1(self,pnl):                   
        x1=31;   y1=15;   w1=400;   w4=45;  h0=20;                
        bx1=wx.StaticBox( pnl, pos=(5, 1), size=(975, 150))
        wx.StaticText(pnl, pos=(x1,  y1), label='連線訊息:')
        self.msg1 = wx.TextCtrl(pnl, pos=(x1+55, y1),    size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)     
        wx.StaticText(pnl, pos=(x1,  y1+25), label='系統訊息:')   
        self.msg2 = wx.TextCtrl(pnl, pos=(x1+55, y1+25), size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)
        wx.StaticText(pnl, pos=(x1+465,  y1), label='路徑:')
        self.path = wx.TextCtrl(pnl, pos=(x1+495, y1),   size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)
        with open(self.fname, "r") as file:
            data1 = os.path.abspath(file.read())
        self.path.SetValue(data1.split('\n')[0])
        self.ext = wx.Button(pnl, wx.ID_ANY, label='結束', pos=(x1+900, y1), size=(w4, h0))  
        self.ext.Bind(wx.EVT_BUTTON, self.on_exit)
        self.Bind(wx.EVT_CLOSE, self.on_timeclose)
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.change_datetime, self.timer)
        self.timer.Start(1000)  # 每隔1秒觸發一次定時器事件  

    def _box2(self,pnl):
        labx=("商品代號","參考價","開盤價","最高價","最低價","成交時間","成交價","單量","總量","Miss")
        rows = 1;     cols = 10;     cell_h = 28;        cell_w = 80
        x2=31;        y2=70;         w2=68;  
        self.grid = wx.grid.Grid(pnl,pos=(x2,y2),size=(883,61))
        self.grid.CreateGrid(rows, cols)
        self.grid.SetDefaultRowSize(cell_h, True)                                                       # 設定預設行高度，並調整現有行的大小
        self.grid.SetDefaultColSize(cell_w, True)                                                       # 設定預設列寬度，並調整現有列的大小                       
        for index, sym in enumerate(labx):                                                              # 設置行標籤內容
            self.grid.SetColLabelValue(index, sym)                                                                  
        self.grid.SetLabelFont(wx.Font(wx.FontInfo(12).Family(wx.FONTFAMILY_DEFAULT).FaceName("標楷體")))
        self.date0=datetime.date.today().strftime("%Y%m%d")
        self.dateL = wx.StaticText(pnl, pos=(x2+1, y2+7), size=(w2, cell_h-5), style=wx.ALIGN_CENTER)  # 設置日期欄位 (角落)
        self.dateL.SetFont(wx.Font(wx.FontInfo(11).Family(wx.FONTFAMILY_DEFAULT).FaceName("Arial Black")))
        self.dateL.SetForegroundColour(wx.BLUE)
        self.dateL.SetLabel(self.date0)
        self.timeL = wx.StaticText(pnl, pos=(x2+2, y2+35), size=(w2, cell_h-5), style=wx.ALIGN_CENTER)  # 設置日期欄位 (RowLabel 0)
        self.timeL.SetFont(wx.Font(wx.FontInfo(12).Family(wx.FONTFAMILY_DEFAULT).FaceName("Arial Black")))
        self.timeL.SetForegroundColour(wx.BLUE)   
        self.grid.SetDefaultCellAlignment(wx.ALIGN_CENTRE, wx.ALIGN_CENTRE)                             # 設置所有儲存格 位置置中
        self.grid.SetDefaultCellTextColour(wx.WHITE)                                                    # 設置所有儲存格 前景顏色為 白
        self.grid.SetDefaultCellBackgroundColour(wx.BLACK)                                              # 設置所有儲存格 背景顏色為 黑
        font0 = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, faceName='Arial')
        self.grid.SetDefaultCellFont(font0)  
        self.grid.SetDefaultCellAlignment(wx.ALIGN_RIGHT, wx.ALIGN_CENTER)
        for i in range(cols):            
            if i == 0:                                                                                  # 設置 商品代號 位置置中  (第0行儲存格()
                attr = wx.grid.GridCellAttr()
                attr.SetAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
                self.grid.SetColAttr(i, attr)        
        self.grid.EnableEditing(False)

    def _box3(self, pnl):
        #DF.df = pd.read_csv('stock_data3.csv')
        bx3 = wx.StaticBox(pnl, pos=(5, 155), size=(1775, 690))
        self.fig = plt.Figure(figsize=(17.75, 6.9))                               # 調整子圖寬度，高度
        self.fig.subplots_adjust(left=0.05, bottom=0.01, right=0.95, top=0.90,hspace=0.01)     # 調整子圖的位置和間距
        self.canvas = FigureCanvasWxAgg(bx3, -1, self.fig)
        sizer = wx.StaticBoxSizer(bx3, wx.VERTICAL)
        sizer.Add(self.canvas, 1, wx.EXPAND)
        bx3.SetSizer(sizer)
        gs = gridspec.GridSpec(2, 1, height_ratios=[4, 0])
        self.ax1 = self.fig.add_subplot(gs[0])
        self.ax1.margins(x=0.001)
        self.ax1.grid(True, which='both', linestyle='--', color='gray')
        xx = DF.df.index.strftime('%H:%M')
        x_ticks = range(0, len(xx), 15)
        self.ax1.set_xticks(x_ticks)
        self.ax1.set_xticklabels(xx[x_ticks])
        self.ax1.xaxis.set_ticks_position('top')
        self.ax2 = self.ax1.twinx()
        self.vop = mpf.make_addplot(DF.df.iloc[:, DF.xV], type='bar',  ax=self.ax2, width=0.6,color=RGB(0, 137, 210),alpha=0.3) 
        self.m2p = mpf.make_addplot(DF.df.iloc[:, DF.x2], type='line', ax=self.ax1, width=1,  color='blue')        
        self.m6p = mpf.make_addplot(DF.df.iloc[:, DF.x6], type='line', ax=self.ax1, width=1,  color='red')
        self.ubp = mpf.make_addplot(DF.df.iloc[:, DF.xU], type='line', ax=self.ax1, width=1,  color='blue', linestyle='dashed')
        self.lbp = mpf.make_addplot(DF.df.iloc[:, DF.xD], type='line', ax=self.ax1, width=1,  color='blue', linestyle='dashed')

    def _update_chart(self):
        mpf.plot(DF.df, type='candle', ax=self.ax1,addplot=[self.vop,self.m2p,self.m6p,self.ubp,self.lbp], style=style)
        self.canvas.draw()
    
    def pre_clean_init(self):
        # 清盤
        self.Symbol0= ''                                                         # 商品代號(小台指-當月)
        self.t00    = '000000'                                                   # API時標(秒)-控制訊息顯示 
        self.time0  = datetime.datetime.now().strftime('%H:%M:%S')               # 當前時間
        self.date0  = datetime.date.today().strftime("%Y%m%d")                   # 當前日期 
        self.t1     = ''                                                         # 即時時標紀錄器
        self.qty    = 0                                                          # 總量紀錄器 
        self.err0   = 0                                                          # 錯誤紀錄器
        self.miss   = 0                                                          # 漏失紀錄器        
        self.fpath  = ''                                                         # 每日數據檔全路徑
        self.t_period   = self.market.get_time_period()                          # 旗標(時間段): -1錯,0日前,1日,2夜前,3夜_1,4,夜_2
        self.f_logon    = False                                                  # 旗標(已登入):
        self.f_download = False                                                  # 旗標(已開始繪圖):
        self.f_plot     = False        
        dd_str = self.market.get_Dayopen(self.date0,1) if (self.market.get_time_period() in (2,3))  else self.date0   # 調整數據日期 為下一開市日
        self.Symbol0 = self.market.get_code(dd_str)              # 商品代號(小台指-當月) 
        self.grid.SetCellValue(0, 0, self.Symbol0)

    def _logon(self):
        # Qapi 登入
        self.pre_clean_init()
        if  not api_flag:
            self.msg2.SetValue('從檔案倒入數據')
            self.f_logon=True 
            return
        self.YTQ = yar_api.YuantaQuoteAXCtrl(self)              # 執行接口連接的相關操作
        time.sleep(1)
        self.msg1.SetValue(' 與元大伺服器網路 連線中... ')
        with open(self.fname, 'r') as file:
            data1 = os.path.abspath(file.read())
        acc =data1.split('\n')[1]
        pwd =data1.split('\n')[2]
        port , reqType  = ('80', 1) if (self.t_period in [0, 1]) else ('82',2)               
        self.YTQ.api_Logon(acc, decrypt_string(pwd), port, reqType)
        self.msg2.SetValue('')
        self.f_logon=True                            

    def _register(self):
        # Qapi 商品代號 註冊
        if  self.Symbol0 != '':
            reqType = True if (self.market.get_time_period()  in [0, 1]) else False
            self.YTQ.api_Register(self.Symbol0, reqType)

    def _unregister(self):
        # Qapi 商品代號 註銷
        if  self.Symbol0 != '':   
            self.YTQ.api_UnRegister(self.Symbol0 )                                                
                         
    def UpdateSymbol(self, RefPri, OpenPri, HighPri, LowPri, t, MatchPri, MatchQty, TolMatchQty):
        # 顯示即時數據
        self.grid.SetCellValue(0, 1,RefPri)                                         # 參考價
        self.grid.SetCellTextColour(0, 1,wx.WHITE)              
        self.grid.SetCellValue(0, 2, OpenPri)                                       # 開盤價
        color = wx.WHITE if OpenPri == RefPri else wx.RED if OpenPri > RefPri else wx.GREEN
        self.grid.SetCellTextColour(0, 2, color)
        self.grid.SetCellValue(0, 3, HighPri)                                       # 最高價
        color = wx.WHITE if HighPri == RefPri else wx.RED if HighPri > RefPri else wx.GREEN
        self.grid.SetCellTextColour(0, 3, color)                
        self.grid.SetCellValue(0, 4, LowPri)                                        # 最低價
        color = wx.WHITE if LowPri == RefPri else wx.RED if LowPri > RefPri else wx.GREEN
        self.grid.SetCellTextColour(0, 4, color)                                 
        self.grid.SetCellValue(0, 5,  f"{t[0:2]}:{t[2:4]}:{t[4:6]}")                # 成交時間               
        self.grid.SetCellTextColour(0, 5, wx.WHITE)
        self.grid.SetCellValue(0, 6, MatchPri)                                      # 成交價
        color = wx.WHITE if MatchPri == RefPri else wx.RED if MatchPri > RefPri else wx.GREEN
        self.grid.SetCellTextColour(0, 6, color)                    
        self.grid.SetCellValue(0, 7, MatchQty)                                      # 單量         
        self.grid.SetCellTextColour(0, 7, wx.WHITE)                                       
        self.grid.SetCellValue(0, 8, TolMatchQty)                                   # 總量          
        self.grid.SetCellTextColour(0, 8, wx.WHITE)                       
        self.grid.SetCellValue(0, 9, str(self.miss))                                # 漏失次數        
        self.grid.SetCellTextColour(0, 9, wx.WHITE)
        if api_flag:
            self._update_chart()                                         

    def change_datetime(self, event): 
        # 顯示/更新 即時時間                                                             
        time1 = datetime.datetime.now().strftime('%H:%M:%S')
        if time1 != self.time0:                                         # 秒改變                          
            self.timeL.SetLabel(time1)                                  # 顯示/更新 即時時間
            if time1[:2] != self.time0[:2]:                             # 時改變
                date1 = datetime.date.today().strftime("%Y%m%d")
                if date1 != self.date0:                                 # 日改變
                    self.dateL.SetLabel(date1)                          # 顯示 即時日期
                    self.date0 = date1                                  # 更新 即時日期
            self.time0 = time1
            self.chktime_Process()                                      # 檢查時間 並做相應特殊處理
     
    def on_timeclose(self, event):
        self.timer.Stop()

    def chktime_Process(self):
        # 檢查時間 處理特殊作業
        tt = datetime.datetime.now().time()                 # 當前時間    
        self.t_period = self.market.get_time_period()       # 旗標(時間段): -1錯,0日前,1日,2夜前,3夜_1,4,夜_2
        # ---------------------------------------------------------------------------------------------------    
        if int(time.time()) % 60 == 0:                      # 每分鐘 顯示訊息
            if not self.market.get_status(self.date0):      # 休市日
                self.msg2.SetValue( self.date0 + ' 本日休市不開盤!' )
                return            
            day = datetime.date.today()
            if self.t_period == 1: 
                t0 = datetime.time(13, 45, 0) if not self.market.get_xday(self.date0) else datetime.time(13, 30, 0)
                msg0 = '日盤開盤中...    距收盤還有 '        
            elif self.t_period in (2,3,4,0):
                t0 = datetime.time(8, 45, 0)
                msg0 = '日盤已收盤!!!    距下次日盤開盤還有 '
            else:
                self.msg2.SetValue(' 時間段區隔錯誤,請洽系統人員處理!! ')
                return
            if self.t_period == 3:
                tt0 = datetime.datetime.combine(day+ datetime.timedelta(days=1), t0) - datetime.datetime.combine(day, tt)
            else:
                tt0 = datetime.datetime.combine(day, t0) - datetime.datetime.combine(day, tt)
            hh = tt0.seconds // 3600;    mm=tt0.seconds % 3600 
            self.msg2.SetValue(msg0 + f"{hh}小時  {(mm) // 60}分鐘!")
        # ----------------------------------------------------(13:46:00~13:46:10) 日盤收盤
        time_start = datetime.time(13, 46, 0) if not self.market.get_xday(self.date0) else datetime.time(13, 31, 0)
        time_end = time_start.replace(second=10)
        if self.market.get_status(self.date0) and time_start <= tt <= time_end:
            self._unregister()                                          # 商品代號註銷     
            self.YTQ = None                                             # 斷開接口連接的相關操作
            self.f_logon=False
            self.msg1.SetValue('元大API離線, 帳號已登出!')
            time.sleep(11)            
            return
        # ----------------------------------------------------(日盤)重新登入
        if (not self.f_logon) and self.market.get_status(self.date0):   # 未登入且為開市日
            if  datetime.time(8,40,0) <= tt <= datetime.time(13,30,00):    
                self._logon()                                           # (8:40:00~13:30:00)          
        # ----------------------------------------------------(6:30:00~~23:59:59) 執行 切換相關旗標
        if not self.f_download and tt >= datetime.time(6,30,0):
            self.market.download_fexdata()                              # 下載期交所前一日數據(自前30開市日至前一開市日)
            self.f_download = True                                      # 設定旗標 (已下載期交所數據)        

    def on_exit(self, event=None):
        # 結束退出 
        # app.keepGoing=False 
        self.Destroy()       
    # ----------------------------------------------------------------------------API主動回報 轉接函式(wrapper)               
    def OnMktStatusChange(self,Status):
        # API主動回傳 (狀態改變:連線/登入)
        if (Status == 1):
            self.msg1.SetValue(f" 與元大伺服器網路 已連線,  帳號登入中...")
        if (Status == 2):
            self.msg1.SetValue(f" 與元大伺服器網路已連線,  帳號登入成功")                   
            self._register()

    def OnGetTimePack(self,strTime):
        # API主動回傳 (時間校正)
        t=("000000000000"+strTime)[-12:][:9]
        t=f"{t[0:2]}:{t[2:4]}:{t[4:6]}.{t[6:9]}"
        self.msg2.SetValue(f" 登入元大伺服器, 時間校正 {t}")

    def OnGetMktData(self, PriType, symbol):
        # API主動回傳 (商品代號註冊)
        txt0 = '   '
        if PriType=='S':
            txt0 = '註冊中...'
        if PriType=='I':             
            txt0 = '註冊成功!'
        self.msg2.SetValue(f" 商品代號 {symbol} {txt0}")

    def OnGetMktAll(self,symbol, RefPri, OpenPri, HighPri, LowPri, UpPri, DnPri, MatchTime, MatchPri, MatchQty, TolMatchQty,BestBuyQty, BestBuyPri, BestSellQty,BestSellPri, FDBPri, FDBQty, FDSPri, FDSQty, ReqType):
        # API主動回傳 (已註冊商品代號的即時數據)       
        self.t_period = self.market.get_time_period()
        if self.t_period != 1:                                                    # 非開盤(日)時段
            return 
        if self.t_period != self.market.get_time_period(MatchTime) or self.Symbol0 != symbol:
            return                                                                # 時段或商品代號 不符合        
        i_tqty = int(TolMatchQty or 0)                                            # 總量
        i_qty = int(MatchQty or 0)                                                # 單量            
        if (i_tqty <=0) or (i_tqty <= self.qty) or (i_qty <=0):                   # 總量<=0 or 總量<=前總量 or 單量<=0
            return       
        if i_tqty != (self.qty+i_qty):                                            # (總量 != 前總量 + 單量)
            self.miss += 1                                                        # 漏失數據 
        self.qty = i_tqty                                                         # 回存總量
        # ----------------------------------------------------------------------------               
        t=("000000000000"+MatchTime)[-12:][:6]                                    # 時標
        DF.update_df(t,int(MatchPri),int(MatchQty))                               # 數據處理
        if int(t[:6]) - int(self.t00) >= 10:                                      # 每10秒顯示一次
            self.UpdateSymbol(RefPri,OpenPri,HighPri,LowPri,t,MatchPri,MatchQty,TolMatchQty)    # 數據顯示
            self.t00 = t[:6]
# ------------------------------------------------------------------------------------------------------ 
if __name__ == "__main__":
    fname0 = '.\\config.sys' 
    DF = FDataFrame()                                                               # 數據框架 實例化
    app = wx.App()
    if not os.path.isfile(fname0):                                                  # 檔案不存在
        Dialog1 = Dialog_input(None, size=(530, 150), pos=(400, 50), fname=fname0)  # 啟動登入畫面                                                 
    path_yuan, path_tfex, db0, db1 = set_path1(fname0)                              # 設定數據放置路徑
    market = yar_market.StockMarket(path_db=db1, path_tfex=path_tfex)               # 實例化 類別:StockMarket
    frame0 = Frame_display(None, size=(1800, 900), pos=(50, 50), fname=fname0, market=market)  # 啟動 主畫面
    app.MainLoop()
# ------------------------------------------------------------------------------------------------------ 
