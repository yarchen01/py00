# 程式名稱     小台指-當月 即時數據 錄製 
# 數據來源     元大期貨-行情API  (YuantaQuote_v2.1.2.9.ocx) (2022/11/29) 
# Python      (3.9.13-Win32)
# wxPython    (4.1.1)
# comtypes    (1.1.11)      無法解決POINTER(IUnknown)問題, 但暫時不影響程式執行
# version:    (20230519-b13)
import wx
import wx.grid
import ctypes
import comtypes
import comtypes.client
import datetime
import time
import os
import csv
import sqlite3
import requests
import struct
# --------------------------------------------------------------------------------------------------------------------------------
class AppFrame(wx.Frame):
    def __init__(self, parent, size, pos, *args, **kw):
        # Frame起始設定 (設置 程式畫面/元大Qapi連結/輸入與顯示處理) 
        style = wx.DEFAULT_FRAME_STYLE & ~(wx.CLOSE_BOX)                         # 移除關閉按鈕
        super(AppFrame, self).__init__(parent, *args, **kw, style=style)                                        
        self.parent = parent                                                     # 父層元件必須是頂層視窗，才能正常接收事件
        ver0 = '(Ver:20230519-b13)' 
        self.symbol = ''                                                         # 商品代號(小台指-當月) 
        self.t00    = ''                                                         # 時標紀錄器 
        self.t1     = ''                                                         # 即時時標紀錄器
        self.qty    = 0                                                          # 總量紀錄器 
        self.err0   = 0                                                          # 錯誤紀錄器
        self.miss   = 0                                                          # 漏失紀錄器
        self.fpath  = ''                                                         # 數據檔全路徑 
        self.f_download = False                                                  # 旗標: 期交所前一日數據已下載
        self.f_logon    = False                                                  # 旗標: 已執行登入        
        Title0 = f"小台指-當月即時數據錄製 {ver0}"
        self.SetTitle(Title0)                                                          
        self.Time_Period = { 
            "day-chk0":   (datetime.time(5,  0,  0), datetime.time(8,  45, 0)),
            "night-chk": (datetime.time(8,  45, 0), datetime.time(13, 45, 0))}
        self.GUIinit(size, pos)
        #--------------------------------------------------------------------------設定 YuantaQuoteCtr API 連接介面                                   
        Iwindow  = ctypes.POINTER(comtypes.IUnknown)()                           # 不影響程式執行  暫時忽略此錯誤
        Icontrol = ctypes.POINTER(comtypes.IUnknown)()
        Ievent   = ctypes.POINTER(comtypes.IUnknown)()
        ctypes.windll.atl.AtlAxCreateControlEx(
                                                "YUANTAQUOTE.YuantaQuoteCtrl.1",  
                                                self.Handle,  
                                                None,
                                                ctypes.byref(Iwindow),
                                                ctypes.byref(Icontrol),
                                                comtypes.byref(comtypes.GUID()),
                                                Ievent)
        self.YuantaQuote = comtypes.client.GetBestInterface(Icontrol)
        self.YuantaQuoteEvents = YuantaQuoteEvents(self)
        self.YuantaQuoteEventsConnect = comtypes.client.GetEvents(self.YuantaQuote, self.YuantaQuoteEvents)
    #---------------------------------------------------------------------------------------------------------         
    def GUIinit(self,size,pos):
        #--------------------------------------------------------------------------設定 Box1 帳號/密碼/登入
        self.SetSize(size)                               
        self.SetPosition(pos)                           
        pnl = wx.Panel(self)
        w0=55;  w1=400;       w2=120;     w3=370;     w4=45;     w5=68;      h0=20;     
        y0=15;  y1=40;        y2=70      
        x0=31;  x1=x0+w1+80;  x2=x1+180;  x3=x2+175; x4=x0+900             
        wx.StaticBox( pnl, pos=(5, 1), size=(size[0]-25, size[1]-50))
        wx.StaticText(pnl, pos=(x0,  y0), label='連線訊息:')
        self.msg = wx.TextCtrl(pnl, pos=(x0+55, y0), size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)
        wx.StaticText(pnl, pos=(x1,  y0), label='路徑:')
        self.path0 = wx.TextCtrl(pnl, pos=(x1+35, y0), size=(w3, h0))     
        wx.StaticText(pnl, pos=(x0,  y1), label='系統訊息:')   
        self.msg2= wx.TextCtrl(pnl, pos=(x0+55, y1), size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)     
        wx.StaticText(pnl, pos=(x1,  y1), label='帳號:')
        self.acc = wx.TextCtrl(pnl, pos=(x1+35, y1), size=(w2, h0))        
        wx.StaticText(pnl, pos=(x2,  y1), label='密碼:')
        self.pwd = wx.TextCtrl(pnl, pos=(x2+35, y1), size=(w2, h0),style=wx.TE_PASSWORD )       
        self.btn = wx.Button(pnl, wx.ID_ANY, label='登入', pos=(x3, y1), size=(w4, h0))
        self.ext = wx.Button(pnl, wx.ID_ANY, label='結束', pos=(x4, y0), size=(w4, h0))  
        self.btn.Bind(wx.EVT_BUTTON, self.logon)
        self.ext.Bind(wx.EVT_BUTTON, self.on_exit)
        self.Bind(wx.EVT_CLOSE, self.on_timeclose)
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.update_time, self.timer)
        self.timer.Start(1000)  # 每隔1秒觸發一次定時器事件       
        #--------------------------------------------------------------------------Box2 顯示錄製商品代號與數據訊息 (wxGrid)
        labx=("商品代號","參考價","開盤價","最高價","最低價","成交時間","成交價","單量","總量","Miss")
        rows = 1;           cols = 10;       
        cell_h = 28;        cell_w = 80
        self.grid = wx.grid.Grid(pnl,pos=(x0,y2),size=(883,61))
        self.grid.CreateGrid(rows, cols)
        self.grid.SetDefaultRowSize(cell_h, True)                                                       # 設定預設行高度，並調整現有行的大小
        self.grid.SetDefaultColSize(cell_w, True)                                                       # 設定預設列寬度，並調整現有列的大小                       
        for index, sym in enumerate(labx):                                                              # 設置行標籤內容
            self.grid.SetColLabelValue(index, sym)                                                                  
        self.grid.SetLabelFont(wx.Font(wx.FontInfo(12).Family(wx.FONTFAMILY_DEFAULT).FaceName("標楷體")))
        self.date0 = wx.StaticText(pnl, pos=(x0+1, y2+7), size=(w5, cell_h-5), style=wx.ALIGN_CENTER)     # 設置日期欄位 (角落)
        self.date0.SetFont(wx.Font(wx.FontInfo(11).Family(wx.FONTFAMILY_DEFAULT).FaceName("Arial Black")))
        self.date0.SetForegroundColour(wx.BLUE)
        self.date0.SetLabel(datetime.date.today().strftime("%Y%m%d")) 
        self.time0 = wx.StaticText(pnl, pos=(x0+2, y2+35), size=(w5, cell_h-5), style=wx.ALIGN_CENTER)    # 設置日期欄位 (RowLabel 0)
        self.time0.SetFont(wx.Font(wx.FontInfo(12).Family(wx.FONTFAMILY_DEFAULT).FaceName("Arial Black")))
        self.time0.SetForegroundColour(wx.BLUE)   
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
            elif i == 5:                                                                                # 設置 成交時間 字形大小=12(第5行儲存格()
                attr = wx.grid.GridCellAttr()
                attr.SetFont(wx.Font(wx.FontInfo(11)))
                self.grid.SetColAttr(i, attr)          
        self.grid.EnableEditing(False)        
           
    #---------------------------------------------------------------------------------------------------------
    def logon(self,event):
        global market
        # Qapi 帳號密碼登入
        # T port 80/443 , T+1 port 82/442 ,  reqType=1 T盤 , reqType=2  T+1盤
        acc0  = self.acc.GetValue()
        pwd0  = self.pwd.GetValue()
        port   = ('80') if (get_time_period() in [0, 1]) else ('82')
        reqType= (1)    if (get_time_period() in [0, 1]) else (2)
        frame.msg.SetValue(f" 與元大伺服器網路 連線中... ")
        self.YuantaQuote.SetMktLogon(acc0, pwd0, 'apiquote.yuantafutures.com.tw', port, reqType, 0)
        self.f_logon  = True
        # ----------------------------------------------------------------------------------------登入後才能做的事情
        path0 = self.path0.GetValue()                                                            # 設定數據放置路徑
        if path0 != '' :
            global path_yuan, path_tfex, db0, db1 
            path_yuan, path_tfex, db0, db1 = set_path0(path0)                                    
        market = StockMarket(1)                                                                  # 設定StockMarket類別-提供開市日等相關訊息功能	      
        # print ('login')

    def RegisterQuoteSymbol(self, Symbol0):
        # Qapi 商品代號 註冊
        # QuoteSymbol: 商品代號;  
        # reqType:     (True:日盤, False:夜盤)
        # UpdateMode:  (1:Snapshot當時最新資料, 2:Update註冊後的所有更新, 4:SnapshotUpd Snapshot + Update) 
        UpdateMode = 4 
        reqType = True if (get_time_period() in [0, 1]) else False            
        self.YuantaQuote.AddMktReg (Symbol0, UpdateMode , reqType, 0)

    def UnRegisterQuoteSymbol (self, Symbol0):
        # Qapi 商品代號 註消
        self.YuantaQuote.DelMktReg (Symbol0, 0)
        time.sleep(1)
        self.YuantaQuote.DelMktReg (Symbol0, 1)
        time.sleep(1) 
                 
    def UpdateSymbol(self, sym, RefPri, OpenPri, HighPri, LowPri, MatchTime, MatchPri, MatchQty, TolMatchQty, BestBuyPri, BestBuyQty, BestSellPri, BestSellQty):
        # 更新處理 Qapi 取得的即時數據 
        global buf
        if sym == self.symbol:
                i_Tqty=int(TolMatchQty)                                                     # 總量
                i_Sqty =int(MatchQty)                                                       # 單量
                if i_Tqty != (self.qty+i_Sqty):                                             # 漏失 (總量 != 前總量 + 單量)
                    self.miss += 1
                self.qty = i_Tqty                                                           # 回存總量
                # ----------------------------------------------------------------------------               
                day2 = int(self.date0.GetLabel()) if self.date0.GetLabel() != "" else 0
                tim2 = int(MatchTime) // 1000 if MatchTime != "" else 0
                pri2 = int(MatchPri) if MatchPri != "" else 0
                qty2 = int(MatchQty) if MatchQty != "" else 0
                qty9 = int(TolMatchQty) if TolMatchQty != "" else 0
                prib = int(BestBuyPri.split(",")[0])  if BestBuyPri  != "" else 0
                qtyb = int(BestBuyQty.split(",")[0])  if BestBuyQty  != "" else 0
                pris = int(BestSellPri.split(",")[0]) if BestSellPri != "" else 0
                qtys = int(BestSellQty.split(",")[0]) if BestSellQty != "" else 0                
                dd = struct.pack('iifiififi',day2,tim2,pri2,qty2,qty9,prib,qtyb,pris,qtys)
                buf.append(dd)                                                              # 存入  暫存數據陣列                   
                t=("000000000000"+MatchTime)[-12:][:9]                                      # 時標
                if  (self.t00 == t[0:6]):                                                   # 每秒顯示一次
                    return
                self.t00 = t[0:6]
                with open(self.fpath, 'ba') as file:
                    for i in range(len(buf)):
                        file.write(buf[i])
                buf=[]                                            
                self.t00 = t[0:6]                                                           # 回存時標
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
                self.grid.SetCellValue(0, 5,  f"{t[0:2]}:{t[2:4]}:{t[4:6]}.{t[6:9]}")       # 成交時間               
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
    
    def SetConnectStatusValue(self,Value):
        # 顯示連線登入訊息                                           
        self.msg.SetValue(Value)
                                                               
    def SetLogonDisable(self):
        # 失效(帳號/密碼/登入)
        self.path0.Disable()                                                      
        self.acc.Disable()
        self.pwd.Disable()
        self.btn.Disable()
   
    def SetRegister_Symbol(self):
        # 取得並註冊 (當月/近月 大/小台指) 商品代號 
        if  self.symbol != '':
            self.UnRegisterQuoteSymbol (self.symbol)                                        # 商品代號註銷       
        day_str = datetime.date.today().strftime("%Y%m%d")                                  # 當前日期
        if datetime.datetime.now().time() > datetime.time(13, 46, 0):                       # 調整日期 (夜盤視為 下一開市日之數據)
            day_str = market.get_Dayopen(day_str,1)                                      
        self.symbol = market.get_code(day_str)                                              # 商品代號(小台指-當月) (月結日下午 商品代號會不同)
        self.fpath = os.path.join(path_yuan, f"{day_str}-{self.symbol}.bin0")               # 數據的全路徑
        self.grid.SetCellValue(0, 0, self.symbol)          
        self.RegisterQuoteSymbol(self.symbol)
        time.sleep(1.5)        
    
    def update_time(self, event):
        # 顯示/更新 即時時間
        now = datetime.datetime.now()
        self.chktime_Process()                                                                  # 檢查時間 並做相應處理
        current_time = now.strftime('%H:%M:%S')
        if self.time0.GetLabel() != current_time:
            self.time0.SetLabel(current_time)
            self.update_date(now)
   
    def update_date(self, now):
        # 顯示/更新 即時日期
        current_date = now.strftime('%Y%m%d')
        if self.date0.GetLabel() != current_date:
            self.date0.SetLabel(current_date)
    
    def chktime_Process(self):
        #檢查時間 處理特殊作業
        if not self.f_logon:                                                # 未登入,不確定數據路徑
            return 
        # [8:30:00~8:30:30] and [14:50:00~14:50:30] 執行 檢查數據歸屬日期(影響數據檔案名稱和商品代號),重命名數據檔; 重註冊商品代號
        tt = datetime.datetime.now().time()
        if (datetime.time(8,30,0) <= tt <= datetime.time(8,30,30)) or (datetime.time(14,50,0) <= tt <= datetime.time(14,50,30)):
            self.SetRegister_Symbol()
            self.msg2.SetValue(self.symbol + ' 商品代號重新註冊,請稍等...')
            time.sleep(30)
            self.msg2.SetValue(self.symbol+ ' 商品代號 已完成註冊!!')    
        # [6:00:00~6:00:30] 執行 切換相關旗標
        if (datetime.time(6,00,0) <= tt <= datetime.time(6,00,30)):
            self.f_download = False                                         # 旗標: 期交所前一日數據已下載
        # [6:30:00~23:59:59] 執行 下載期交所前一日數據
        if self.f_logon and  not self.f_download:                           # 已登入 且 未下載
            if datetime.time(6,30,0) <= tt:         
                download_fexdata()                                          # 下載期交所前一日數據
                self.f_download = True
        
    def on_timeclose(self, event):
        self.timer.Stop()
                   
    def on_exit(self, event):
        # 結束退出 
        # app.keepGoing=False 
        self.Destroy()       
# --------------------------------------------------------------------------------------------------------------------------------
class YuantaQuoteEvents(object):
    # 元大Qapi的相關函數 (主動回傳)
    def __init__(self, parent):
        self.parent = parent
        self.TRADING_HOURS_MASK = 0b1101  # 定义时间范围的二进制掩码，每位代表不同的时间段
        self.TRADING_HOURS = {
            "day":   (datetime.time(8, 45,  0), datetime.time(13, 45, 0)),
            "night": (datetime.time(15, 0,  0), datetime.time(5,   0, 0))}
        
    def is_within_trading_hours(self, MatchTime):
        t0=("000000000000"+MatchTime)[-12:][:6]
        time_obj = datetime.datetime.strptime(t0, "%H%M%S").time()       
        if self.TRADING_HOURS["day"][0] <= time_obj <= self.TRADING_HOURS["day"][1]:            # 检查时间是否在交易时间范围内
            return bool(self.TRADING_HOURS_MASK & 0b0001)                                       # 检查二进制掩码的第一位
        elif self.TRADING_HOURS["night"][0] <= time_obj or time_obj <= self.TRADING_HOURS["night"][1]:
            return bool(self.TRADING_HOURS_MASK & 0b1000)                                       # 检查二进制掩码的第四位
        else:
            return False

    def OnMktStatusChange (self, this, Status, Msg, ReqType):
        msg0 = f"OnMktStatusChange {ReqType},{Msg},{Status}"
        # print (msg0)
        frame.SetConnectStatusValue(msg0)
        if (Status == 1):
            frame.msg.SetValue(f" 與元大伺服器網路 已連線,  帳號登入中...")
            frame.SetLogonDisable()
            frame.SetRegister_Symbol()
        if (Status == 2):
            frame.msg.SetValue(f" 與元大伺服器網路已連線,  帳號登入成功")
            frame.SetLogonDisable()
            frame.SetRegister_Symbol()

    def OnRegError(self, this, symbol, updmode, ErrCode, ReqType):
        print ('OnRegError {},{},{},{}'.format (ReqType, ErrCode, symbol, updmode))

    def OnGetMktData(self, this, PriType, symbol, Qty, Pri, ReqType):
        frame.msg2.SetValue(f" 商品代號 {symbol} 註冊中...")
        # print ('OnGetMktData', this, PriType, symbol, Qty, Pri, ReqType)

    def OnGetMktQuote(self, this, symbol, DisClosure, Duration, ReqType):
        print ('OnGetMktQuote')

    def OnGetMktAll(self, this, symbol, RefPri, OpenPri, HighPri, LowPri, UpPri, DnPri, MatchTime, MatchPri, MatchQty, TolMatchQty,BestBuyQty, BestBuyPri, BestSellQty,BestSellPri, FDBPri, FDBQty, FDSPri, FDSQty, ReqType):
        # API 即時數據
        if (not self.is_within_trading_hours(MatchTime)):                                   # 不在開盤時段
            return
        i_tqty = int(TolMatchQty) if TolMatchQty != '' else 0
        i_qty  = int(MatchQty)    if MatchQty    != '' else 0
        if (i_tqty <=0) or (i_tqty <= frame.qty) or (int(MatchQty) <=0):                    # 總量<=0 or 總量<=前總量 or 單量<=0
            return       
        frame.UpdateSymbol(symbol,RefPri,OpenPri,HighPri,LowPri,MatchTime,MatchPri,MatchQty,TolMatchQty,BestBuyPri,BestBuyQty,BestSellPri,BestSellQty)
        # print ('OnGetMktAll')

    def OnGetDelayClose(self, this, symbol, DelayClose, ReqType):
        print ('OnGetDelayClose')

    def OnGetBreakResume(self, this, symbol, BreakTime, ResumeTime, ReqType):
        print ('OnGetBreakResume')

    def OnGetTradeStatus(self, this, symbol, TradeStatus, ReqType):
        print ('OnGetTradeStatus')

    def OnTickRegError(self, this, strSymbol, lMode, lErrCode, ReqType):
        print ('OnTickRegError')

    def OnGetTickData(self, this, strSymbol, strTickSn, strMatchTime, strBuyPri, strSellPri, strMatchPri, strMatchQty, strTolMatQty,
        strMatchAmt, strTolMatAmt, ReqType):
        print ('OnGetTickData')

    def OnTickRangeDataError(self, this, strSymbol, lErrCode, ReqType):
        print ('OnTickRangeDataError')

    def OnGetTickRangeData(self, this, strSymbol, strStartTime, strEndTime, strTolMatQty, strTolMatAmt, ReqType):
        print ('OnGetTickRangeData')

    def OnGetTimePack(self, this, strTradeType, strTime, ReqType):
        t=("000000000000"+strTime)[-12:][:9]
        t=f"{t[0:2]}:{t[2:4]}:{t[4:6]}.{t[6:9]}"
        frame.msg2.SetValue(f" 登入元大伺服器, 時間校正 {t}")
        # print ('OnGetTimePack {},{}'.format (strTradeType, strTime))

    def OnGetDelayOpen(self, this, symbol, DelayOpen, ReqType):
        print ('OnGetDelayOpen')

    def OnGetFutStatus(self, this, symbol, FunctionCode, BreakTime, StartTime, ReopenTime, ReqType):
        frame.msg2.SetValue(f" 商品代號 {symbol} 註冊成功")
        # print ('OnGetFutStatus', this, symbol, FunctionCode, BreakTime, StartTime, ReopenTime, ReqType)

    def OnGetLimitChange(self, this, symbol, FunctionCode, StatusTime, Level, ExpandType, ReqType):
        print ('OnGetLimitChange')        
# --------------------------------------------------------------------------------------------------------------------------------
class StockMarket:
    # 類別 近兩年每日開市狀態與方法
    def __init__(self, init=0):
        if init == 0:
            return
        self.T01_NAME = "t01_market"
        self.status = {}
        now = datetime.datetime.now()
        start_date = datetime.datetime(now.year - 1, 1, 1)
        end_date = datetime.datetime(now.year, 12, 31)
        delta = end_date - start_date 
        days = delta.days + 1        
        start_date0 = int(start_date.strftime('%Y%m%d'))
        end_date0 = int(end_date.strftime('%Y%m%d'))
        sqlstr = f"CREATE TABLE IF NOT EXISTS {self.T01_NAME} (date01 CHAR(8) PRIMARY KEY, open01 INTEGER)"
        with sqlite3.connect(db1) as conn:
            c = conn.cursor()
            c.execute(sqlstr)
            time.sleep(1)
            c = conn.cursor()
            sqlstr = f"SELECT date01, open01 FROM {self.T01_NAME} WHERE date01 BETWEEN {start_date0} AND {end_date0}"
            c.execute(sqlstr)
            rows = c.fetchall()
            if  len(rows) < days :
                self._t01_download_insert(now.year - 1)
                self._t01_download_insert(now.year)
                c.execute(sqlstr)
                rows = c.fetchall()              
            for row in rows:
                self.status[row[0]] = (row[1] == 1)

    def _t01_download_insert(self,year1: int):
        # 下載指定年的每日開市資料,並新增入t01數據表
        url = f"https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear={year1}"
        fname_csv = os.path.join(path_tfex, f"holidaySchedule_{year1}.csv")
        if  not os.path.isfile(fname_csv):
            for i in range(3):
                response = requests.get(url)
                if response.status_code == 200 and len(response.content) >= 1024:
                    with open(fname_csv, "wb") as f:
                        f.write(response.content)            
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
        with sqlite3.connect(db1) as conn:
            # conn.execute(f"CREATE TABLE IF NOT EXISTS {self.T01_NAME} (date01 INTEGER PRIMARY KEY, open01 INTEGER)")
            # conn.commit()
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
            conn.executemany(f"INSERT INTO {self.T01_NAME} (date01, open01) VALUES (?, ?)", rows)

    def get_status(self, day_str):
        # 指定日期 得到是否為開市日 
        if day_str in self.status:
            return self.status[day_str]
        else:
            return None

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
                if self.status[d]:
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
        if not self.get_status(day_str) :                                 # 判別 day是休市日
            day_str = self.get_Dayopen(day_str, 1)                        # 取得 下一個開市日
            day = datetime.datetime.strptime(dd_str, '%Y%m%d')
        if day0 > day:                                                      # 如果 day0 >月結日  則 day 設定為下一個月
            year, month = (day.year + 1, 1) if day.month == 12 else (day.year, day.month + 1)       # 處理跨年問題
        else:
            year, month =(day.year,day.month)     
        cod0 = f"MXF{chr(month + 64)}{year%10}"                             # 商品代號 (小台指-當月)
        return cod0                                                         
# --------------------------------------------------------------------------------------------------------------------------------
def set_path0(path0='C:\\') -> tuple[str, str, str, str]:
    drive_letter = path0[:2]  
    if not os.path.exists(drive_letter):                        # 檢查磁碟機是否存在 
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
# --------------------------------------------------------------------------------------------------------------------------------
def get_time_period():
    # 檢查當前時間段 輸出時間段編號: 0=日盤前(05:01:00~08:44:59); 1=日盤(08:45:00~13:45:59); 2=夜盤前(13:46:00~14:59:59); 3=夜盤(15:00:00~05:00:59)
    now = time.localtime().tm_hour * 3600 + time.localtime().tm_min * 60 + time.localtime().tm_sec  # 當前時刻（秒）
    periods = {
        range(0,     18059): 3,                     # 夜盤
        range(18060, 31499): 0,                     # 日盤前
        range(31500, 50759): 1,                     # 日盤
        range(50820, 53999): 2,                     # 夜盤前
        range(54000, 86399): 3                      # 夜盤
    } 
    for p, i in periods.items():
        if now in p:
            return i
# --------------------------------------------------------------------------------------------------------------------------------
def fex_dayzip_download(d_str) -> bool:
    # 從期交所下載 指定日期的 交易原始數據壓縮檔
    fname=os.path.join(path_tfex, f"{d_str}Fex.zip")
    if  os.path.isfile(fname):                                                  # 檔案已存在 無須下載
        return True     
    d = datetime.datetime.strptime(d_str, '%Y%m%d').date()                      
    if (datetime.date.today() - d) >= datetime.timedelta(days=40):              # 無法下載 期交所已下架(超過30天開市日)數據檔 
        return False
    url0 = "Daily_" + d_str[0:4] + "_" + d_str[4:6] + "_" + d_str[6:] + ".zip"
    url = f"https://www.taifex.com.tw/file/taifex/Dailydownload/Dailydownload/{url0}" 
    return download_file(url, fname, retries=3, wait_time=1, min_size=1024)
# --------------------------------------------------------------------------------------------------------------------------------
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
# --------------------------------------------------------------------------------------------------------------------------------
def download_fexdata():
    # 下載期交所 40天內 每日壓縮數據檔 (30天開市日)  
    s_date = datetime.date.today() - datetime.timedelta(days=40)                # 起始日期為為40天前
    e_date = datetime.date.today() - datetime.timedelta(days=1)                 # 結束日期為昨天                                                           # 生成日期序列
    cur_date = s_date
    while cur_date <= e_date:
        d0_str=cur_date.strftime("%Y%m%d")
        if market.get_status(d0_str):                                           # 開市日
            fname=os.path.join(path_tfex, f"{d0_str}Fex.zip")
            if not os.path.isfile(fname):
                fex_dayzip_download(d0_str)
        cur_date += datetime.timedelta(days=1)
# --------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    path_yuan, path_tfex, db0, db1 = set_path0()            # 設定數據放置路徑
    buf = []                                                # 數據陣列 (暫存,每秒寫出清空)                                             
    market = StockMarket()
    app = wx.App()    
    frame = AppFrame(None, size=(1000, 200), pos=(400, 50))
    frame.Show(True)
    app.MainLoop()