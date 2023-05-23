# 程式名稱     小台指-當月 即時數據 錄製 
# 數據來源     元大期貨-行情API  
# 套件限制:    Python(3.9.13-Win32)  wxPython(4.1.1)  comtypes(1.1.11)     
# 版本:       (20230519-b13) (20230523-b15)
import wx
import wx.grid
import datetime
import time
import psutil
import os
import struct
import yar_api                      # 元大行情API 接口 (YuantaQuote_v2.1.2.9.ocx) (2022/11/29)
import yar_market                   # 期貨相關公用函式, 資料庫操作函式
#---------------------------------------------------------------------------------------------------------
class AppFrame(wx.Frame):
    def __init__(self, parent, size, pos, *args, **kw):
        # Frame起始設定 (設置 程式畫面/元大Qapi連結/輸入與顯示處理) 
        style = wx.DEFAULT_FRAME_STYLE & ~(wx.CLOSE_BOX)                         # 移除關閉按鈕
        super(AppFrame, self).__init__(parent, *args, **kw, style=style)                                        
        self.parent = parent                                                     # 父層元件必須是頂層視窗，才能正常接收事件
        ver0 = '(Ver:20230519-b13)' 
        self.Symbol0= ''                                                         # 商品代號(小台指-當月) 
        self.t00    = ''                                                         # 時標紀錄器 
        self.t1     = ''                                                         # 即時時標紀錄器
        self.qty    = 0                                                          # 總量紀錄器 
        self.err0   = 0                                                          # 錯誤紀錄器
        self.miss   = 0                                                          # 漏失紀錄器
        self.fpath  = ''                                                         # 每日數據檔全路徑 
        self.f_download = False                                                  # 旗標: 期交所前一日數據已下載
        Title0 = f"小台指-當月即時數據錄製 {ver0}"
        self.SetTitle(Title0)                                                          
        self.GUIinit(size, pos)
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
        self.msg1 = wx.TextCtrl(pnl, pos=(x0+55, y0), size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)
        wx.StaticText(pnl, pos=(x1,  y0), label='路徑:')
        self.path0 = wx.TextCtrl(pnl, pos=(x1+35, y0), size=(w3, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)
        self.path0.SetValue(path_root)      
        wx.StaticText(pnl, pos=(x0,  y1), label='系統訊息:')   
        self.msg2 = wx.TextCtrl(pnl, pos=(x0+55, y1), size=(w1, h0),style=wx.TE_READONLY | wx.BORDER_SUNKEN)     
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
        # ----------------------------------- 刪除目錄,建立同名檔案,並設唯讀(使其無法寫入任何log)
        dir0="./Logs"
        if os.path.isdir(dir0):
            os.rmdir(dir0)
        if not os.path.exists(dir0):        
            open(dir0, "w").close()
        os.chmod(dir0, 0o444)
        # --------------------------------                
        self.YTQ = yar_api.YuantaQuoteAXCtrl(self)
    #---------------------------------------------------------------------------------------------------------
    def logon(self,event):
        global market
        # Qapi 帳號密碼登入   
        acc0  = self.acc.GetValue()
        pwd0  = self.pwd.GetValue()
        port , reqType  = ('80', 1) if (yar_market.get_time_period() in [0, 1]) else ('82',2)                
        self.msg1.SetValue(f" 與元大伺服器網路 連線中... ")
        self.YTQ.api_Logon(acc0, pwd0, port, reqType)                            

    def RegisterQuoteSymbol(self, symbol):
        # Qapi 商品代號 註冊
        reqType = True if (yar_market.get_time_period() in [0, 1]) else False
        self.YTQ.api_Register(symbol , reqType)
        time.sleep(1.5)             
                         
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
                                                               
    def SetLogonDisable(self):
        # 失效(帳號/密碼/登入)
        self.path0.Disable()                                                      
        self.acc.Disable()
        self.pwd.Disable()
        self.btn.Disable()
   
    def SetRegister_Symbol(self):
        # 取得並註冊 (小台指-當月) 商品代號 
        if  self.Symbol0 != '':     
            self.YTQ.api_UnRegister(self.Symbol0)                                    # 商品代號註銷
            time.sleep(1.5)         
        day_str = datetime.date.today().strftime("%Y%m%d")                          # 當前日期
        if not market.get_status(day_str):                                          # 休市日 
            day_str = market.get_Dayopen(day_str,1)                                 # 調整 視為下一開市日
        else:    
            if datetime.datetime.now().time() > datetime.time(13, 46, 0):           # 夜盤 
                day_str = market.get_Dayopen(day_str,1)                             # 調整 視為下一開市日之數據                                     
        self.Symbol0 = market.get_code(day_str)                                     # 商品代號(小台指-當月) (月結日下午 商品代號會不同)
        self.fpath = os.path.join(path_yuan, f"{day_str}-{self.Symbol0}.bin0")      # 數據的全路徑
        self.grid.SetCellValue(0, 0, self.Symbol0)          
        self.RegisterQuoteSymbol(self.Symbol0)
         
    def update_time(self, event):
        # 顯示/更新 即時時間
        now = datetime.datetime.now()
        self.chktime_Process()                                                      # 檢查時間 並做相應特殊處理
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
        # [8:30:00~8:30:30] and [14:50:00~14:50:30] 執行 檢查數據歸屬日期(影響數據檔案名稱和商品代號),重命名數據檔; 重註冊商品代號
        tt = datetime.datetime.now().time()
        if (datetime.time(8,30,0) <= tt <= datetime.time(8,30,30)) or (datetime.time(14,50,0) <= tt <= datetime.time(14,50,30)):
            self.SetRegister_Symbol()
            self.msg2.SetValue(self.Symbol0 + ' 商品代號重新註冊,請稍等...')
            time.sleep(31)
            self.msg2.SetValue(self.Symbol0+ ' 商品代號 已完成註冊!!')    
        # [6:00:00~6:00:30] 執行 切換相關旗標
        if  (datetime.time(6,00,0) <= tt <= datetime.time(6,00,30)):
            if self.f_download:
                self.f_download = False                                    
        # [6:30:00~23:59:59] 執行 下載期交所前一日數據
        if not self.f_download:                                            
            if datetime.time(6,30,0) <= tt:                                 # 6:30:00 之後         
                market.download_fexdata()                                   # 下載期交所前一日數據(自前30開市日至前一開市日)
                self.f_download = True
        
    def on_timeclose(self, event):
        self.timer.Stop()
                   
    def on_exit(self, event):
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
            self.SetLogonDisable()
            self.SetRegister_Symbol()
            time.sleep(2)

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
        if yar_market.get_time_period(MatchTime) not in (1, 3, 4) or self.Symbol0 != symbol:
            return                                                                      # 不在開盤時段 或 商品代號不符合        
        i_tqty = int(TolMatchQty or 0)                                                  # 總量
        i_qty = int(MatchQty or 0)                                                      # 單量            
        if (i_tqty <=0) or (i_tqty <= self.qty) or (i_qty <=0):                         # 總量<=0 or 總量<=前總量 or 單量<=0
            return       
        if i_tqty != (self.qty+i_qty):                                                  # (總量 != 前總量 + 單量)
            self.miss += 1                                                              # 漏失數據 
        self.qty = i_tqty                                                               # 回存總量
        # ----------------------------------------------------------------------------               
        global buf                                                                      # 暫存
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
        buf.append(dd)                                                              # 存入 暫存數據陣列                   
        t=("000000000000"+MatchTime)[-12:][:9]                                      # 時標
        if  (self.t00 != t[0:6]):                                                   # 每秒 輸出即時數據到檔案
            if len(buf) > 0:
                with open(self.fpath, 'ba') as file:
                    for i in range(len(buf)):
                        file.write(buf[i])
                    buf=[]        
            self.t00 = t[0:6]
            self.UpdateSymbol(RefPri,OpenPri,HighPri,LowPri,t,MatchPri,MatchQty,TolMatchQty)

    def OnGetFutStatus(self,symbol):
        self.msg2.SetValue(f" 開始傳送商品代號 {symbol} 的即時數據") 
# --------------------------------------------------------------------------------------------------------------------------------
def set_path0() -> tuple[str, str, str, str, str]:
    # 從檔案取得路徑資料 設定數據存放全路徑名稱
    path0='C:\\'
    fname = 'api_datapath.txt'
    if  os.path.isfile(fname):
        with open(fname, "r") as file:
            path0 = os.path.abspath(file.read())
            path0=path0.split('\n')[0]
            if not os.path.exists(path0[:2]):                   # 檢查磁碟機是否存在 
                path0='C:\\'          
    path_root = os.path.join(path0, 'data00')                   # 數據主目錄  
    path_yuan = os.path.join(path_root, 'yuanta')               # 元大數據下載路徑
    path_tfex = os.path.join(path_root, 'taifex')               # 期交所數據下載路徑
    db0 = os.path.join(path_root, "fex00.db")                   # 數據庫 (存放 期交所-逐筆數據)
    db1 = os.path.join(path_root, "fex01.db")                   # 數據庫 (存放 期交所-每秒/每分/日夜/彙整數據)
    os.makedirs(path_root, exist_ok=True)
    os.makedirs(path_yuan, exist_ok=True)
    os.makedirs(path_tfex, exist_ok=True)
    return path_root, path_yuan, path_tfex, db0, db1
# --------------------------------------------------------------------------------------------
if __name__ == "__main__":
    path_root, path_yuan, path_tfex, db0, db1 = set_path0()             # 設定數據放置路徑
    buf = []                                                            # 數據陣列 (暫存,每秒寫出清空)  
    market = yar_market.StockMarket(path_db=db1, path_tfex=path_tfex)
    app = wx.App()    
    frame = AppFrame(None, size=(1000, 200), pos=(400, 50))
    frame.Show(True)
    # YTquote = yar_api.YuantaQuoteAXCtrl(frame)
    app.MainLoop()

    
    