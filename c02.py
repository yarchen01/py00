import re
import ctypes
import comtypes
import comtypes.client
import datetime
import dateutil.relativedelta
import tkinter as tk
from tkinter import ttk
import logging
from logging.handlers import TimedRotatingFileHandler

class YuantaQuoteAXCtrl:
    def __init__(self, parent):
        # 父层元件必须是顶层视窗，才能正常接收事件
        self.parent = parent

        # 配置日志记录器
        logging.basicConfig(filename='yuanta.log', level=logging.INFO,format='%(asctime)s %(levelname)s %(message)s')
        
        # 準備呼叫 ATL API 時所需參數
        container = ctypes.POINTER(comtypes.IUnknown)()
        control = ctypes.POINTER(comtypes.IUnknown)()
        guid = comtypes.GUID()
        sink = ctypes.POINTER(comtypes.IUnknown)()

        # 建立 ActiveX 控制元件实体
        ctypes.windll.atl.AtlAxCreateControlEx(
            'YUANTAQUOTE.YuantaQuoteCtrl.1',
            self.parent.winfo_id(),
            None,
            ctypes.byref(container),
            ctypes.byref(control),
            ctypes.byref(guid),
            sink
        )

        # 取得 ActiveX 控制元件实体
        self.ctrl = comtypes.client.GetBestInterface(control)

        # 綁定 ActiveX 控制元件事件
        self.sink = comtypes.client.GetEvents(self.ctrl, self)

        # 連線資訊
        self.Host = None
        self.Port = None
        self.Username = None
        self.Password = None

        # 记录初始化完成的信息
        logging.info("YuantaQuoteAXCtrl initialized")       

    # 設定連線資訊
    def Config(self, host, port, username, password):
        self.Host = host
        self.Port = port
        self.Username = username
        self.Password = password
    # ActiveX 控制元件事件
    def OnGetBreakResume(self,
        symbol,
        breakTime,
        resumeTime,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetDelayClose(self,
        symbol,
        delayClose,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetDelayOpen(self,
        symbol,
        delayOpen,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetFutStatus(self,
        symbol,
        functionCode,
        breakTime,
        startTime,
        reopenTime,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetLimitChange(self,
        symbol,
        functionCode,
        statusTime,
        level,
        expandType,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetMktAll(self,
        symbol,
        refPri,
        openPri,
        highPri,
        lowPri,
        upPri,
        dnPri,
        matchTime,
        matchPri,
        matchQty,
        tolMatchQty,
        bestBuyQty,
        bestBuyPri,
        bestSellQty,
        bestSellPri,
        fdbPri,
        fdbQty,
        fdsPri,
        fdsQty,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetMktData(self,
        priType,
        symbol,
        qty,
        pri,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetMktQuote(self,
        symbol,
        disClosure,
        duration,
        reqType):
        pass

    # ActiveX 控制元件事件
    # 連線行為造成的狀態改變會從這個事件通知
    def OnMktStatusChange(self,
        status,
        msg,
        reqType):
        # 取出訊息開頭可能存在的連線狀態代碼
        code = ''
        find = re.match(r'^(?P<code>\d).+$', msg)
        if find:
            code = msg[0]
            msg = msg[1:]
        # 去除訊息結尾可能存在的驚嘆號
        if msg.endswith('!'):
            msg = msg[:-1]
        # 計算文字寬度並補上結尾空白
        msg = msg + ' ' * (50 - wcwidth.wcswidth(msg))
        # 產出欄位值
        clX = f' {datetime.datetime.now():%H:%M:%S.%f} '
        cl01 = f' {reqType: >7} '
        cl02 = f' {status: >6} '
        cl03 = f' {code: >4} '
        cl04 = f' {msg} '
        
        logger.info('OnMktStatusChange\n' +
            '--------------------------------------------------------------------------------------------------\n' +
            '|                 | reqType | status | code | msg                                                |\n' +
            '--------------------------------------------------------------------------------------------------\n' +
            f'|{           clX}|{   cl01}|{  cl02}|{cl03}|{cl04                                              }|\n' +
            '--------------------------------------------------------------------------------------------------\n')

    # ActiveX 控制元件事件
    # 登入後不定期會接收到元大服務主機發送的時戳，供客戶端進行校時使用
    def OnGetTimePack(self,
        strTradeType,
        strTime,
        reqType):
        clX = f' {datetime.datetime.now():%H:%M:%S.%f} '
        cl01 = f' {reqType: >7} '
        cl02 = f' {strTime[:2]}:{strTime[2:4]}:{strTime[4:6]}.{strTime[6:]} '
        cl03 = f' {strTradeType: >12} '
        logger.info(
            f'OnGetTimePack\n' +
            f'--------------------------------------------------------------\n' +
            f'|                 | reqType |         strTime | strTradeType |\n' +
            f'--------------------------------------------------------------\n' +
            f'|{            clX}|{   cl01}|{           cl02}|{        cl03}|\n' +
            f'--------------------------------------------------------------\n'
        )

    # ActiveX 控制元件事件
    def OnGetTradeStatus(self,
        symbol,
        tradeStatus,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnRegError(self,
        symbol,
        updMode,
        errCode,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnTickRangeDataError(self,
        symbol,
        errCode,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnTickRegError(self,
        symbol,
        mode,
        errCode,
        reqType):
        pass

def main():
    root = tk.Tk()
    root.title("Yuanta.Quote")
    root.withdraw()  # 隱藏 root 視窗

    # 配置日誌記錄器
    log_formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
    log_file = f'yuanta_{datetime.date.today():%Y%m%d}.log'
    log_handler = TimedRotatingFileHandler(log_file, when='D', interval=1, backupCount=7)
    log_handler.setFormatter(log_formatter)
    log_handler.setLevel(logging.DEBUG)

    logger = logging.getLogger('yuanta_logger')
    logger.addHandler(log_handler)
    logger.setLevel(logging.DEBUG)

    # 建立封裝後的元大期貨 API 報價元件
    quote = YuantaQuoteAXCtrl(root)
    # 設定連線資訊
    quote.Config(
        host="apiquote.yuantafutures.com.tw",
        port="80",
        username="A122881975",
        password="31415926"
    )
    # 登入元大服務
    quote.Logon()

    logger.info("YuantaQuoteAXCtrl initialized")

    root.mainloop()