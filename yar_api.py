# 名稱:         yar_api.py 
# 功能:         元大行情API接口   YuantaQuoteAXCtrl  (YuantaQuote_v2.1.2.9.ocx) (2022/11/29)
# 限制:         Python(3.9.13-Win32)  wxPython(4.1.1)  comtypes(1.1.11)      
# 版本:         20230523-b15
import logging
logging.basicConfig(
    level=logging.WARNING,
    filename='yar.log',
    filemode='w',
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    style='%',
    force=True
)
# 建立 logger 物件
logger = logging.getLogger(__name__)

import ctypes 
import comtypes
import comtypes.client
import time
from datetime import date
class YuantaQuoteAXCtrl:
    def __init__(self, parent):
        self.parent = parent
        self.frame_wrapper = parent
        Iwindow  = ctypes.POINTER(comtypes.IUnknown)()                           
        Icontrol = ctypes.POINTER(comtypes.IUnknown)()
        Ievent   = ctypes.POINTER(comtypes.IUnknown)()
        ctypes.windll.atl.AtlAxCreateControlEx(
                "YUANTAQUOTE.YuantaQuoteCtrl.1",  
                self.parent.Handle,  
                None,
                ctypes.byref(Iwindow),
                ctypes.byref(Icontrol),
                comtypes.byref(comtypes.GUID()),
                Ievent)         
        self.ctrl = comtypes.client.GetBestInterface(Icontrol)              # 取得 ActiveX 控制元件實體
        self.sink = comtypes.client.GetEvents(self.ctrl, self)              # 綁定 ActiveX 控制元件事件 
    # ------------------------------------------------------------------------API 提供呼叫
    def api_Logon(self,acc, pwd, port, reqType):
        # api 帳號密碼登入      reqType:(日盤 1, 夜盤 2)    port:(日盤 80/443, 夜盤 82/442)
        ret=self.ctrl.SetMktLogon(acc, pwd, 'apiquote.yuantafutures.com.tw', port, reqType, 0)
        time.sleep(3)
        logging.info(f"api_Logon:{acc}, {pwd}, {port}, {reqType}  ({ret})")        

    def api_Register(self, symbol, reqType):
        # Qapi 商品代號 註冊    Symbol:商品代號;   reqType:(日盤 True, 夜盤 False)
        ret=self.ctrl.AddMktReg (symbol, '4' , reqType, 0)   # 註冊方式UpdateMode 1:Snapshot當時最新資料,2:Update註冊後的所有更新,4:(1+2)
        time.sleep(1)
        logging.info(f"api_register:{symbol}, {reqType}  ({ret})")
    
    def api_UnRegister(self, symbol):
        # Qapi 商品代號 註消
        ret=self.ctrl.DelMktReg (symbol, 0)
        time.sleep(1)
        #logging.info(f"api_UnRegister:{symbol}  ({ret}) ")
    # ------------------------------------------------------------------------API 主動回報
    def OnMktStatusChange(self,this,Status,Msg,ReqType):
        #狀態回報(連線/登入)
        #logging.info(f"OnMktStatusChange:{this},{Status},{Msg},{ReqType}")
        self.frame_wrapper.OnMktStatusChange(Status)                                        

    def OnGetMktData(self,this,PriType,symbol,Qty,Pri,ReqType):                             
        #註冊回報(商品代號) 
        #logging.info(F"OnGetMktData:{this},{PriType},{symbol},{Qty},{Pri},{ReqType}")
        self.frame_wrapper.OnGetMktData(PriType,symbol)

    def OnGetMktAll(self, this, symbol, RefPri, OpenPri, HighPri, LowPri, UpPri, DnPri, MatchTime, MatchPri, MatchQty, TolMatchQty,BestBuyQty, BestBuyPri, BestSellQty,BestSellPri, FDBPri, FDBQty, FDSPri, FDSQty, ReqType):
        #即時數據回報(API)
        #logging.info(f"OnGetMktAll:{this},{symbol},{RefPri},{OpenPri},{HighPri},{LowPri},{UpPri},{DnPri},{MatchTime},{MatchPri},{MatchQty},{TolMatchQty},{BestBuyQty},{BestBuyPri},{BestSellQty},{BestSellPri},{FDBPri},{FDBQty},{FDSPri},{FDSQty},{ReqType}")    
        self.frame_wrapper.OnGetMktAll(symbol,RefPri,OpenPri,HighPri,LowPri,UpPri,DnPri,MatchTime,MatchPri,MatchQty,TolMatchQty,BestBuyQty,BestBuyPri,BestSellQty,BestSellPri,FDBPri,FDBQty,FDSPri,FDSQty,ReqType)

    def OnGetTimePack(self,this,strTradeType,strTime,ReqType):
        #時間校正回報
        #logging.info(f"OnGetTimePack:{this},{strTradeType},{strTime},{ReqType}")
        self.frame_wrapper.OnGetTimePack(strTime)
        
    def OnGetFutStatus(self, this,symbol,FunctionCode,BreakTime,StartTime,ReopenTime,ReqType):
        #logging.info(f"OnGetFutStatus:{this},{symbol},{FunctionCode},{BreakTime},{StartTime},{ReopenTime},{ReqType}")
        return                                          
    
    def OnGetMktQuote(self,this,symbol,DisClosure,Duration,ReqType):
        #logging.info(f"OnGetMktQuote:{this},{symbol},{DisClosure},{Duration},{ReqType}")
        return
    def OnGetDelayClose(self,this,symbol,DelayClose,ReqType):
        #logging.info(f"OnGetDelayClose:{this},{symbol},{DelayClose},{ReqType}")
        return
    def OnGetBreakResume(self,this,symbol,BreakTime,ResumeTime,ReqType):
        #logging.info(f"OnGetBreakResume:{this},{symbol},{BreakTime},{ResumeTime},{ReqType}")
        return
    def OnGetTradeStatus(self,this,symbol,TradeStatus,ReqType):
        #logging.info(f"OnGetTradeStatus:{this},{symbol},{TradeStatus},{ReqType}")
        return
    def OnTickRegError(self,this,strSymbol,lMode,lErrCode,ReqType):
        #logging.info(f"OnTickRegError:{this},{strSymbol},{lMode},{lErrCode},{ReqType}")
        return
    def OnGetTickData(self,this,strSymbol,strTickSn,strMatchTime,strBuyPri,strSellPri,strMatchPri,strMatchQty,strTolMatQty,strMatchAmt,strTolMatAmt,ReqType):
        #logging.info(f"OnGetTickData:{this},{strSymbol},{strTickSn},{strMatchTime},{strBuyPri},{strSellPri},{strMatchPri},{strMatchQty},{strTolMatQty},{strMatchAmt},{strTolMatAmt},{ReqType}")
        return
    def OnTickRangeDataError(self,this,strSymbol,lErrCode,ReqType):
        #logging.info(f"OnTickRangeDataError:{this},{strSymbol},{lErrCode},{ReqType}")
        return
    def OnGetTickRangeData(self,this,strSymbol,strStartTime,strEndTime,strTolMatQty,strTolMatAmt,ReqType):
        #logging.info(f"OnGetTickRangeData:{this},{strSymbol},{strStartTime},{strEndTime},{strTolMatQty},{strTolMatAmt},{ReqType}")
        return
    def OnGetDelayOpen(self,this,symbol,DelayOpen,ReqType):
        #logging.info(f"OnGetDelayOpen:{this},{symbol},{DelayOpen},{ReqType}")
        return
    def OnGetLimitChange(self,this,symbol,FunctionCode,StatusTime,Level,ExpandType,ReqType):
        #logging.info(f"OnGetLimitChange:{this},{symbol},{FunctionCode},{StatusTime},{Level},{ExpandType},{ReqType}")
        return
    def OnRegError(self,this,symbol,updmode,ErrCode,ReqType):
        #logging.info(f"OnRegError:{this},{symbol},{updmode},{ErrCode},{ReqType}")
        return