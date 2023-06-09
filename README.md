20230524   b15.py release
  1. 增加 api_datapath.txt 儲存 數據路徑
  2. 分割 b14.py  成為   b15.py,  yar_api.py,  yar_market.py
  3. 元大行情API (YuantaQuote_v2.1.2.9.ocx) envent.log 大量紀錄每筆數據 佔用處理時間及儲存空間
  4. 花時間處理 強制遮蔽 元大log的輸出

20230609 yardf_002.py  release
  1. yardf_data_002.py  使用 pandas DataFrame 處理數據, 用 df.iat[] 整數索引 提高數據更新的速度
  2. yardf_002.py       使用 mplfinance  繪圖 繪製k線圖/BBand/移動平均線  並啓用wx.Frame 設置畫面 處理輸入和顯示的資料
  3. config.sys         存放數據目錄全路徑 以及登入賬號和密碼  密碼有經過特殊加密處理
  4. yar_api.py,  yar_market.py 沿用上次 並無大改 
  5. Logs 強制遮蔽 元大API log 霸道的輸出! 
