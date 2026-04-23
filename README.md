auto_fillet_90 :
    1. PL90F 可框選/多選後做圓角（僅接受 LWPOLYLINE / POLYLINE）
    2. PL90FA 畫完線後自動做圓角，繪製期間強制使用正交（ORTHO），按 ESC 可取消

datalink_auto :
    1. 載入 `datalink_auto.lsp`
    2. 執行 `DLAUTO` 或 `AUTODATALINK`
    3. 先選盤名（可點選文字或手動輸入）
    4. 再選 Excel 來源檔（預設 `.xlsx`，仍可選 `.xls/.csv`）
    5. 會自動依盤名對應同名分頁，並建立 Data Link（若找不到同名分頁，會請手動指定）
       若已有同名 Data Link，會直接覆蓋舊的 Data Link
    6. 需安裝 Microsoft Excel（用於讀取工作表名稱）
    7. Data Link 會使用相對於目前 DWG 的相對路徑
    8. 批量建立可用 `DLAUTOBATCH` 或 `AUTODATALINKBATCH`：
       選 Excel 後，框選多個盤名文字（TEXT/MTEXT/ATTRIB/ATTDEF），會自動逐一建立並回報結果
    9. 若要檢查目前 DWG 內 Data Link 名稱，可用 `DLLINKLIST` / `AUTODATALINKLIST` / `DATALINKLIST` / `DATALINKLISTT`
    10. 若清單異常為空，可用 `DLLINKLISTDBG` 顯示各種掃描來源的偵錯結果
    11. 若要一次建立 Excel 全部分頁的 Data Link，可用 `DLAUTOSHEETS` / `AUTODATALINKSHEETS` / `DLAUTOALLSHEETS`：
       會用「分頁名稱（去除前後空白後）」當 Data Link 名稱；同名時會直接覆蓋舊資料
    12. 建立 Data Link 時，程式會以唯讀模式讀取 Excel，並使用「不回寫來源檔 + 保留 Excel 格式(含合併儲存格)」設定

datalink_table_auto :
    1. 載入 `datalink_auto.lsp` 與 `datalink_table_auto.lsp`
    2. 執行 `DLTABLEAUTO` / `AUTODATALINKTABLE` / `DLLINKTABLE`
    3. 僅支援「指定 Data Link 建表」：
       預設先「選取圖面文字（TEXT/MTEXT/ATTRIB/ATTDEF）」指定 Data Link 名稱；
       按 Enter/Space 不選取時，會改為手動輸入 Data Link 名稱/序號
    4. 建表後會自動套用格式：
       儲存格寬度 `2025`、儲存格高度 `486`、文字高度 `200`、
       文字樣式 `微軟正黑體`、表格顏色 `ByLayer`（邊框跟隨圖層顏色）
       若系統找不到 `msjh.ttc`，程式會跳過強制字型並提示警告
    5. 為避免誤拆合併列，建表後「預設不自動刪空白列」
       若確定你的表格沒有合併列，才可手動設 `dlt:*remove-empty-rows-after-create*` 為 `T` 啟用刪除
    6. 批次重綁既有表格可用 `DLTABLEREBIND` / `DLTABLEREBINDBATCH` / `AUTODATALINKTABLEREBIND` / `DLLINKTABLEREBIND`
       可先框選既有表格；若直接 Enter，會抓目前頁籤全部 ACAD_TABLE
       程式會優先讀取舊表原本綁定的 Data Link 名稱來對應同名目標；讀不到時才回退用表格內容推斷
       「鄰近文字比對」僅接受目前 DWG 已存在的 Data Link 名稱，並採用邊界安全的包含比對（例如 `盤名:R11A1` 可對應 `R11A1`）
       若不希望使用鄰近文字回退，可先設 `(setq dlt:*enable-near-text-fallback* nil)`
       重綁方式為「新建同名 Data Link 表格 + 刪除舊表格」
       完成後會列出成功/失敗/略過清單
    7. 半自動重綁可用 `DLTABLEREBINDMANUAL` / `DLTABLEREBINDPICK` / `AUTODATALINKTABLEREBINDPICK` / `DLLINKTABLEREBINDPICK`
       逐張操作：先選舊表，再選盤名文字（或手動輸入 Data Link 名稱/序號）
       適合舊表已是靜態表、無法自動辨識對應 Data Link 的情境

workstation_auto_connect :
    1. 載入 `workstation_auto_connect.lsp`
    2. 執行 `WSAUTOCONNECT` / `AUTOWSCONNECT` / `WORKSTATIONAUTOCONNECT`
    3. 依序完成 1~5 步：
       1/5 框選/多選工作站圖例（方式同 PL90F）
       2/5 框選/多選要連線區域（方式同 PL90F）
       3/5 自動過濾掉工作站圖例以外物件
       4/5 設定連線範例（先框選/多選兩個工作站決定方向，再於此步驟切到「電氣圖層」用 PLINE 畫範例線）
       5/5 完成連線（依範例方向分列/分欄連線）
    4. 所有連線皆由工作站「四分點（上/右/下/左中點）」起訖
    5. 連線樣式會套用第 4 步新畫 PLINE 的圖層、線型、顏色、線寬
    6. 自動連線僅建立水平或垂直線段；若兩點只能形成斜線會自動略過
    7. 支援圖例已 Explode 的情境：
       圖例自動判定優先序為 Block/圓/文字；若第 1 步框選內含文字（如 `E3`），第 3 步會把該文字納入篩選條件
       若精準過濾為 0，會自動用忽略圖層的寬鬆比對重試
