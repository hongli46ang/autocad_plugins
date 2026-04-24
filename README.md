- windows資料夾連動
   ```shell
   New-Item -ItemType SymbolicLink -Path "$env:AppData\Autodesk\ApplicationPlugins\MyTool.bundle" -Target "C:\Users\user\路徑\autocad_plugins\MyTool.bundle"
   ```

- auto_fillet_90 :
    1. PL90F 可框選/多選後做圓角（僅接受 LWPOLYLINE / POLYLINE）
    2. PL90FA 畫完線後自動做圓角，繪製期間強制使用正交（ORTHO），按 ESC 可取消

- datalink_auto :
    1. 載入 `datalink_auto.lsp`
    2. 執行 `DLAUTO` 或 `AUTODATALINK`
    3. 先選盤名（可點選文字或手動輸入）
    4. 再選 Excel 來源檔（預設 `.xlsx`，仍可選 `.xls/.csv`）
    5. 會自動依盤名對應同名分頁，並建立 Data Link（若找不到同名分頁，會請手動指定）
       若已有同名 Data Link，會直接覆蓋舊的 Data Link
    6. 需安裝 Microsoft Excel（用於讀取工作表名稱）
    7. Data Link 會使用相對於目前 DWG 的相對路徑
    8. 批量建立可用 `DLAUTOBATCH` 或 `AUTODATALINKBATCH`：
       選 Excel 後，框選多個盤名文字（TEXT/MTEXT/ATTRIB/ATTDEF，選取方式同 `PL90F`），會自動逐一建立並回報結果
    9. 若要檢查目前 DWG 內「有效 Data Link 物件」名稱，可用 `DLLINKLIST` / `AUTODATALINKLIST` / `DATALINKLIST` / `DATALINKLISTT`
    10. 若要查殘留名稱（只有 key、沒有物件），可用 `DLLINKLISTDBG` 看 `VLA_KEY`
    11. 若要一次建立 Excel 全部分頁的 Data Link，可用 `DLAUTOSHEETS` / `AUTODATALINKSHEETS` / `DLAUTOALLSHEETS`：
       會用「分頁名稱（去除前後空白後）」當 Data Link 名稱；同名時會直接覆蓋舊資料
    12. 建立 Data Link 時，程式會以唯讀模式讀取 Excel，並使用「不回寫來源檔 + 保留 Excel 格式(含合併儲存格)」設定
    13. 若 `DLTABLELINKCHECK` 顯示大量 `NO_OBJECT`，可用 `DLDATALINKREPAIR` / `AUTODATALINKREPAIR` / `DATALINKREPAIR`：
       可連續選多個 Excel，會自動以同名分頁重建 Data Link（先清除殘留 key，再以新資料為主）
       若目前圖面「完全沒有 Data Link」，同指令會自動切換成全量重建模式（由所選 Excel 全部分頁建立）
    14. 若你要忽略目前狀態、直接覆蓋重建全部 Data Link，可用
       `DLDATALINKREBUILDALL` / `AUTODATALINKREBUILDALL` / `DATALINKREBUILDALL`

- datalink_table_auto :
    1. 載入 `datalink_auto.lsp` 與 `datalink_table_auto.lsp`
    2. 執行 `DLTABLEUI` / `DLTABLEAUTO` / `AUTODATALINKTABLE` / `DLLINKTABLE`
       建議流程：先執行 `DLAUTOSHEETS` 選 Excel，一次產生所有分頁 Data Link；
       再執行 `DLTABLEUI`，程式會開啟 AutoCAD `TABLE` 視窗；
       開啟前會先將 Data Link 改為「只更新資料、跳過 Excel 格式」；
       請在視窗中手動選擇 `From a data link` / `從資料連結`、選 Data Link 後指定插入點。
       成功建立的表格會保留 Data Link 綁定，後續可用 AutoCAD 內建 `DATALINKUPDATE` 更新；
       `DATALINKUPDATE` 開始前會自動把 Data Link 改為「只更新資料、跳過 Excel 格式」；
       結束後會連續延後重套目前頁籤表格格式，避免欄寬/文字/邊框被 Excel 格式覆蓋。
       格式化採用 Data Link 安全模式，不逐格改寫連結儲存格，避免跳出「覆蓋連結資料」提示。
    3. `DLTABLEDIALOG` 走同一個 `TABLE` 視窗流程；建完後同樣會自動格式化
    4. `TBFIX` 可手動選一個既有表格並套用同樣格式；`TBFIXALL` 可重套目前頁籤全部表格
       `DLDATAONLY` 可手動把目前 DWG 的 Data Link 改為只更新資料、跳過 Excel 格式
    5. 建表後會自動套用格式：
       儲存格寬度 `2025`、儲存格高度 `486`、文字高度 `200`、
       文字樣式 `Standard`、對齊方式置中

- datalink_batch_table :
    1. 載入 `datalink_auto.lsp` 與 `datalink_batch_table.lsp`
    2. 建議先執行 `DLAUTOSHEETS` / `AUTODATALINKSHEETS`，從 Excel 全部分頁建立 Data Link
    3. 執行 `DLLIST` / `DLBATCHLIST` 可確認目前圖面內可用的 Data Link
    4. 執行 `DLBATCHAUTO` / `DLBATCHTABLE` / `DLBATCHTABLES`：
       指定第一張表格插入點與垂直間距後，會依 Data Link 清單逐張建立表格並套用格式
    5. `DLBATCHFIX` 可手動選一個既有表格並套用 batch table 的格式；一般格式化也可用 `TBFIX` / `TBFIXALL`
    6. 若 AutoCAD 版本未透過 ActiveX 暴露 `Table.SetDataLink`，請改用 `DLTABLEUI` 互動建立後再格式化

- workstation_auto_connect :
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
