- windows資料夾連動
   ```shell
   New-Item -ItemType SymbolicLink -Path "$env:AppData\Autodesk\ApplicationPlugins\MyTool.bundle" -Target "C:\Users\user\路徑\autocad_plugins\MyTool.bundle"
   ```

- auto_fillet_90 :
    1. PL90F 可框選/多選後做圓角（僅接受 LWPOLYLINE / POLYLINE）
    2. PL90FA 畫完線後自動做圓角，繪製期間強制使用正交（ORTHO），按 ESC 可取消

- datalink_auto :
    1. 載入 `datalink_auto.lsp`
    2. 執行 `AUTODATALINKSHEETS`
    3. 選 Excel 來源檔（預設 `.xlsx`，仍可選 `.xls/.csv`）
    4. 會一次建立 Excel 全部分頁的 Data Link：
       會用「分頁名稱（去除前後空白後）」當 Data Link 名稱；同名時會直接覆蓋舊資料
    5. 需安裝 Microsoft Excel（用於讀取工作表名稱）
    6. Data Link 會使用相對於目前 DWG 的相對路徑
    7. 建立 Data Link 時，程式會以唯讀模式讀取 Excel，並使用「不回寫來源檔 + 保留 Excel 格式(含合併儲存格)」設定

- workstation_auto_connect :
    1. 載入 `workstation_auto_connect.lsp`
    2. 執行 `AUTOWSCONNECTRANGE`
    3. 依序完成 1~5 步：
       1/5 框選/多選工作站圖例（方式同 PL90F）
       2/5 框選/多選要連線區域（方式同 PL90F；只使用這一次框選）
       3/5 自動過濾掉工作站圖例以外物件
       4/5 手動選擇水平/垂直連線方向
       5/5 完成連線（依選定方向分列/分欄連線）
    4. 所有連線皆由工作站「四分點（上/右/下/左中點）」起訖
    5. 連線樣式使用 `電氣管線` 圖層 ByLayer 樣式
    6. 自動連線距離限制為 `2000` 圖面單位（2.0m，圖面以 mm 繪製時適用）
    7. 自動連線允許水平/垂直方向誤差 `5` 度內；超過角度或超過距離會略過並回報數量
       水平連線只使用工作站左/右中點；垂直連線只使用工作站上/下中點
    8. 支援圖例已 Explode 的情境：
       圖例自動判定優先序為 Block/圓/文字；若第 1 步框選內含文字（如 `E3`），第 3 步會把該文字納入篩選條件
       若精準過濾為 0，會自動用忽略圖層的寬鬆比對重試
