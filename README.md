- windows資料夾連動
   ```shell
   New-Item -ItemType SymbolicLink -Path "$env:AppData\Autodesk\ApplicationPlugins\MyTool.bundle" -Target "C:\Users\user\路徑\autocad_plugins\MyTool.bundle"
   ```
   放到 `ApplicationPlugins` 後，開啟 AutoCAD 會自動載入小工具，並載入/開啟 `MyToolbar` 工具列。

- auto_fillet_90 :
    1. PL90F 可框選/多選後做圓角（僅接受 LWPOLYLINE / POLYLINE）
    2. PL90FA 畫完線後自動做圓角，繪製期間強制使用正交（ORTHO），按 ESC 可取消

- datalink_auto :
    1. 執行 `AUTODATALINKSHEETS`
    2. 選 Excel 來源檔（預設 `.xlsx`，仍可選 `.xls/.csv`）
    3. 會一次建立 Excel 全部分頁的 Data Link：
       會用「分頁名稱（去除前後空白後）」當 Data Link 名稱；同名時會直接覆蓋舊資料
    4. 需安裝 Microsoft Excel（用於讀取工作表名稱）
    5. Data Link 會使用相對於目前 DWG 的相對路徑
    6. 建立 Data Link 時，程式會以唯讀模式讀取 Excel，並使用「不回寫來源檔 + 保留 Excel 格式(含合併儲存格)」設定

- workstation_auto_connect :
   - 目前僅適用大量排列方是一致的平板燈或工作站使用
    1. 執行 `AUTOWSCONNECTRANGE`
    2. 依序完成 1~5 步：
       1/5 框選/多選工作站圖例（方式同 PL90F）
       2/5 框選/多選要連線區域（方式同 PL90F；只使用這一次框選）
       3/5 自動過濾掉工作站圖例以外物件
       4/5 手動選擇水平/垂直連線方向
       5/5 完成連線（依選定方向分列/分欄連線）
    3. 所有連線皆由工作站「四分點（上/右/下/左中點）」起訖
    4. 連線樣式使用 `電氣管線` 圖層 ByLayer 樣式
    5. 自動連線距離限制為 `2000` 圖面單位（2.0m，圖面以 mm 繪製時適用）
    6. 自動連線允許水平/垂直方向誤差 `5` 度內；超過角度或超過距離會略過並回報數量
       水平連線只使用工作站左/右中點；垂直連線只使用工作站上/下中點
    7. 支援圖例已 Explode 的情境：
       圖例自動判定優先序為 Block/圓/文字；若第 1 步框選內含文字（如 `E3`），第 3 步會把該文字納入篩選條件
       若精準過濾為 0，會自動用忽略圖層的寬鬆比對重試

- connected_symbol_groups :
   1. 執行 `AUTOCONNECTGROUPS`（短指令 `AUTOCG`）
   2. 依序完成 1~4 步：
      1/4 框選/多選要分組的圖示圖例（圖例判定方式同 `AUTOWSCONNECTRANGE`）
      2/4 框選/多選要分析的圖示與連線範圍
      3/4 自動過濾符合圖例的圖示，並抓取連線
      4/4 依連線關係找出不同群組並標註編號
   3. 連線會優先使用 `電氣管線` 圖層上的 `LINE / LWPOLYLINE / POLYLINE`；若範圍內沒有該圖層線段，會改用範圍內全部線段
   4. 編號會建立在 `連線群組編號` 圖層，格式為 `G1`, `G2`, `G3`...
   5. 執行時可選擇是否刪除舊的 `連線群組編號` 圖層文字，避免重複標註
   6. 另一種單點箭頭方式：執行 `AUTOCONNECTGROUPSALL`（短指令 `AUTOCGA`）
      1/5 先框選/多選適用此排序的圖例，可一次選多種圖例；圖例條件會包含顏色
      2/5 選取一個箭頭範例；直接 Enter 則使用預設箭頭
      3/5 選取盤體物件，或指定盤體位置點
      4/5 框選/多選要分析的全部圖示與連線範圍
      5/5 程式只抓符合第 1 步圖例的圖示，依連線分組，且每組只在一個代表圖示上放置箭頭
      箭頭會貼到該組離盤體最近圖示的四分點：左右中點使用水平箭頭，上下中點使用垂直箭頭；箭頭會建立在 `連線群組編號` 圖層
