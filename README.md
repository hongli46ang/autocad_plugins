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
       會用「分頁名稱」當 Data Link 名稱，且同名 Data Link 會自動覆蓋
    12. 建立 Data Link 時，程式會以唯讀模式讀取 Excel，並使用「不回寫來源檔 + 保留 Excel 格式(含合併儲存格)」設定

datalink_table_auto :
    1. 載入 `datalink_auto.lsp` 與 `datalink_table_auto.lsp`
    2. 執行 `DLTABLEAUTO` / `AUTODATALINKTABLE` / `DLLINKTABLE`
    3. 僅支援「指定 Data Link 建表」：
       輸入 Data Link 名稱或序號，建立單一表格
    4. 為避免誤拆合併列，建表後「預設不自動刪空白列」
       若確定你的表格沒有合併列，才可手動設 `dlt:*remove-empty-rows-after-create*` 為 `T` 啟用刪除
