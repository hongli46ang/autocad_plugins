(vl-load-com)

(setq dlt:*last-removed-empty-rows* 0)
; 預設關閉：避免刪空白列時誤拆到合併列
(setq dlt:*remove-empty-rows-after-create* nil)
(setq dlt:*last-create-failure-reason* nil)
(setq dlt:*has-command-s* nil)
(setq dlt:*cell-width* 2025.0)
(setq dlt:*cell-height* 486.0)
(setq dlt:*text-height* 200.0)
(setq dlt:*text-style-name* "DLT_MSJH")
(setq dlt:*text-font-file* "msjh.ttc")
(setq dlt:*text-font-candidates* '("msjh.ttc" "MSJH.TTC" "msjh.ttf" "MSJH.TTF"))
(setq dlt:*msjh-font-resolved* nil)
(setq dlt:*msjh-warning-shown* nil)

(defun dlt:trim (s)
  (if s
    (vl-string-trim " \t\r\n" s)
    ""
  )
)

(defun dlt:list-has-ci (item lst / key)
  (setq key (strcase (dlt:trim item)))
  (if (vl-some '(lambda (x) (= (strcase (dlt:trim x)) key)) lst)
    T
    nil
  )
)

(defun dlt:find-ci (item lst / key)
  (setq key (strcase (dlt:trim item)))
  (vl-some
    '(lambda (x)
       (if (= (strcase (dlt:trim x)) key)
         x
       )
     )
    lst
  )
)

(defun dlt:digits-only-p (s / i ok ch)
  (setq s  (dlt:trim s)
        i  1
        ok (/= s ""))
  (while (and ok (<= i (strlen s)))
    (setq ch (substr s i 1))
    (if (null (wcmatch ch "#"))
      (setq ok nil)
    )
    (setq i (1+ i))
  )
  ok
)

(defun dlt:resolve-link-name (input names / txt idx)
  (setq txt (dlt:trim input))
  (cond
    ((= txt "") nil)
    ((dlt:find-ci txt names))
    ((dlt:digits-only-p txt)
     (setq idx (atoi txt))
     (if (and (> idx 0) (<= idx (length names)))
       (nth (1- idx) names)
       nil
     )
    )
    (T nil)
  )
)

(defun dlt:set-dxf (ed code val / old)
  (if (setq old (assoc code ed))
    (subst (cons code val) old ed)
    (append ed (list (cons code val)))
  )
)

(defun dlt:print-link-names (names / i)
  (setq i 1)
  (prompt "\n--- 目前 Data Link 清單 ---")
  (foreach n names
    (prompt (strcat "\n  " (itoa i) ". " n))
    (setq i (1+ i))
  )
)

(defun dlt:init-command-s-flag ()
  ; atoms-family 1 為已定義函數符號，避免直接呼叫不存在函數
  (if (vl-position "COMMAND-S" (atoms-family 1))
    (setq dlt:*has-command-s* T)
    (setq dlt:*has-command-s* nil)
  )
)

(defun dlt:run-safe-command (args / ret)
  ; 優先用 command-s，避免互動命令卡在等待輸入
  (if dlt:*has-command-s*
    (vl-catch-all-apply 'command-s args)
    'dlt_no_command_s
  )
)

(defun dlt:flush-command (/ tries)
  (setq tries 0)
  (while (and (> (getvar "CMDACTIVE") 0) (< tries 8))
    (if dlt:*has-command-s*
      (vl-catch-all-apply 'command-s (list (strcat (chr 3) (chr 3))))
      (setq tries 8)
    )
    (setq tries (1+ tries))
  )
  (if (> (getvar "CMDACTIVE") 0)
    (setq dlt:*last-create-failure-reason* "cancel_incomplete")
  )
)

(defun dlt:get-new-table-ename-after (entBefore / en ed lastTable)
  (setq lastTable nil
        en (if entBefore (entnext entBefore) (entnext)))
  (while en
    (setq ed (entget en))
    (if (and ed (= (cdr (assoc 0 ed)) "ACAD_TABLE"))
      (setq lastTable en)
    )
    (setq en (entnext en))
  )
  lastTable
)

(defun dlt:get-cell-text-safe (tableObj row col / v)
  (setq v (vl-catch-all-apply 'vlax-invoke-method (list tableObj 'GetText row col)))
  (if (vl-catch-all-error-p v)
    ""
    (dlt:trim v)
  )
)

(defun dlt:get-text-from-entity (ename / obj oname txt)
  (setq txt "")
  (if ename
    (progn
      (setq obj (vlax-ename->vla-object ename))
      (if obj
        (progn
          (setq oname (vla-get-ObjectName obj))
          (if (wcmatch oname "AcDbText,AcDbMText,AcDbAttributeDefinition,AcDbAttribute")
            (setq txt (vla-get-TextString obj))
          )
        )
      )
    )
  )
  (dlt:trim txt)
)

(defun dlt:get-input-by-pick (/ ent ename txt)
  (setq txt "")
  (if (setq ent (entsel "\n選取作為 Data Link 名稱的文字物件 (Enter/Space改手動輸入): "))
    (progn
      (setq ename (car ent)
            txt   (dlt:get-text-from-entity ename))
      (if (= txt "")
        (prompt "\n所選物件不是可用文字（TEXT/MTEXT/ATTRIB/ATTDEF）或內容為空。")
      )
    )
  )
  txt
)

(defun dlt:cell-merged-p (tableObj row col / v)
  (setq v (vl-catch-all-apply 'vlax-invoke-method (list tableObj 'IsMergedCell row col)))
  (if (vl-catch-all-error-p v)
    nil
    (if (or (= v :vlax-true) (= v T) (= v 1))
      T
      nil
    )
  )
)

(defun dlt:row-empty-p (tableObj row / cols col hasText hasMerged)
  (setq cols (vlax-get-property tableObj 'Columns)
        col  0
        hasText nil
        hasMerged nil)
  (while (and (< col cols) (not hasText) (not hasMerged))
    (if (dlt:cell-merged-p tableObj row col)
      (setq hasMerged T)
      (if (/= (dlt:get-cell-text-safe tableObj row col) "")
        (setq hasText T)
      )
    )
    (setq col (1+ col))
  )
  (if hasMerged
    nil
    (not hasText)
  )
)

(defun dlt:remove-empty-rows-from-table (tableEname / tableObj row removed canDelete ret)
  (setq removed 0
        tableObj (if tableEname (vlax-ename->vla-object tableEname)))
  (if (and tableObj (= (vla-get-ObjectName tableObj) "AcDbTable"))
    (progn
      ; 自後往前刪，避免索引位移；保留第0列，避免刪到標題列
      (setq row (1- (vlax-get-property tableObj 'Rows)))
      (while (>= row 1)
        (setq canDelete (dlt:row-empty-p tableObj row))
        (if canDelete
          (progn
            (setq ret (vl-catch-all-apply 'vlax-invoke-method (list tableObj 'DeleteRows row 1)))
            (if (not (vl-catch-all-error-p ret))
              (setq removed (1+ removed))
            )
          )
        )
        (setq row (1- row))
      )
    )
  )
  (if tableObj (vlax-release-object tableObj))
  removed
)

(defun dlt:resolve-msjh-font-file (/ found)
  (if (and dlt:*msjh-font-resolved* (/= dlt:*msjh-font-resolved* ""))
    dlt:*msjh-font-resolved*
    (progn
      (setq found
            (vl-some
              '(lambda (f)
                 (if (or (findfile f)
                         (findfile (strcat "C:\\Windows\\Fonts\\" f)))
                   f
                 )
               )
              dlt:*text-font-candidates*))
      (if found
        (progn
          (setq dlt:*msjh-font-resolved* found)
          found
        )
        nil
      )
    )
  )
)

(defun dlt:ensure-msjh-text-style (/ stEn stEd fontFile)
  (setq fontFile (dlt:resolve-msjh-font-file))
  (if (null fontFile)
    nil
    (progn
      (setq stEn (tblobjname "STYLE" dlt:*text-style-name*))
      (if (null stEn)
        (progn
          (entmake
            (list
              '(0 . "STYLE")
              (cons 2 dlt:*text-style-name*)
              '(70 . 0)
              '(40 . 0.0)
              '(41 . 1.0)
              '(50 . 0.0)
              '(71 . 0)
              '(42 . 2.5)
              (cons 3 fontFile)
              '(4 . "")
            )
          )
          (setq stEn (tblobjname "STYLE" dlt:*text-style-name*))
        )
      )
      (if stEn
        (progn
          (setq stEd (entget stEn))
          (setq stEd (dlt:set-dxf stEd 3 fontFile))
          (setq stEd (dlt:set-dxf stEd 4 ""))
          (setq stEd (dlt:set-dxf stEd 40 0.0))
          (setq stEd (dlt:set-dxf stEd 41 1.0))
          (setq stEd (dlt:set-dxf stEd 50 0.0))
          (setq stEd (dlt:set-dxf stEd 71 0))
          (entmod stEd)
          (entupd stEn)
          T
        )
        nil
      )
    )
  )
)

(defun dlt:set-entity-color-bylayer (ename / ed)
  (if ename
    (progn
      (setq ed (entget ename))
      (if ed
        (progn
          (setq ed
                (vl-remove-if
                  '(lambda (x) (or (= (car x) 420) (= (car x) 430) (= (car x) 440)))
                  ed))
          (setq ed (dlt:set-dxf ed 62 256))
          (entmod ed)
          (entupd ename)
        )
      )
    )
  )
  T
)

(defun dlt:safe-invoke-method (obj method args / ret)
  (setq ret (vl-catch-all-apply 'vlax-invoke-method (append (list obj method) args)))
  (if (vl-catch-all-error-p ret)
    nil
    T
  )
)

(defun dlt:apply-table-format (tableEname / tableObj rows cols row col)
  (setq tableObj (if tableEname (vlax-ename->vla-object tableEname)))
  (if (and tableObj (= (vla-get-ObjectName tableObj) "AcDbTable"))
    (progn
      ; 表格物件顏色改為 ByLayer（邊框顏色跟隨圖層）
      (vl-catch-all-apply 'vla-put-Color (list tableObj 256))

      (setq rows (vlax-get-property tableObj 'Rows)
            cols (vlax-get-property tableObj 'Columns))

      ; 套用欄寬
      (setq col 0)
      (while (< col cols)
        (dlt:safe-invoke-method tableObj 'SetColumnWidth (list col dlt:*cell-width*))
        (setq col (1+ col))
      )

      ; 套用列高
      (setq row 0)
      (while (< row rows)
        (dlt:safe-invoke-method tableObj 'SetRowHeight (list row dlt:*cell-height*))
        (setq row (1+ row))
      )

      ; 優先逐儲存格設定文字高；若版本不支援，再回退到列型設定
      (if (vlax-method-applicable-p tableObj 'SetCellTextHeight)
        (progn
          (setq row 0)
          (while (< row rows)
            (setq col 0)
            (while (< col cols)
              (dlt:safe-invoke-method tableObj 'SetCellTextHeight (list row col dlt:*text-height*))
              (setq col (1+ col))
            )
            (setq row (1+ row))
          )
        )
        (progn
          (dlt:safe-invoke-method tableObj 'SetTextHeight (list 1 dlt:*text-height*)) ; Title
          (dlt:safe-invoke-method tableObj 'SetTextHeight (list 2 dlt:*text-height*)) ; Header
          (dlt:safe-invoke-method tableObj 'SetTextHeight (list 4 dlt:*text-height*)) ; Data
        )
      )

      ; 文字樣式統一為微軟正黑體
      (if (dlt:ensure-msjh-text-style)
        (progn
          (dlt:safe-invoke-method tableObj 'SetTextStyle (list 1 dlt:*text-style-name*)) ; Title
          (dlt:safe-invoke-method tableObj 'SetTextStyle (list 2 dlt:*text-style-name*)) ; Header
          (dlt:safe-invoke-method tableObj 'SetTextStyle (list 4 dlt:*text-style-name*)) ; Data
        )
        (if (not dlt:*msjh-warning-shown*)
          (progn
            (prompt "\n警告：找不到微軟正黑體字型檔(msjh.ttc)，本次未強制套用該字型。")
            (setq dlt:*msjh-warning-shown* T)
          )
        )
      )

      (dlt:safe-invoke-method tableObj 'Update '())
    )
  )
  (if tableObj (vlax-release-object tableObj))
  ; 再次清除 true color override，確保是 ByLayer
  (dlt:set-entity-color-bylayer tableEname)
  T
)

(defun dlt:try-create-table-by-link (linkName insPt / patterns args ret ok entBefore newTable)
  (setq ok nil)
  (setq dlt:*last-removed-empty-rows* 0)
  (setq dlt:*last-create-failure-reason* nil)

  (if (not dlt:*has-command-s*)
    (progn
      (setq dlt:*last-create-failure-reason* "command_s_unavailable")
      nil
    )
    (progn
      (dlt:flush-command)
      (setq entBefore (entlast))
      (setq patterns
            (list
              ; 依你目前介面的實際流程：
              ; 1) -TABLE 2) 在「行數」題按 L(資料連結) 3) 輸入 Data Link 名稱 4) 指定插入點
              (list "_.-TABLE" "_L" linkName insPt)
              (list "_.-TABLE" "_L" linkName insPt "")
              (list "_.-TABLE" "L" linkName insPt)
              (list "_.-TABLE" "L" linkName insPt "")
            )
      )

      (foreach args patterns
        (if (not ok)
          (progn
            (setq ret (dlt:run-safe-command args))
            (if (vl-catch-all-error-p ret)
              (progn
                (setq dlt:*last-create-failure-reason* "command_error")
                (dlt:flush-command)
              )
              (progn
                (if (> (getvar "CMDACTIVE") 0)
                  (dlt:flush-command)
                )
                (setq newTable (dlt:get-new-table-ename-after entBefore))
                (if newTable
                  (progn
                    (setq ok T)
                    (dlt:apply-table-format newTable)
                  )
                  (setq dlt:*last-create-failure-reason* "no_table_created")
                )
              )
            )
          )
        )
      )

      (if (and ok dlt:*remove-empty-rows-after-create*)
        (progn
          (setq newTable (dlt:get-new-table-ename-after entBefore))
          (if newTable
            (setq dlt:*last-removed-empty-rows*
                  (dlt:remove-empty-rows-from-table newTable))
          )
        )
      )

      (if (and (not ok) (null dlt:*last-create-failure-reason*))
        (setq dlt:*last-create-failure-reason* "unknown")
      )
      ok
    )
  )
)

(defun dlt:get-datalink-names-safe (/ res)
  (setq res (vl-catch-all-apply 'dl:get-datalink-names '()))
  (if (vl-catch-all-error-p res)
    (progn
      (if (findfile "datalink_auto.lsp")
        (load "datalink_auto" nil)
      )
      (setq res (vl-catch-all-apply 'dl:get-datalink-names '()))
      (if (vl-catch-all-error-p res)
        'dlt_error
        res
      )
    )
    res
  )
)

(defun dlt:create-selected-link-table (names / input pickedText linkName insPt)
  (dlt:print-link-names names)
  (setq pickedText (dlt:get-input-by-pick))
  (if (/= (dlt:trim pickedText) "")
    (progn
      (setq input pickedText)
      (prompt (strcat "\n已選取文字: " pickedText))
    )
    (setq input (getstring T "\n輸入要建立表格的 Data Link 名稱或序號: "))
  )
  (setq linkName (dlt:resolve-link-name input names))
  (if (null linkName)
    (prompt "\n找不到指定的 Data Link，已取消。")
    (progn
      (setq insPt (getpoint "\n指定表格插入點: "))
      (if (null insPt)
        (prompt "\n未指定插入點，已取消。")
        (if (dlt:try-create-table-by-link linkName insPt)
          (progn
            (prompt (strcat "\n已建立表格: " linkName))
            (if (> dlt:*last-removed-empty-rows* 0)
              (prompt (strcat "\n已自動移除空白列: " (itoa dlt:*last-removed-empty-rows*)))
            )
          )
          (progn
            (prompt (strcat "\n建立表格失敗: " linkName))
            (prompt (strcat "\n失敗原因: " dlt:*last-create-failure-reason*))
            (if (= dlt:*last-create-failure-reason* "no_table_created")
              (prompt "\n此版本可能不支援以命令列直接由 Data Link 建表，請用 TABLE 對話框手動選「From a data link」。")
            )
            (if (= dlt:*last-create-failure-reason* "command_s_unavailable")
              (prompt "\n目前環境不支援 command-s，為避免卡住，已停止自動建表。")
            )
          )
        )
      )
    )
  )
)

(defun c:DLTABLEAUTO (/ names)
  (vl-load-com)
  (setvar "cmdecho" 0)
  (dlt:init-command-s-flag)

  (setq names (dlt:get-datalink-names-safe))
  (if (eq names 'dlt_error)
    (prompt "\n找不到 datalink_auto.lsp 或 dl:get-datalink-names，請先載入 datalink_auto.lsp。")
    (if (null names)
      (prompt "\n目前找不到任何 Data Link。")
      (progn
        (if (not dlt:*has-command-s*)
          (prompt "\n此 AutoCAD/LISP 環境不支援 command-s，為避免 TABLE 指令卡住，已停用 DLTABLEAUTO。")
          (dlt:create-selected-link-table names)
        )
      )
    )
  )
  (princ)
)

(defun c:AUTODATALINKTABLE ()
  (c:DLTABLEAUTO)
)

(defun c:DLLINKTABLE ()
  (c:DLTABLEAUTO)
)

(princ "\nDLTABLEAUTO / AUTODATALINKTABLE / DLLINKTABLE 載入完成。")
(princ)
