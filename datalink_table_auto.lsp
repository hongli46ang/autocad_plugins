(vl-load-com)

(setq dlt:*table-gap* 5000.0)
(setq dlt:*last-removed-empty-rows* 0)
; 預設關閉：避免刪空白列時誤拆到合併列
(setq dlt:*remove-empty-rows-after-create* nil)

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
    ((dlt:digits-only-p txt)
     (setq idx (atoi txt))
     (if (and (> idx 0) (<= idx (length names)))
       (nth (1- idx) names)
       nil
     )
    )
    (T (dlt:find-ci txt names))
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

(defun dlt:flush-command ()
  (while (> (getvar "CMDACTIVE") 0)
    (command "")
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

(defun dlt:try-create-table-by-link (linkName insPt / patterns args ret ok entBefore newTable)
  (setq ok nil)
  (setq dlt:*last-removed-empty-rows* 0)
  (setq entBefore (entlast))
  (setq patterns
        (list
          (list "_.-TABLE" "_FromDataLink" linkName insPt)
          (list "_.-TABLE" "_FromDataLink" linkName insPt "")
          (list "_.-TABLE" "_FromDataLink" "_Name" linkName insPt)
          (list "_.-TABLE" "_FromDataLink" "_Name" linkName insPt "")
          (list "_.TABLE" "_FromDataLink" linkName insPt)
          (list "_.TABLE" "_FromDataLink" linkName insPt "")
          (list "_.TABLE" "_FromDataLink" "_Name" linkName insPt)
          (list "_.TABLE" "_FromDataLink" "_Name" linkName insPt "")
        )
  )
  (foreach args patterns
    (if (not ok)
      (progn
        (setq ret (vl-catch-all-apply 'vl-cmdf args))
        (if (vl-catch-all-error-p ret)
          (dlt:flush-command)
          (setq ok T)
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
  ok
)

(defun dlt:offset-point-x (pt dx / z)
  (setq z (if (caddr pt) (caddr pt) 0.0))
  (list (+ (car pt) dx) (cadr pt) z)
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

(defun dlt:create-selected-link-table (names / input linkName insPt)
  (dlt:print-link-names names)
  (setq input (getstring T "\n輸入要建立表格的 Data Link 名稱或序號: "))
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
          (prompt (strcat "\n建立表格失敗: " linkName))
        )
      )
    )
  )
)

(defun dlt:create-all-link-tables (names / basePt idx linkName insPt okList failList okMsg)
  (setq basePt (getpoint "\n指定第一張表格插入點: "))
  (if (null basePt)
    (prompt "\n未指定插入點，已取消。")
    (progn
      (setq idx 0
            okList '()
            failList '())
      (foreach linkName names
        (setq insPt (dlt:offset-point-x basePt (* idx dlt:*table-gap*)))
        (if (dlt:try-create-table-by-link linkName insPt)
          (progn
            (setq okMsg linkName)
            (if (> dlt:*last-removed-empty-rows* 0)
              (setq okMsg
                    (strcat okMsg
                            " (移除空白列: "
                            (itoa dlt:*last-removed-empty-rows*)
                            ")"))
            )
            (setq okList (append okList (list okMsg)))
          )
          (setq failList (append failList (list linkName)))
        )
        (setq idx (1+ idx))
      )

      (prompt (strcat "\n批量建立完成，總數: " (itoa (length names))))
      (prompt (strcat "\n成功: " (itoa (length okList))
                      "，失敗: " (itoa (length failList))))
      (prompt (strcat "\n表格間距: " (rtos dlt:*table-gap* 2 0) " mm (X方向)"))

      (if okList
        (progn
          (prompt "\n--- 建立成功 ---")
          (foreach x okList
            (prompt (strcat "\n" x))
          )
        )
      )
      (if failList
        (progn
          (prompt "\n--- 建立失敗 ---")
          (foreach x failList
            (prompt (strcat "\n" x))
          )
        )
      )
    )
  )
)

(defun c:DLTABLEAUTO (/ names mode)
  (vl-load-com)
  (setvar "cmdecho" 0)

  (setq names (dlt:get-datalink-names-safe))
  (if (eq names 'dlt_error)
    (prompt "\n找不到 datalink_auto.lsp 或 dl:get-datalink-names，請先載入 datalink_auto.lsp。")
    (if (null names)
      (prompt "\n目前找不到任何 Data Link。")
      (progn
        (initget "S A")
        (setq mode (getkword "\n建立模式 [指定(S)/全部(A)] <S>: "))
        (if (or (null mode) (= mode "S"))
          (dlt:create-selected-link-table names)
          (dlt:create-all-link-tables names)
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
