(vl-load-com)

(setq dlt:*last-removed-empty-rows* 0)
; 預設關閉：避免刪空白列時誤拆到合併列
(setq dlt:*remove-empty-rows-after-create* nil)
(setq dlt:*last-create-failure-reason* nil)
(setq dlt:*last-created-table-ename* nil)
(setq dlt:*has-command-s* nil)
(setq dlt:*cell-width* 2025.0)
(setq dlt:*cell-height* 486.0)
(setq dlt:*text-height* 200.0)
(setq dlt:*text-style-name* "Standard")
(setq dlt:*cell-h-margin* 2.0)
(setq dlt:*cell-v-margin* 2.0)
(setq dlt:*cell-align-middle-center* 5)
(setq dlt:*text-font-file* "msjh.ttc")
(setq dlt:*text-font-candidates* '("msjh.ttc" "MSJH.TTC" "msjh.ttf" "MSJH.TTF"))
(setq dlt:*msjh-font-resolved* nil)
(setq dlt:*msjh-warning-shown* nil)
; 批次重綁時是否啟用「鄰近文字比對」回退
(setq dlt:*enable-near-text-fallback* T)

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

(defun dlt:ename-p (x)
  (= (type x) 'ENAME)
)

(defun dlt:valid-point-p (p)
  (and (= (type p) 'LIST)
       (>= (length p) 2)
       (numberp (car p))
       (numberp (cadr p))
       (or (= (length p) 2)
           (numberp (caddr p))))
)

(defun dlt:function-defined-p (fnName / key)
  (setq key (strcase (dlt:trim fnName)))
  (if (vl-position key (atoms-family 1))
    T
    nil
  )
)

(defun dlt:extract-datalink-name-from-detail (detail / p1 p2 out)
  (setq out "")
  (if (and detail (setq p1 (vl-string-search "\n" detail)))
    (progn
      (setq p2 (vl-string-search "\n" detail (1+ p1)))
      (if p2
        (setq out (substr detail (+ p1 2) (- p2 p1 1)))
        (setq out (substr detail (+ p1 2)))
      )
    )
  )
  (dlt:trim out)
)

(defun dlt:get-datalink-name-from-object-local (en / ed out)
  (setq out "")
  (if (dlt:ename-p en)
    (progn
      (setq ed (entget en))
      (if (and ed (= (cdr (assoc 0 ed)) "DATALINK"))
        (progn
          (setq out (dlt:extract-datalink-name-from-detail (cdr (assoc 301 ed))))
          (if (= out "")
            (setq out (dlt:trim (cdr (assoc 3 ed))))
          )
          (if (= out "")
            (setq out (dlt:trim (cdr (assoc 300 ed))))
          )
        )
      )
    )
  )
  out
)

(defun dlt:get-datalink-name-from-object-safe (en / out ret)
  (setq out "")
  (if (dlt:function-defined-p "DL:GET-DATALINK-NAME-FROM-OBJECT")
    (progn
      (setq ret (vl-catch-all-apply 'dl:get-datalink-name-from-object (list en)))
      (if (not (vl-catch-all-error-p ret))
        (setq out (dlt:trim ret))
      )
    )
  )
  (if (= out "")
    (setq out (dlt:get-datalink-name-from-object-local en))
  )
  out
)

(defun dlt:collect-ref-enames-from-ed (ed / out pr code val)
  (setq out '())
  (foreach pr ed
    (setq code (car pr)
          val  (cdr pr))
    (if (and (member code '(331 340 341 350 360))
             (dlt:ename-p val)
             (null (vl-position val out)))
      (setq out (cons val out))
    )
  )
  (reverse out)
)

(defun dlt:find-first-datalink-under-entity (root / queue visited en ed typ refs found steps)
  (setq queue (if (dlt:ename-p root) (list root) '())
        visited '()
        found nil
        steps 0)
  (while (and queue (null found) (< steps 500))
    (setq en (car queue)
          queue (cdr queue)
          steps (1+ steps))
    (if (and (dlt:ename-p en) (null (vl-position en visited)))
      (progn
        (setq visited (cons en visited)
              ed      (entget en)
              typ     (if ed (cdr (assoc 0 ed)) ""))
        (if (= typ "DATALINK")
          (setq found en)
          (progn
            (setq refs (dlt:collect-ref-enames-from-ed ed))
            (if refs
              (setq queue (append queue refs))
            )
          )
        )
      )
    )
  )
  found
)

(defun dlt:get-linked-datalink-name-from-table (tableEname / dlObj)
  (setq dlObj (dlt:find-first-datalink-under-entity tableEname))
  (if dlObj
    (dlt:get-datalink-name-from-object-safe dlObj)
    ""
  )
)

(defun dlt:get-entity-point (ename / ed p)
  (setq ed (if ename (entget ename))
        p  (if ed (cdr (assoc 10 ed))))
  (if (dlt:valid-point-p p)
    p
    nil
  )
)

(defun dlt:word-char-p (ch)
  (if (and ch (= (type ch) 'STR) (= (strlen ch) 1))
    (if (wcmatch (strcase ch) "[0-9A-Z_]")
      T
      nil
    )
    nil
  )
)

(defun dlt:text-has-name-with-boundary-p (txtNorm nameNorm / from pos nLen tLen bIdx aIdx bCh aCh ok)
  (setq ok nil
        from 0
        nLen (strlen nameNorm)
        tLen (strlen txtNorm))
  (while (and (not ok) (setq pos (vl-string-search nameNorm txtNorm from)))
    (setq bIdx pos
          aIdx (+ pos nLen 1)
          bCh  (if (>= bIdx 1) (substr txtNorm bIdx 1) "")
          aCh  (if (<= aIdx tLen) (substr txtNorm aIdx 1) ""))
    (if (and (or (= bCh "") (not (dlt:word-char-p bCh)))
             (or (= aCh "") (not (dlt:word-char-p aCh))))
      (setq ok T)
      (setq from (1+ pos))
    )
  )
  ok
)

(defun dlt:resolve-link-in-text (txt names / direct txtNorm best bestLen n nn)
  ; 鄰近文字比對僅接受「已存在 Data Link 名稱」。
  ; 允許在長字串中命中，但需符合前後邊界，避免亂配。
  (setq direct (dlt:resolve-link-by-text txt names))
  (if direct
    direct
    (progn
      (setq txtNorm (strcase (dlt:normalize-link-name txt))
            best    nil
            bestLen 0)
      (foreach n names
        (setq nn (strcase (dlt:normalize-link-name n)))
        (if (and (/= nn "")
                 (dlt:text-has-name-with-boundary-p txtNorm nn)
                 (> (strlen nn) bestLen))
          (setq best n
                bestLen (strlen nn))
        )
      )
      best
    )
  )
)

(defun dlt:collect-link-text-candidates (names / flt ss i en txt link p out)
  (setq out '())
  (setq flt (list '(0 . "TEXT,MTEXT,ATTRIB,ATTDEF") (cons 410 (getvar "CTAB"))))
  (setq ss (ssget "_X" flt))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq en   (ssname ss i)
              txt  (dlt:get-text-from-entity en)
              link (if (/= txt "") (dlt:resolve-link-in-text txt names) nil)
              p    (if link (dlt:get-entity-point en) nil))
        (if (and link p)
          ; 項目格式: (linkName point text ename)
          (setq out (cons (list link p txt en) out))
        )
        (setq i (1+ i))
      )
    )
  )
  out
)

(defun dlt:short-text (s maxLen / t)
  (setq t (dlt:trim s))
  (if (and (> maxLen 3) (> (strlen t) maxLen))
    (strcat (substr t 1 (- maxLen 3)) "...")
    t
  )
)

(defun dlt:best-near-text-candidate-for-table (tableEname candidates / tp best bestD c cp d)
  (setq tp (or (dlt:get-table-insertion-point tableEname)
               (dlt:get-entity-point tableEname))
        best nil
        bestD nil)
  (if (and (dlt:valid-point-p tp) candidates)
    (foreach c candidates
      (setq cp (cadr c))
      (if (dlt:valid-point-p cp)
        (progn
          (setq d (distance tp cp))
          (if (or (null bestD) (< d bestD))
            (setq best c
                  bestD d)
          )
        )
      )
    )
  )
  best
)

(defun dlt:normalize-link-name (s)
  (if (dlt:function-defined-p "DL:SANITIZE-NAME")
    (dl:sanitize-name s)
    (dlt:trim s)
  )
)

(defun dlt:resolve-link-by-text (txt names / txtNorm)
  (setq txt (dlt:trim txt))
  (if (= txt "")
    nil
    (or
      (dlt:find-ci txt names)
      (progn
        (setq txtNorm (strcase (dlt:normalize-link-name txt)))
        (vl-some
          '(lambda (n)
             (if (= (strcase (dlt:normalize-link-name n)) txtNorm)
               n
             )
           )
          names
        )
      )
    )
  )
)

(defun dlt:get-table-handle (ename / ed h)
  (setq ed (if ename (entget ename))
        h  (if ed (cdr (assoc 5 ed))))
  (if h h "")
)

(defun dlt:get-table-insertion-point (tableEname / obj p)
  (setq obj (if tableEname (vlax-ename->vla-object tableEname)))
  (if obj
    (progn
      (setq p (vl-catch-all-apply 'vlax-get-property (list obj 'InsertionPoint)))
      (if (= (type obj) 'VLA-OBJECT)
        (vl-catch-all-apply 'vlax-release-object (list obj))
      )
      (if (vl-catch-all-error-p p)
        nil
        (progn
          (if (= (type p) 'VARIANT)
            (setq p (vlax-variant-value p))
          )
          (setq p
                (cond
                  ((= (type p) 'LIST) p)
                  ((= (type p) 'SAFEARRAY) (vlax-safearray->list p))
                  (T nil)
                )
          )
          (if (dlt:valid-point-p p)
            p
            nil
          )
        )
      )
    )
    nil
  )
)

(defun dlt:copy-vla-property-safe (src dst prop / v)
  (if (and src dst
           (vlax-property-available-p src prop)
           (vlax-property-available-p dst prop T))
    (progn
      (setq v (vl-catch-all-apply 'vlax-get-property (list src prop)))
      (if (not (vl-catch-all-error-p v))
        (vl-catch-all-apply 'vlax-put-property (list dst prop v))
      )
    )
  )
  T
)

(defun dlt:copy-table-basic-props (srcEname dstEname / src dst)
  (setq src (if srcEname (vlax-ename->vla-object srcEname))
        dst (if dstEname (vlax-ename->vla-object dstEname)))
  (if (and src dst)
    (progn
      (dlt:copy-vla-property-safe src dst 'Layer)
      (dlt:copy-vla-property-safe src dst 'Linetype)
      (dlt:copy-vla-property-safe src dst 'LinetypeScale)
      (dlt:copy-vla-property-safe src dst 'Lineweight)
      (dlt:copy-vla-property-safe src dst 'Color)
      (dlt:copy-vla-property-safe src dst 'Rotation)
    )
  )
  (if (= (type src) 'VLA-OBJECT)
    (vl-catch-all-apply 'vlax-release-object (list src))
  )
  (if (= (type dst) 'VLA-OBJECT)
    (vl-catch-all-apply 'vlax-release-object (list dst))
  )
  T
)

(defun dlt:guess-link-name-from-table (tableEname names / tableObj rows cols maxRows maxCols row col txt hit)
  (setq tableObj (if tableEname (vlax-ename->vla-object tableEname))
        hit nil)
  (if (and tableObj (= (vla-get-ObjectName tableObj) "AcDbTable"))
    (progn
      (setq rows (vlax-get-property tableObj 'Rows)
            cols (vlax-get-property tableObj 'Columns)
            maxRows (min rows 12)
            maxCols (min cols 8))
      ; 以左上區域做快速比對（標題/表頭通常在這裡）
      (setq row 0)
      (while (and (< row maxRows) (null hit))
        (setq col 0)
        (while (and (< col maxCols) (null hit))
          (setq txt (dlt:get-cell-text-safe tableObj row col))
          (if (/= txt "")
            (setq hit (dlt:resolve-link-by-text txt names))
          )
          (setq col (1+ col))
        )
        (setq row (1+ row))
      )
    )
  )
  (if (= (type tableObj) 'VLA-OBJECT)
    (vl-catch-all-apply 'vlax-release-object (list tableObj))
  )
  hit
)

(defun dlt:rebuild-table-by-link (tableEname linkName / insPt ok newTable)
  (setq insPt (dlt:get-table-insertion-point tableEname))
  (if (null insPt)
    (progn
      (setq dlt:*last-create-failure-reason* "no_insertion_point")
      nil
    )
    (progn
      (setq ok (dlt:try-create-table-by-link linkName insPt))
      (if ok
        (progn
          (setq newTable dlt:*last-created-table-ename*)
          (if newTable
            (dlt:copy-table-basic-props tableEname newTable)
          )
          (entdel tableEname)
          T
        )
        nil
      )
    )
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

(defun dlt:set-entity-table-props (ename / ed)
  (if ename
    (progn
      (setq ed (entget ename))
      (if ed
        (progn
          (setq ed
                (vl-remove-if
                  '(lambda (x) (or (= (car x) 420) (= (car x) 430) (= (car x) 440)))
                  ed))
          ; 框線顏色：ByLayer
          (setq ed (dlt:set-dxf ed 62 256))
          ; 框線線型：ByBlock
          (setq ed (dlt:set-dxf ed 6 "ByBlock"))
          ; 框線線粗：ByBlock
          (setq ed (dlt:set-dxf ed 370 -2))
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

(defun dlt:safe-put-property (obj prop value / ret)
  (setq ret (vl-catch-all-apply 'vlax-put-property (list obj prop value)))
  (if (vl-catch-all-error-p ret)
    nil
    T
  )
)

(defun dlt:apply-gridcolor-bylayer (tableObj rows cols / clr row col)
  ; 先嘗試用 AcCmColor 物件設定（較完整）
  (setq clr (vl-catch-all-apply 'vla-get-TrueColor (list tableObj)))
  (if (not (vl-catch-all-error-p clr))
    (progn
      (dlt:safe-put-property clr 'ColorIndex 256)

      ; 列型層級（Title/Header/Data）框線顏色
      (dlt:safe-invoke-method tableObj 'SetGridColor (list 1 63 clr))
      (dlt:safe-invoke-method tableObj 'SetGridColor (list 2 63 clr))
      (dlt:safe-invoke-method tableObj 'SetGridColor (list 4 63 clr))
      (dlt:safe-invoke-method tableObj 'SetGridColor (list 7 63 clr))

      ; 儲存格層級覆寫：把每格框線都設回 ByLayer
      (setq row 0)
      (while (< row rows)
        (setq col 0)
        (while (< col cols)
          (dlt:safe-invoke-method tableObj 'SetCellGridColor (list row col 15 clr))
          (dlt:safe-invoke-method tableObj 'SetCellGridColor (list row col 63 clr))
          (setq col (1+ col))
        )
        (setq row (1+ row))
      )
      (if (= (type clr) 'VLA-OBJECT)
        (vl-catch-all-apply 'vlax-release-object (list clr))
      )
    )
  )

  ; 部分版本只支援整數色碼 API，做相容回退
  (dlt:safe-invoke-method tableObj 'SetGridColor2 (list 1 63 256))
  (dlt:safe-invoke-method tableObj 'SetGridColor2 (list 2 63 256))
  (dlt:safe-invoke-method tableObj 'SetGridColor2 (list 4 63 256))
  (dlt:safe-invoke-method tableObj 'SetGridColor2 (list 7 63 256))

  (setq row 0)
  (while (< row rows)
    (setq col 0)
    (while (< col cols)
      (dlt:safe-invoke-method tableObj 'SetCellGridColor2 (list row col 15 256))
      (dlt:safe-invoke-method tableObj 'SetCellGridColor2 (list row col 63 256))
      (setq col (1+ col))
    )
    (setq row (1+ row))
  )
  T
)

(defun dlt:apply-table-format (tableEname / tableObj rows cols row col)
  (setq tableObj (if tableEname (vlax-ename->vla-object tableEname)))
  (if (and tableObj (= (vla-get-ObjectName tableObj) "AcDbTable"))
    (progn
      ; 框線顏色：ByLayer
      (vl-catch-all-apply 'vla-put-Color (list tableObj 256))
      ; 框線線型/線粗：ByBlock
      (vl-catch-all-apply 'vla-put-Linetype (list tableObj "ByBlock"))
      (vl-catch-all-apply 'vla-put-Lineweight (list tableObj -2))

      (setq rows (vlax-get-property tableObj 'Rows)
            cols (vlax-get-property tableObj 'Columns))

      ; 儲存格邊距
      (dlt:safe-put-property tableObj 'HorzCellMargin dlt:*cell-h-margin*)
      (dlt:safe-put-property tableObj 'VertCellMargin dlt:*cell-v-margin*)

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

      ; 對齊方式：正中
      (if (vlax-method-applicable-p tableObj 'SetCellAlignment)
        (progn
          (setq row 0)
          (while (< row rows)
            (setq col 0)
            (while (< col cols)
              (dlt:safe-invoke-method tableObj 'SetCellAlignment (list row col dlt:*cell-align-middle-center*))
              (setq col (1+ col))
            )
            (setq row (1+ row))
          )
        )
        (progn
          (dlt:safe-invoke-method tableObj 'SetAlignment (list 1 dlt:*cell-align-middle-center*)) ; Title
          (dlt:safe-invoke-method tableObj 'SetAlignment (list 2 dlt:*cell-align-middle-center*)) ; Header
          (dlt:safe-invoke-method tableObj 'SetAlignment (list 4 dlt:*cell-align-middle-center*)) ; Data
        )
      )

      ; 文字樣式：Standard
      (if (tblsearch "STYLE" dlt:*text-style-name*)
        (progn
          (dlt:safe-invoke-method tableObj 'SetTextStyle (list 1 dlt:*text-style-name*)) ; Title
          (dlt:safe-invoke-method tableObj 'SetTextStyle (list 2 dlt:*text-style-name*)) ; Header
          (dlt:safe-invoke-method tableObj 'SetTextStyle (list 4 dlt:*text-style-name*)) ; Data
        )
        (prompt "\n警告：找不到文字樣式 Standard，本次未強制套用文字樣式。")
      )

      ; 文字顏色：ByBlock（0）
      (dlt:safe-invoke-method tableObj 'SetColor (list 1 0)) ; Title
      (dlt:safe-invoke-method tableObj 'SetColor (list 2 0)) ; Header
      (dlt:safe-invoke-method tableObj 'SetColor (list 4 0)) ; Data
      (dlt:safe-invoke-method tableObj 'SetContentColor (list 1 0)) ; Title
      (dlt:safe-invoke-method tableObj 'SetContentColor (list 2 0)) ; Header
      (dlt:safe-invoke-method tableObj 'SetContentColor (list 4 0)) ; Data

      ; 文字旋轉：0
      (dlt:safe-invoke-method tableObj 'SetTextRotation (list 1 0.0)) ; Title
      (dlt:safe-invoke-method tableObj 'SetTextRotation (list 2 0.0)) ; Header
      (dlt:safe-invoke-method tableObj 'SetTextRotation (list 4 0.0)) ; Data

      ; 框線顏色：強制清掉儲存格黑色覆寫，統一回 ByLayer
      (dlt:apply-gridcolor-bylayer tableObj rows cols)

      (dlt:safe-invoke-method tableObj 'Update '())
    )
  )
  (if tableObj (vlax-release-object tableObj))
  ; 再次清除 true color override，並套用框線 ByLayer + 線型/線粗 ByBlock
  (dlt:set-entity-table-props tableEname)
  T
)

(defun dlt:try-create-table-by-link (linkName insPt / patterns args ret ok entBefore newTable)
  (setq ok nil)
  (setq dlt:*last-removed-empty-rows* 0)
  (setq dlt:*last-create-failure-reason* nil)
  (setq dlt:*last-created-table-ename* nil)

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
                    (setq dlt:*last-created-table-ename* newTable)
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
          (if dlt:*last-created-table-ename*
            (setq dlt:*last-removed-empty-rows*
                  (dlt:remove-empty-rows-from-table dlt:*last-created-table-ename*))
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

(defun dlt:ss->ename-list (ss / i out)
  (setq out '())
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq out (cons (ssname ss i) out))
        (setq i (1+ i))
      )
    )
  )
  (reverse out)
)

(defun dlt:select-target-tables (/ flt ss)
  (setq flt (list '(0 . "ACAD_TABLE") (cons 410 (getvar "CTAB"))))
  (prompt "\n選取要重綁的表格 (Enter=目前頁籤全部 ACAD_TABLE): ")
  (setq ss (ssget flt))
  (if ss
    ss
    (ssget "_X" flt)
  )
)

(defun dlt:print-batch-lines (title lines)
  (if lines
    (progn
      (prompt title)
      (foreach s lines
        (prompt (strcat "\n" s))
      )
    )
  )
)

(defun dlt:pick-one-table-entity (/ ent en ed)
  (setq ent (entsel "\n選取要重綁的舊表格 (Enter 結束): "))
  (if (null ent)
    nil
    (progn
      (setq en (car ent)
            ed (if en (entget en)))
      (if (and ed (= (cdr (assoc 0 ed)) "ACAD_TABLE"))
        en
        'dlt_retry
      )
    )
  )
)

(defun dlt:ask-link-name-for-manual-rebind (names / input pickedText linkName)
  (setq pickedText (dlt:get-input-by-pick))
  (if (/= (dlt:trim pickedText) "")
    (progn
      (setq input pickedText)
      (prompt (strcat "\n已選取文字: " pickedText))
    )
    (setq input (getstring T "\n輸入對應 Data Link 名稱或序號: "))
  )
  (setq linkName (dlt:resolve-link-name input names))
  (if (null linkName)
    (progn
      (prompt "\n找不到對應 Data Link，該表格本次略過。")
      nil
    )
    linkName
  )
)

(defun dlt:batch-rebind-existing-tables (names / ss enames total done okCount failCount skipCount
                                                en h linkName reason linkedRaw sourceTag
                                                nearCand nearTxt textCandidates
                                                okLines failLines skipLines)
  (setq ss (dlt:select-target-tables))
  (if (null ss)
    (prompt "\n目前頁籤找不到可重綁的表格。")
    (progn
      (if dlt:*enable-near-text-fallback*
        (progn
          (setq textCandidates (dlt:collect-link-text-candidates names))
          (prompt (strcat "\n目前頁籤可用盤名字文字候選: " (itoa (length textCandidates))))
        )
        (progn
          (setq textCandidates '())
          (prompt "\n目前已關閉鄰近文字比對（僅用舊表 Data Link 與表格內容）。")
        )
      )

      (setq enames    (dlt:ss->ename-list ss)
            total     (length enames)
            done      0
            okCount   0
            failCount 0
            skipCount 0
            okLines   '()
            failLines '()
            skipLines '())

      (foreach en enames
        (setq done (1+ done)
              h    (dlt:get-table-handle en))
        (prompt (strcat "\n處理中 " (itoa done) "/" (itoa total) "，Handle: " h))

        (setq linkName nil
              reason   ""
              sourceTag ""
              linkedRaw (dlt:get-linked-datalink-name-from-table en))
        (if (/= (dlt:trim linkedRaw) "")
          (progn
            (setq linkName (dlt:resolve-link-by-text linkedRaw names))
            (if linkName
              (setq sourceTag (strcat "[來源:舊表 Data Link=" linkedRaw "]"))
              (setq reason (strcat "舊表 Data Link: " linkedRaw "，目前 DWG 找不到同名"))
            )
          )
        )

        (if (null linkName)
          (progn
            ; 回退 1：舊表若查不到 Data Link，再嘗試用表格內容推斷
            (setq linkName (dlt:guess-link-name-from-table en names))
            (if linkName
              (setq sourceTag "[來源:表格內容推斷]")
            )
          )
        )

        (if (and (null linkName) dlt:*enable-near-text-fallback*)
          (progn
            ; 回退 2：用表格附近文字推斷
            (setq nearCand (dlt:best-near-text-candidate-for-table en textCandidates))
            (if nearCand
              (progn
                (setq linkName (car nearCand)
                      nearTxt  (dlt:short-text (caddr nearCand) 40)
                      sourceTag (strcat "[來源:附近文字=" nearTxt "]"))
              )
              (if (/= (dlt:trim linkedRaw) "")
                (setq reason (strcat "舊表 Data Link: " linkedRaw "，目前 DWG 找不到同名"))
              )
            )
          )
        )

        (if (and (null linkName) (= reason ""))
          (if dlt:*enable-near-text-fallback*
            (setq reason "表格未綁 Data Link，表格內容與附近文字皆無法對應現有 Data Link")
            (setq reason "表格未綁 Data Link，且表格內容無法對應現有 Data Link（鄰近文字比對關閉）")
          )
        )

        (if (and (null linkName) (/= (dlt:trim linkedRaw) "") (= reason ""))
          (setq reason (strcat "舊表 Data Link: " linkedRaw "，目前 DWG 找不到同名"))
        )

        (if (and (null linkName) (= (dlt:trim linkedRaw) "") (= reason ""))
          (setq reason "表格未綁 Data Link")
        )

        (if (null linkName)
          (progn
            (setq skipCount (1+ skipCount))
            (setq skipLines (cons (strcat h "  (" reason ")") skipLines))
          )
          (if (dlt:rebuild-table-by-link en linkName)
            (progn
              (setq okCount (1+ okCount))
              (setq okLines (cons (strcat h " -> " linkName "  " sourceTag) okLines))
            )
            (progn
              (setq failCount (1+ failCount)
                    reason (if dlt:*last-create-failure-reason*
                             dlt:*last-create-failure-reason*
                             "unknown"))
              (setq failLines (cons (strcat h " -> " linkName "  (" reason ")") failLines))
            )
          )
        )
      )

      (prompt "\n--- 批次重綁完成 ---")
      (prompt (strcat "\n總表格數: " (itoa total)
                      "，成功: " (itoa okCount)
                      "，失敗: " (itoa failCount)
                      "，略過: " (itoa skipCount)))
      (dlt:print-batch-lines "\n--- 成功清單 ---" (reverse okLines))
      (dlt:print-batch-lines "\n--- 失敗清單 ---" (reverse failLines))
      (dlt:print-batch-lines "\n--- 略過清單 ---" (reverse skipLines))
      (if (> failCount 0)
        (prompt "\n若失敗原因為 no_table_created，請確認此版本是否支援命令列 -TABLE 由 Data Link 建表。")
      )
    )
  )
)

(defun c:DLTABLEREBINDMANUAL (/ names tableEn linkName handle reason
                                done okCount failCount skipCount
                                okLines failLines skipLines)
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
          (prompt "\n此 AutoCAD/LISP 環境不支援 command-s，為避免 TABLE 指令卡住，已停用 DLTABLEREBINDMANUAL。")
          (progn
            (dlt:print-link-names names)
            (prompt "\n開始手動重綁：每次先選舊表，再選盤名文字（或手動輸入名稱/序號）。")
            (setq done      nil
                  okCount   0
                  failCount 0
                  skipCount 0
                  okLines   '()
                  failLines '()
                  skipLines '())

            (while (not done)
              (setq tableEn (dlt:pick-one-table-entity))
              (cond
                ((null tableEn)
                 (setq done T)
                )
                ((eq tableEn 'dlt_retry)
                 (prompt "\n所選物件不是表格，請重選。")
                )
                (T
                 (setq handle   (dlt:get-table-handle tableEn)
                       linkName (dlt:ask-link-name-for-manual-rebind names))
                 (if (null linkName)
                   (progn
                     (setq skipCount (1+ skipCount))
                     (setq skipLines (cons (strcat handle "  (找不到指定 Data Link)") skipLines))
                   )
                   (if (dlt:rebuild-table-by-link tableEn linkName)
                     (progn
                       (setq okCount (1+ okCount))
                       (setq okLines (cons (strcat handle " -> " linkName) okLines))
                     )
                     (progn
                       (setq failCount (1+ failCount)
                             reason (if dlt:*last-create-failure-reason*
                                      dlt:*last-create-failure-reason*
                                      "unknown"))
                       (setq failLines (cons (strcat handle " -> " linkName "  (" reason ")") failLines))
                     )
                   )
                 )
                )
              )
            )

            (prompt "\n--- 手動重綁完成 ---")
            (prompt (strcat "\n成功: " (itoa okCount)
                            "，失敗: " (itoa failCount)
                            "，略過: " (itoa skipCount)))
            (dlt:print-batch-lines "\n--- 成功清單 ---" (reverse okLines))
            (dlt:print-batch-lines "\n--- 失敗清單 ---" (reverse failLines))
            (dlt:print-batch-lines "\n--- 略過清單 ---" (reverse skipLines))
          )
        )
      )
    )
  )
  (princ)
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

(defun c:DLTABLEREBIND (/ names)
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
          (prompt "\n此 AutoCAD/LISP 環境不支援 command-s，為避免 TABLE 指令卡住，已停用 DLTABLEREBIND。")
          (dlt:batch-rebind-existing-tables names)
        )
      )
    )
  )
  (princ)
)

(defun c:DLTABLEREBINDBATCH ()
  (c:DLTABLEREBIND)
)

(defun c:DLTABLEREBINDPICK ()
  (c:DLTABLEREBINDMANUAL)
)

(defun c:AUTODATALINKTABLE ()
  (c:DLTABLEAUTO)
)

(defun c:DLLINKTABLE ()
  (c:DLTABLEAUTO)
)

(defun c:AUTODATALINKTABLEREBIND ()
  (c:DLTABLEREBIND)
)

(defun c:DLLINKTABLEREBIND ()
  (c:DLTABLEREBIND)
)

(defun c:AUTODATALINKTABLEREBINDPICK ()
  (c:DLTABLEREBINDMANUAL)
)

(defun c:DLLINKTABLEREBINDPICK ()
  (c:DLTABLEREBINDMANUAL)
)

(princ "\nDLTABLEAUTO / AUTODATALINKTABLE / DLLINKTABLE / DLTABLEREBIND / DLTABLEREBINDBATCH / AUTODATALINKTABLEREBIND / DLLINKTABLEREBIND / DLTABLEREBINDMANUAL / DLTABLEREBINDPICK / AUTODATALINKTABLEREBINDPICK / DLLINKTABLEREBINDPICK 載入完成。")
(princ)
