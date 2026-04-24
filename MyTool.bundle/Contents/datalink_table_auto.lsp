(vl-load-com)

;; DLTABLEUI
;; 流程：啟動 TABLE 互動介面 -> 使用者選 Data Link / 插入點 -> 指令結束後自動格式化新表格
;; 備用：TBFIX 可手動選一個既有表格並套用相同格式

(if (and (boundp 'dlt:*update-reactor*) dlt:*update-reactor* (vlr-added-p dlt:*update-reactor*))
  (vlr-remove dlt:*update-reactor*)
)
(setq dlt:*table-before-ent* nil)
(setq dlt:*table-reactor* nil)
(setq dlt:*update-reactor* nil)
(setq dlt:*pending-format-mode* nil)
(setq dlt:*pending-table-ename* nil)
(setq dlt:*pending-format-source* "")
(setq dlt:*pending-format-remaining* 0)

;; ====== 你主要改這幾個值 ======
(setq dlt:*table-row-height* 486.0)     ;; 列高
(setq dlt:*table-col-width* 2025.0)     ;; 欄寬
(setq dlt:*table-text-height* 200.0)    ;; 文字高度
(setq dlt:*table-text-style* "Standard") ;; 文字型式
(setq dlt:*table-style-name* "DLT_DATALINK_TABLE") ;; 自動套用的表格型式
(setq dlt:*table-alignment* 5)          ;; 5 = 置中
(setq dlt:*table-color-index* 256)      ;; 256 = ByLayer
(setq dlt:*table-linetype* "ByLayer")
(setq dlt:*table-lineweight* -1)        ;; -1 = ByLayer
(setq dlt:*table-grid-types-all* 63)    ;; 所有格線
(setq dlt:*format-delay-ms* 1200)       ;; 每次重套格式前等待時間
(setq dlt:*format-repeat-count* 6)      ;; Data Link 覆蓋較慢時連續補格式次數

;; Data Link 更新選項 bit：
;; SkipFormat=131072, UpdateRowHeight=262144, UpdateColumnWidth=524288,
;; AllowSourceUpdate=1048576, OverwriteContentModifiedAfterUpdate=4194304,
;; OverwriteFormatModifiedAfterUpdate=8388608
(setq dlt:*datalink-skip-format-bit* 131072)
(setq dlt:*datalink-update-row-height-bit* 262144)
(setq dlt:*datalink-update-col-width-bit* 524288)
(setq dlt:*datalink-allow-source-update-bit* 1048576)
(setq dlt:*datalink-overwrite-content-bit* 4194304)
(setq dlt:*datalink-overwrite-format-bit* 8388608)

(defun dlt:remove-reactor ()
  (if (and dlt:*table-reactor* (not (vlr-added-p dlt:*table-reactor*)))
    (setq dlt:*table-reactor* nil)
  )
  (if (and dlt:*table-reactor* (vlr-added-p dlt:*table-reactor*))
    (vlr-remove dlt:*table-reactor*)
  )
  (setq dlt:*table-reactor* nil)
)

(defun dlt:table-ename-p (e / ed objname)
  (if (and e (setq ed (entget e)))
    (or (= (cdr (assoc 0 ed)) "ACAD_TABLE")
        (and
          (not (vl-catch-all-error-p
                 (setq objname
                   (vl-catch-all-apply
                     '(lambda () (vla-get-ObjectName (vlax-ename->vla-object e)))
                   )
                 )
               )
          )
          (= objname "AcDbTable")
        )
    )
    nil
  )
)

(defun dlt:find-new-table-after (before / e found last)
  (setq found nil)
  (setq e (if before (entnext before) (entnext)))
  (while e
    (if (dlt:table-ename-p e)
      (setq found e)
    )
    (setq e (entnext e))
  )
  ;; 保險：如果掃不到，就檢查最後一個物件
  (if (null found)
    (progn
      (setq last (entlast))
      (if (dlt:table-ename-p last)
        (setq found last)
      )
    )
  )
  found
)

(defun dlt:function-defined-p (fn / name)
  (if (= (type fn) 'SYM)
    (progn
      (setq name (strcase (vl-symbol-name fn)))
      (if (vl-position name (atoms-family 1)) T nil)
    )
    T
  )
)

(defun dlt:safe-call (fn args)
  (if (dlt:function-defined-p fn)
    (vl-catch-all-apply fn args)
    nil
  )
)

(defun dlt:safe-put (obj prop val)
  (if obj
    (vl-catch-all-apply 'vlax-put-property (list obj prop val))
    nil
  )
)

(defun dlt:safe-method (obj method args)
  (if obj
    (vl-catch-all-apply 'vlax-invoke-method (append (list obj method) args))
    nil
  )
)

(defun dlt:try-method (obj method args / res)
  (if obj
    (progn
      (setq res (vl-catch-all-apply 'vlax-invoke-method (append (list obj method) args)))
      (if (vl-catch-all-error-p res) nil T)
    )
    nil
  )
)

(defun dlt:try-put (obj prop val / res)
  (if obj
    (progn
      (setq res (vl-catch-all-apply 'vlax-put-property (list obj prop val)))
      (if (vl-catch-all-error-p res) nil T)
    )
    nil
  )
)

(defun dlt:get-active-document (/ acad)
  (setq acad (vlax-get-acad-object))
  (if acad
    (vla-get-ActiveDocument acad)
    nil
  )
)

(defun dlt:ename-p (x)
  (= (type x) 'ENAME)
)

(defun dlt:trim (s)
  (if s
    (vl-string-trim " \t\r\n" s)
    ""
  )
)

(defun dlt:release (obj)
  (if (and obj (= (type obj) 'VLA-OBJECT))
    (vl-catch-all-apply 'vlax-release-object (list obj))
  )
)

(defun dlt:bit-clear (flags bitMask)
  (if (/= (logand flags bitMask) 0)
    (- flags bitMask)
    flags
  )
)

(defun dlt:normalize-datalink-update-flag91 (flags / out)
  (setq out (if flags flags 0))
  ;; Data-only update: keep linked content updating, but do not import Excel/table formatting.
  (setq out (logior out dlt:*datalink-skip-format-bit*))
  (setq out (logior out dlt:*datalink-overwrite-content-bit*))
  (setq out (dlt:bit-clear out dlt:*datalink-update-row-height-bit*))
  (setq out (dlt:bit-clear out dlt:*datalink-update-col-width-bit*))
  (setq out (dlt:bit-clear out dlt:*datalink-allow-source-update-bit*))
  (setq out (dlt:bit-clear out dlt:*datalink-overwrite-format-bit*))
  out
)

(defun dlt:set-first-dxf (ed code val / out done pair)
  (setq out '()
        done nil)
  (foreach pair ed
    (if (and (not done) (= (car pair) code))
      (progn
        (setq out (cons (cons code val) out))
        (setq done T)
      )
      (setq out (cons pair out))
    )
  )
  (reverse (if done out (cons (cons code val) out)))
)

(defun dlt:set-all-dxf (ed code val / out pair)
  (setq out '())
  (foreach pair ed
    (if (= (car pair) code)
      (setq out (cons (cons code val) out))
      (setq out (cons pair out))
    )
  )
  (reverse out)
)

(defun dlt:get-existing-datalink-dict-ename (/ nod rec en)
  (setq nod (namedobjdict)
        rec (if (dlt:ename-p nod) (dictsearch nod "ACAD_DATALINK")))
  (if rec
    (progn
      (setq en (or (cdr (assoc 360 rec))
                   (cdr (assoc 350 rec))))
      (if (dlt:ename-p en) en nil)
    )
    nil
  )
)

(setq dlt:*dict-visited* nil)
(setq dlt:*dict-vla-visited* nil)

(defun dlt:safe-vla-objectname (obj / r)
  (setq r (vl-catch-all-apply 'vla-get-ObjectName (list obj)))
  (if (vl-catch-all-error-p r)
    ""
    (strcase (dlt:trim r))
  )
)

(defun dlt:safe-vla-handle (obj / r)
  (setq r (vl-catch-all-apply 'vla-get-Handle (list obj)))
  (if (vl-catch-all-error-p r)
    ""
    (dlt:trim r)
  )
)

(defun dlt:vla-object->ename-safe (obj / r)
  (setq r (vl-catch-all-apply 'vlax-vla-object->ename (list obj)))
  (if (and (not (vl-catch-all-error-p r)) (dlt:ename-p r))
    r
    nil
  )
)

(defun dlt:collect-datalinks-from-dict (dictEname / rec key objEn objEd typ handle out)
  (setq out '())
  (if (dlt:ename-p dictEname)
    (progn
      (setq handle (cdr (assoc 5 (entget dictEname))))
      (if (and handle (member handle dlt:*dict-visited*))
        out
        (progn
          (if handle
            (setq dlt:*dict-visited* (cons handle dlt:*dict-visited*))
          )
          (setq rec (dictnext dictEname T))
          (while rec
            (setq key   (cdr (assoc 3 rec))
                  objEn (or (cdr (assoc 360 rec))
                            (cdr (assoc 350 rec)))
                  objEd (if (dlt:ename-p objEn) (entget objEn))
                  typ   (if objEd (cdr (assoc 0 objEd))))
            (cond
              ((= typ "DATALINK")
               (setq out (cons objEn out))
              )
              ((= typ "DICTIONARY")
               (setq out (append (dlt:collect-datalinks-from-dict objEn) out))
              )
            )
            (setq rec (dictnext dictEname))
          )
        )
      )
    )
  )
  out
)

(defun dlt:collect-datalinks-from-vla-dict (dictObj / out item oname handle en)
  (setq out '())
  (if (and dictObj (= (type dictObj) 'VLA-OBJECT))
    (progn
      (setq handle (dlt:safe-vla-handle dictObj))
      (if (and (/= handle "") (member handle dlt:*dict-vla-visited*))
        out
        (progn
          (if (/= handle "")
            (setq dlt:*dict-vla-visited* (cons handle dlt:*dict-vla-visited*))
          )
          (vlax-for item dictObj
            (setq oname (dlt:safe-vla-objectname item))
            (cond
              ((wcmatch oname "*DICTIONARY*")
               (foreach en (dlt:collect-datalinks-from-vla-dict item)
                 (setq out (dlt:add-unique-ename en out))
               )
              )
              ((wcmatch oname "*DATALINK*")
               (setq en (dlt:vla-object->ename-safe item))
               (if en
                 (setq out (dlt:add-unique-ename en out))
               )
              )
            )
          )
        )
      )
    )
  )
  out
)

(defun dlt:get-datalink-enames-vla (/ doc db dicts dictObj out)
  (setq out '()
        doc (dlt:get-active-document))
  (if doc
    (progn
      (setq db (vl-catch-all-apply 'vla-get-Database (list doc)))
      (if (not (vl-catch-all-error-p db))
        (progn
          (setq dicts (vl-catch-all-apply 'vla-get-Dictionaries (list db)))
          (if (not (vl-catch-all-error-p dicts))
            (progn
              (setq dictObj (vl-catch-all-apply 'vla-Item (list dicts "ACAD_DATALINK")))
              (if (not (vl-catch-all-error-p dictObj))
                (progn
                  (setq dlt:*dict-vla-visited* nil)
                  (setq out (dlt:collect-datalinks-from-vla-dict dictObj))
                  (dlt:release dictObj)
                )
              )
              (dlt:release dicts)
            )
          )
        )
      )
    )
  )
  out
)

(defun dlt:add-unique-ename (en lst)
  (if (vl-position en lst)
    lst
    (cons en lst)
  )
)

(defun dlt:get-nod-dictionary-enames (/ nod rec en recEd out)
  (setq out '()
        nod (namedobjdict))
  (if (dlt:ename-p nod)
    (progn
      (setq rec (dictnext nod T))
      (while rec
        (setq en (or (cdr (assoc 360 rec))
                     (cdr (assoc 350 rec))))
        (if (and (dlt:ename-p en)
                 (setq recEd (entget en))
                 (= (cdr (assoc 0 recEd)) "DICTIONARY"))
          (setq out (append out (list en)))
        )
        (setq rec (dictnext nod))
      )
    )
  )
  out
)

(defun dlt:get-datalink-enames-all (/ out dictE dictEnames d en ed item items)
  (setq out '())

  ;; Step 1 - reuse the datalink_auto scanner if it is already loaded.
  (if (dlt:function-defined-p 'dl:get-datalink-items-all)
    (progn
      (setq items (vl-catch-all-apply 'dl:get-datalink-items-all '()))
      (if (and (not (vl-catch-all-error-p items)) items)
        (foreach item items
          (if (and (= (type item) 'LIST) (dlt:ename-p (car item)))
            (setq out (dlt:add-unique-ename (car item) out))
          )
        )
      )
    )
  )

  ;; Step 2 - ActiveX Dictionary enumeration, same source as Data Link Manager.
  (if (null out)
    (foreach item (dlt:get-datalink-enames-vla)
      (setq out (dlt:add-unique-ename item out))
    )
  )

  ;; Step 3 - standard ACAD_DATALINK dictionary.
  (if (null out)
    (progn
      (setq dictE (dlt:get-existing-datalink-dict-ename))
      (if (dlt:ename-p dictE)
        (progn
          (setq dlt:*dict-visited* nil)
          (foreach item (dlt:collect-datalinks-from-dict dictE)
            (setq out (dlt:add-unique-ename item out))
          )
        )
      )
    )
  )

  ;; Step 4 - fallback: scan all top-level NOD dictionaries.
  (if (null out)
    (progn
      (setq dictEnames (dlt:get-nod-dictionary-enames)
            dlt:*dict-visited* nil)
      (foreach d dictEnames
        (foreach item (dlt:collect-datalinks-from-dict d)
          (setq out (dlt:add-unique-ename item out))
        )
      )
    )
  )

  ;; Step 5 - last fallback: scan database entities.
  (if (null out)
    (progn
      (setq en (entnext))
      (while en
        (setq ed (entget en))
        (if (and ed (= (cdr (assoc 0 ed)) "DATALINK"))
          (setq out (dlt:add-unique-ename en out))
        )
        (setq en (entnext en))
      )
    )
  )
  out
)

(defun dlt:set-datalink-data-only (en / ed old91 new91 newEd)
  (if (and (dlt:ename-p en)
           (setq ed (entget en))
           (= (cdr (assoc 0 ed)) "DATALINK"))
    (progn
      (setq old91 (cdr (assoc 91 ed))
            new91 (dlt:normalize-datalink-update-flag91 old91)
            newEd ed)
      ;; DATALINK usually stores update options in both the entity body and its custom data map.
      ;; All 91 groups in this entity are update-option flags in the Data Links we create/read.
      (setq newEd (dlt:set-all-dxf newEd 91 new91))
      (setq newEd (dlt:set-first-dxf newEd 92 1))
      (if (entmod newEd)
        (progn
          (entupd en)
          T
        )
        nil
      )
    )
    nil
  )
)

(defun dlt:set-all-datalinks-data-only (/ links count en)
  (setq count 0
        links (dlt:get-datalink-enames-all))
  (foreach en links
    (if (dlt:set-datalink-data-only en)
      (setq count (1+ count))
    )
  )
  count
)

(defun dlt:get-table-style-dict (/ doc db dicts dict)
  (setq doc (dlt:get-active-document))
  (if doc
    (progn
      (setq db (vl-catch-all-apply 'vla-get-Database (list doc)))
      (if (vl-catch-all-error-p db)
        nil
        (progn
          (setq dicts (vl-catch-all-apply 'vla-get-Dictionaries (list db)))
          (if (vl-catch-all-error-p dicts)
            nil
            (progn
              (setq dict (dlt:safe-method dicts 'Item (list "ACAD_TABLESTYLE")))
              (if (= (type dict) 'VLA-OBJECT) dict nil)
            )
          )
        )
      )
    )
    nil
  )
)

(defun dlt:get-or-create-table-style (/ dict sty)
  (setq dict (dlt:get-table-style-dict))
  (if dict
    (progn
      (setq sty (dlt:safe-method dict 'Item (list dlt:*table-style-name*)))
      (if (/= (type sty) 'VLA-OBJECT)
        (setq sty (dlt:safe-method dict 'AddObject (list dlt:*table-style-name* "AcDbTableStyle")))
      )
      (if (= (type sty) 'VLA-OBJECT) sty nil)
    )
    nil
  )
)

(defun dlt:set-row-format-on-target (target rowType / ok)
  (setq ok nil)
  (if target
    (progn
      (if (dlt:try-method target 'SetTextHeight (list rowType dlt:*table-text-height*))
        (setq ok T)
      )
      (if (dlt:try-method target 'SetAlignment (list rowType dlt:*table-alignment*))
        (setq ok T)
      )
      (if (dlt:try-method target 'SetTextStyle (list rowType dlt:*table-text-style*))
        (setq ok T)
      )
    )
  )
  ok
)

(defun dlt:set-grid-format-on-target (target clr / rowType ok)
  (setq ok nil)
  (foreach rowType '(1 2 4 7)
    (if clr
      (progn
        (if (dlt:try-method target 'SetGridColor (list rowType dlt:*table-grid-types-all* clr))
          (setq ok T)
        )
        ;; Some AutoCAD table-style builds expose the argument order as grid type, row type.
        (if (dlt:try-method target 'SetGridColor (list dlt:*table-grid-types-all* rowType clr))
          (setq ok T)
        )
      )
    )
    (if (dlt:try-method target 'SetGridLineWeight (list rowType dlt:*table-grid-types-all* dlt:*table-lineweight*))
      (setq ok T)
    )
    (if (dlt:try-method target 'SetGridLineWeight (list dlt:*table-grid-types-all* rowType dlt:*table-lineweight*))
      (setq ok T)
    )
  )
  ok
)

(defun dlt:ensure-table-style-format (tableObj / sty clr rowType)
  (setq sty (dlt:get-or-create-table-style))
  (if sty
    (progn
      (dlt:try-put sty 'Name dlt:*table-style-name*)
      (dlt:try-put sty 'Description "Data Link table style managed by DLTABLEUI")
      (setq clr (dlt:safe-call 'vla-get-TrueColor (list tableObj)))
      (if (and clr (not (vl-catch-all-error-p clr)) (= (type clr) 'VLA-OBJECT))
        (dlt:safe-put clr 'ColorIndex dlt:*table-color-index*)
        (setq clr nil)
      )
      (foreach rowType '(1 2 4 7)
        (dlt:set-row-format-on-target sty rowType)
      )
      (dlt:set-grid-format-on-target sty clr)
      (if (and clr (= (type clr) 'VLA-OBJECT))
        (vl-catch-all-apply 'vlax-release-object (list clr))
      )
      dlt:*table-style-name*
    )
    nil
  )
)

(defun dlt:clear-table-style-overrides (obj)
  ;; Safe attempts only. Some AutoCAD ActiveX builds expose this, others do not.
  (or
    (dlt:try-method obj 'ClearTableStyleOverrides '())
    (dlt:try-method obj 'ClearTableStyleOverrides (list 0))
    (dlt:try-method obj 'ClearTableStyleOverrides (list 7))
    (dlt:try-method obj 'RemoveAllOverrides '())
  )
)

(defun dlt:apply-cell-text-format (obj rows cols / row col)
  (if (and obj (vlax-method-applicable-p obj 'SetCellTextHeight))
    (progn
      (setq row 0)
      (while (< row rows)
        (setq col 0)
        (while (< col cols)
          (dlt:try-method obj 'SetCellTextHeight (list row col dlt:*table-text-height*))
          (setq col (1+ col))
        )
        (setq row (1+ row))
      )
    )
  )
  (if (and obj (vlax-method-applicable-p obj 'SetCellAlignment))
    (progn
      (setq row 0)
      (while (< row rows)
        (setq col 0)
        (while (< col cols)
          (dlt:try-method obj 'SetCellAlignment (list row col dlt:*table-alignment*))
          (setq col (1+ col))
        )
        (setq row (1+ row))
      )
    )
  )
  T
)

(defun dlt:apply-cell-grid-format (obj clr rows cols / row col)
  (if (and obj clr (= (type clr) 'VLA-OBJECT) (vlax-method-applicable-p obj 'SetCellGridColor))
    (progn
      (setq row 0)
      (while (< row rows)
        (setq col 0)
        (while (< col cols)
          (dlt:try-method obj 'SetCellGridColor (list row col 15 clr))
          (dlt:try-method obj 'SetCellGridColor (list row col dlt:*table-grid-types-all* clr))
          (setq col (1+ col))
        )
        (setq row (1+ row))
      )
    )
  )
  (if (and obj (vlax-method-applicable-p obj 'SetCellGridLineWeight))
    (progn
      (setq row 0)
      (while (< row rows)
        (setq col 0)
        (while (< col cols)
          (dlt:try-method obj 'SetCellGridLineWeight (list row col 15 dlt:*table-lineweight*))
          (dlt:try-method obj 'SetCellGridLineWeight (list row col dlt:*table-grid-types-all* dlt:*table-lineweight*))
          (setq col (1+ col))
        )
        (setq row (1+ row))
      )
    )
  )
  T
)

(defun dlt:valid-table-ename-p (e)
  (if (and e (entget e) (dlt:table-ename-p e))
    T
    nil
  )
)

(defun dlt:send-pending-format-command (/ doc cmd res)
  (setq doc (dlt:get-active-document))
  (if doc
    (progn
      (setq cmd
        (strcat
          "_.DELAY " (itoa dlt:*format-delay-ms*) "\n"
          "DLTAPPLYPENDING\n"
        )
      )
      (setq res (vl-catch-all-apply 'vla-SendCommand (list doc cmd)))
      (if (vl-catch-all-error-p res) nil T)
    )
    nil
  )
)

(defun dlt:queue-pending-format (mode tableEname source)
  (setq dlt:*pending-format-mode* mode
        dlt:*pending-table-ename* tableEname
        dlt:*pending-format-source* source
        dlt:*pending-format-remaining* dlt:*format-repeat-count*)
  (dlt:send-pending-format-command)
)

(defun dlt:apply-grid-format (obj / clr rows cols)
  (dlt:safe-call 'vla-put-Color (list obj dlt:*table-color-index*))
  (dlt:safe-call 'vla-put-Linetype (list obj dlt:*table-linetype*))
  (dlt:safe-call 'vla-put-Lineweight (list obj dlt:*table-lineweight*))

  (setq rows (vla-get-Rows obj)
        cols (vla-get-Columns obj))

  (setq clr (dlt:safe-call 'vla-get-TrueColor (list obj)))
  (if (and clr (not (vl-catch-all-error-p clr)) (= (type clr) 'VLA-OBJECT))
    (dlt:safe-put clr 'ColorIndex dlt:*table-color-index*)
    (setq clr nil)
  )

  (dlt:set-grid-format-on-target obj clr)
  (dlt:apply-cell-grid-format obj clr rows cols)

  (if (and clr (not (vl-catch-all-error-p clr)) (= (type clr) 'VLA-OBJECT))
    (vl-catch-all-apply 'vlax-release-object (list clr))
  )
  T
)

(defun dlt:format-table-object (obj / rows cols r c styleName)
  (setq rows (vla-get-Rows obj))
  (setq cols (vla-get-Columns obj))

  (dlt:clear-table-style-overrides obj)
  (setq styleName (dlt:ensure-table-style-format obj))
  (if styleName
    (dlt:try-put obj 'StyleName styleName)
  )

  ;; 每列列高
  (setq r 0)
  (while (< r rows)
    (dlt:safe-method obj 'SetRowHeight (list r dlt:*table-row-height*))
    (setq r (1+ r))
  )

  ;; 每欄欄寬
  (setq c 0)
  (while (< c cols)
    (dlt:safe-method obj 'SetColumnWidth (list c dlt:*table-col-width*))
    (setq c (1+ c))
  )

  ;; 文字高度/對齊/文字型式：1=Title, 2=Header, 4=Data, 7=全部
  (foreach r '(1 2 4 7)
    (dlt:set-row-format-on-target obj r)
  )
  (dlt:apply-cell-text-format obj rows cols)

  (dlt:apply-grid-format obj)
  (dlt:safe-method obj 'GenerateLayout '())
  (dlt:safe-method obj 'Update '())
)

(defun dlt:format-table-ename (e / obj)
  (if (dlt:table-ename-p e)
    (progn
      (dlt:set-all-datalinks-data-only)
      (setq obj (vlax-ename->vla-object e))
      (dlt:format-table-object obj)
      (vl-catch-all-apply 'vlax-release-object (list obj))
      (princ "\n表格格式已自動調整完成。")
    )
    (princ "\n沒有找到可格式化的表格。")
  )
  (princ)
)

(defun dlt:format-all-tables-current-tab (/ ss i count e)
  (setq ss (ssget "_X" (list '(0 . "ACAD_TABLE") (cons 410 (getvar "CTAB"))))
        count 0)
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (if (dlt:table-ename-p e)
          (progn
            (dlt:format-table-ename e)
            (setq count (1+ count))
          )
        )
        (setq i (1+ i))
      )
    )
  )
  count
)

(defun dlt:on-datalinkupdate-ended (reactor params / cmd)
  (setq cmd (strcase (car params)))
  (if (wcmatch cmd "*DATALINKUPDATE*")
    (progn
      (if (dlt:queue-pending-format 'all nil "DATALINKUPDATE")
        (princ "\nDATALINKUPDATE 完成，已排程延後重套表格格式。")
        (princ "\nDATALINKUPDATE 完成。若表格被資料連結覆蓋縮回，請執行 TBFIXALL。")
      )
    )
  )
  (princ)
)

(defun dlt:on-datalinkupdate-started (reactor params / cmd count)
  (setq cmd (strcase (car params)))
  (if (wcmatch cmd "*DATALINKUPDATE*")
    (progn
      (setq count (dlt:set-all-datalinks-data-only))
      (princ (strcat "\nDATALINKUPDATE 開始前，已將 Data Link 改為只更新資料。數量: " (itoa count)))
    )
  )
  (princ)
)

(defun dlt:ensure-update-reactor ()
  (if (or (null dlt:*update-reactor*) (not (vlr-added-p dlt:*update-reactor*)))
    (setq dlt:*update-reactor*
      (vlr-command-reactor
        nil
        '((:vlr-commandWillStart . dlt:on-datalinkupdate-started)
          (:vlr-commandEnded . dlt:on-datalinkupdate-ended))
      )
    )
  )
  T
)

(defun dlt:on-table-command-ended (reactor params / cmd newtab)
  (setq cmd (strcase (car params)))
  (if (wcmatch cmd "*TABLE*")
    (progn
      (setq newtab (dlt:find-new-table-after dlt:*table-before-ent*))
      (dlt:remove-reactor)
      (setq dlt:*table-before-ent* nil)
      (if newtab
        (if (dlt:queue-pending-format 'one newtab "TABLE")
          (princ "\nTABLE 已結束，已排程延後格式化新表格。")
          (dlt:format-table-ename newtab)
        )
        (princ "\nTABLE 已結束，但沒有偵測到新表格。可改用 TBFIX 手動選表格格式化。")
      )
    )
  )
  (princ)
)

(defun c:DLTAPPLYPENDING (/ mode tableEname source remaining count)
  (vl-load-com)
  (setq mode dlt:*pending-format-mode*
        tableEname dlt:*pending-table-ename*
        source dlt:*pending-format-source*
        remaining dlt:*pending-format-remaining*)
  (cond
    ((or (null mode) (<= remaining 0))
     (princ "\n沒有待處理的表格格式化。")
    )
    (T
     (cond
       ((and (= mode 'one) (dlt:valid-table-ename-p tableEname))
        (dlt:format-table-ename tableEname)
        (princ (strcat "\n延後格式化完成。剩餘補格式次數: " (itoa (1- remaining))))
       )
       ((or (= mode 'all) (= source "DATALINKUPDATE") (= source "TABLE"))
        (setq count (dlt:format-all-tables-current-tab))
        (princ (strcat "\n延後重套表格格式完成。數量: " (itoa count) "，剩餘補格式次數: " (itoa (1- remaining))))
       )
       (T
        (princ "\n沒有可格式化的表格。")
       )
     )
     (setq dlt:*pending-format-remaining* (1- remaining))
     (if (> dlt:*pending-format-remaining* 0)
       (if (not (dlt:send-pending-format-command))
         (setq dlt:*pending-format-remaining* 0)
       )
       (setq dlt:*pending-format-mode* nil
             dlt:*pending-table-ename* nil
             dlt:*pending-format-source* "")
     )
    )
  )
  (princ)
)

(defun dlt:on-table-command-cancelled (reactor params / cmd)
  (setq cmd (strcase (car params)))
  (if (wcmatch cmd "*TABLE*")
    (progn
      (dlt:remove-reactor)
      (setq dlt:*table-before-ent* nil)
      (princ "\nTABLE 已取消，未調整格式。")
    )
  )
  (princ)
)

(defun dlt:prepare-table-reactor ()
  (vl-load-com)
  (dlt:remove-reactor)
  (setq dlt:*table-before-ent* (entlast))
  (setq dlt:*table-reactor*
    (vlr-command-reactor
      nil
      '((:vlr-commandEnded . dlt:on-table-command-ended)
        (:vlr-commandCancelled . dlt:on-table-command-cancelled)
        (:vlr-commandFailed . dlt:on-table-command-cancelled))
    )
  )
)

(defun c:DLTABLEUI (/)
  (if (> (getvar "CMDACTIVE") 0)
    (princ "\n目前有指令正在執行，請先 Esc 結束後再執行 DLTABLEUI。")
    (progn
      (princ (strcat "\n已先將 Data Link 改為只更新資料。數量: " (itoa (dlt:set-all-datalinks-data-only))))
      (dlt:prepare-table-reactor)
      (princ "\n請在 TABLE 視窗中選 From a data link / 從資料連結，選 Data Link 後指定插入點。完成後會自動格式化。")
      (initdia)
      (command "_.TABLE")
    )
  )
  (princ)
)

(defun c:DLTABLEDIALOG ()
  (c:DLTABLEUI)
)

(defun c:DLTABLEAUTO ()
  (c:DLTABLEUI)
)

(defun c:AUTODATALINKTABLE ()
  (c:DLTABLEUI)
)

(defun c:DLLINKTABLE ()
  (c:DLTABLEUI)
)

(defun c:TBFIX (/ ent)
  (vl-load-com)
  (setq ent (car (entsel "\n選取要調整格式的表格: ")))
  (if ent
    (dlt:format-table-ename ent)
    (princ "\n未選取表格。")
  )
  (princ)
)

(defun c:TBFIXALL (/ count)
  (vl-load-com)
  (setq count (dlt:format-all-tables-current-tab))
  (princ (strcat "\n已重套目前頁籤表格格式。數量: " (itoa count)))
  (princ)
)

(defun c:DLDATAONLY (/ count)
  (vl-load-com)
  (setq count (dlt:set-all-datalinks-data-only))
  (princ (strcat "\n已將 Data Link 改為只更新資料、跳過 Excel 格式。數量: " (itoa count)))
  (princ)
)

(dlt:ensure-update-reactor)

(princ "\n已載入：DLTABLEUI/DLTABLEAUTO/DLTABLEDIALOG = 開啟 TABLE 視窗建 Data Link 表格後延後格式化；TBFIX/TBFIXALL = 重套表格格式；DLDATAONLY = Data Link 只更新資料。")
(princ)
