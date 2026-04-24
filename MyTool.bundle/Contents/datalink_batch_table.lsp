(vl-load-com)

;; ==========================================================
;; datalink_batch_table.lsp
;; Commands:
;;   DLBATCHAUTO   - create one linked AutoCAD table for every Data Link
;;   DLBATCHTABLE  - alias of DLBATCHAUTO
;;   DLBATCHTABLES - alias of DLBATCHAUTO
;;   DLLIST        - list usable Data Link objects in the current drawing
;;   DLBATCHLIST   - alias of DLLIST
;;   DLBATCHFIX    - manually format one selected table
;;
;; Note:
;;   DLBATCHAUTO uses AutoCAD Table.SetDataLink. If a specific AutoCAD build
;;   does not expose this method to ActiveX, use DLTABLEUI + TBFIX instead.
;; ==========================================================

;; Main table format values.
(setq dlbt:*table-row-height* 486.0)
(setq dlbt:*table-col-width* 2025.0)
(setq dlbt:*table-text-height* 200.0)
(setq dlbt:*table-text-style* "Standard")
(setq dlbt:*table-alignment* 5)          ;; acMiddleCenter
(setq dlbt:*table-color-index* 256)      ;; ByLayer
(setq dlbt:*table-linetype* "ByLayer")
(setq dlbt:*table-lineweight* -1)        ;; ByLayer
(setq dlbt:*table-grid-types-all* 63)
(setq dlbt:*table-cell-grid-types* 63)
(setq dlbt:*default-spacing* 3500.0)
(setq dlbt:*last-create-error* nil)

(defun dlbt:trim (s)
  (if s
    (vl-string-trim " \t\r\n" s)
    ""
  )
)

(defun dlbt:ename-p (x)
  (= (type x) 'ENAME)
)

(defun dlbt:release (obj)
  (if (and obj (= (type obj) 'VLA-OBJECT))
    (vl-catch-all-apply 'vlax-release-object (list obj))
  )
)

(defun dlbt:safe-call (fn args)
  (vl-catch-all-apply fn args)
)

(defun dlbt:safe-put (obj prop val)
  (vl-catch-all-apply 'vlax-put-property (list obj prop val))
)

(defun dlbt:catch-message (res fallback)
  (if (vl-catch-all-error-p res)
    (vl-catch-all-error-message res)
    fallback
  )
)

(defun dlbt:get-doc ()
  (vla-get-ActiveDocument (vlax-get-acad-object))
)

(defun dlbt:get-ms ()
  (vla-get-ModelSpace (dlbt:get-doc))
)

(defun dlbt:3dpt (p)
  (vlax-3d-point (list (car p) (cadr p) (if (caddr p) (caddr p) 0.0)))
)

(defun dlbt:set-dxf (ed code val / old)
  (if (setq old (assoc code ed))
    (subst (cons code val) old ed)
    (append ed (list (cons code val)))
  )
)

(defun dlbt:safe-vla-getname (dictObj childObj / r)
  (setq r (vl-catch-all-apply 'vla-GetName (list dictObj childObj)))
  (if (vl-catch-all-error-p r)
    ""
    (dlbt:trim r)
  )
)

(defun dlbt:safe-vla-objectname (obj / r)
  (setq r (vl-catch-all-apply 'vla-get-ObjectName (list obj)))
  (if (vl-catch-all-error-p r)
    ""
    (strcase (dlbt:trim r))
  )
)

(defun dlbt:safe-vla-handle (obj / r)
  (setq r (vl-catch-all-apply 'vla-get-Handle (list obj)))
  (if (vl-catch-all-error-p r)
    ""
    (dlbt:trim r)
  )
)

(defun dlbt:extract-datalink-name-from-detail (detail / p1 p2 out)
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
  (dlbt:trim out)
)

(defun dlbt:get-datalink-name-from-ename (en / ed out)
  (setq out "")
  (if (dlbt:ename-p en)
    (progn
      (setq ed (entget en))
      (if (and ed (= (cdr (assoc 0 ed)) "DATALINK"))
        (progn
          (setq out (dlbt:extract-datalink-name-from-detail (cdr (assoc 301 ed))))
          (if (= out "")
            (setq out (dlbt:trim (cdr (assoc 3 ed))))
          )
          (if (= out "")
            (setq out (dlbt:trim (cdr (assoc 300 ed))))
          )
        )
      )
    )
  )
  out
)

(defun dlbt:add-datalink-pair (name en pairs / clean target exists ed)
  (setq clean (dlbt:trim name))
  (if (and (= clean "") (dlbt:ename-p en))
    (setq clean (dlbt:get-datalink-name-from-ename en))
  )
  (if (and (/= clean "")
           (dlbt:ename-p en)
           (setq ed (entget en))
           (= (cdr (assoc 0 ed)) "DATALINK"))
    (progn
      (setq target (strcase clean)
            exists nil)
      (foreach p pairs
        (if (= (strcase (dlbt:trim (car p))) target)
          (setq exists T)
        )
      )
      (if exists
        pairs
        (append pairs (list (cons clean en)))
      )
    )
    pairs
  )
)

(setq dlbt:*dict-vla-visited* nil)

(defun dlbt:get-datalink-pairs-from-vla-dict-recursive (dictObj pairs / item key oname handle en)
  (if (and dictObj (= (type dictObj) 'VLA-OBJECT))
    (progn
      (setq handle (dlbt:safe-vla-handle dictObj))
      (if (and (/= handle "") (member handle dlbt:*dict-vla-visited*))
        pairs
        (progn
          (if (/= handle "")
            (setq dlbt:*dict-vla-visited* (cons handle dlbt:*dict-vla-visited*))
          )
          (vlax-for item dictObj
            (setq key   (dlbt:safe-vla-getname dictObj item)
                  oname (dlbt:safe-vla-objectname item))
            (cond
              ((wcmatch oname "*DICTIONARY*")
               (setq pairs (dlbt:get-datalink-pairs-from-vla-dict-recursive item pairs))
              )
              ((wcmatch oname "*DATALINK*")
               (setq en (vl-catch-all-apply 'vlax-vla-object->ename (list item)))
               (if (not (vl-catch-all-error-p en))
                 (setq pairs (dlbt:add-datalink-pair key en pairs))
               )
              )
            )
          )
        )
      )
    )
  )
  pairs
)

(defun dlbt:get-datalink-pairs-vla (/ acad doc db dicts dictObj pairs)
  (setq pairs '())
  (setq acad (vlax-get-acad-object)
        doc  (vla-get-ActiveDocument acad)
        db   (vla-get-Database doc)
        dicts (vla-get-Dictionaries db))
  (setq dictObj (vl-catch-all-apply 'vla-Item (list dicts "ACAD_DATALINK")))
  (if (not (vl-catch-all-error-p dictObj))
    (progn
      (setq dlbt:*dict-vla-visited* nil)
      (setq pairs (dlbt:get-datalink-pairs-from-vla-dict-recursive dictObj pairs))
      (dlbt:release dictObj)
    )
  )
  (dlbt:release dicts)
  pairs
)

(defun dlbt:get-existing-datalink-dict-ename (/ nod rec en ed)
  (setq nod (namedobjdict))
  (if (dlbt:ename-p nod)
    (progn
      (setq rec (dictsearch nod "ACAD_DATALINK"))
      (if rec
        (progn
          (setq en (cdr (assoc -1 rec))
                ed (if (dlbt:ename-p en) (entget en)))
          (if (or (not ed) (/= (cdr (assoc 0 ed)) "DICTIONARY"))
            (setq en (or (cdr (assoc 360 rec))
                         (cdr (assoc 350 rec))))
          )
          (if (dlbt:ename-p en) en nil)
        )
        nil
      )
    )
    nil
  )
)

(setq dlbt:*dict-visited* nil)

(defun dlbt:get-datalink-pairs-from-dict-recursive (dictEname pairs / rec key objEn objEd typ handle)
  (if (dlbt:ename-p dictEname)
    (progn
      (setq handle (cdr (assoc 5 (entget dictEname))))
      (if (and handle (member handle dlbt:*dict-visited*))
        pairs
        (progn
          (if handle
            (setq dlbt:*dict-visited* (cons handle dlbt:*dict-visited*))
          )
          (setq rec (dictnext dictEname T))
          (while rec
            (setq key   (dlbt:trim (cdr (assoc 3 rec)))
                  objEn (or (cdr (assoc 360 rec))
                            (cdr (assoc 350 rec)))
                  objEd (if (dlbt:ename-p objEn) (entget objEn))
                  typ   (if objEd (cdr (assoc 0 objEd))))
            (cond
              ((= typ "DATALINK")
               (setq pairs (dlbt:add-datalink-pair key objEn pairs))
              )
              ((= typ "DICTIONARY")
               (setq pairs (dlbt:get-datalink-pairs-from-dict-recursive objEn pairs))
              )
            )
            (setq rec (dictnext dictEname))
          )
        )
      )
    )
  )
  pairs
)

(defun dlbt:get-datalink-pairs-dict (/ dictEname pairs)
  (setq pairs '()
        dictEname (dlbt:get-existing-datalink-dict-ename))
  (if (dlbt:ename-p dictEname)
    (progn
      (setq dlbt:*dict-visited* nil)
      (setq pairs (dlbt:get-datalink-pairs-from-dict-recursive dictEname pairs))
    )
  )
  pairs
)

(defun dlbt:get-nod-dictionary-enames (/ nod rec en recEd out)
  (setq out '()
        nod (namedobjdict))
  (if (dlbt:ename-p nod)
    (progn
      (setq rec (dictnext nod T))
      (while rec
        (setq en (or (cdr (assoc 360 rec))
                     (cdr (assoc 350 rec))))
        (if (and (dlbt:ename-p en)
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

(defun dlbt:get-datalink-pairs-by-nod-scan (/ pairs dictEname)
  (setq pairs '()
        dlbt:*dict-visited* nil)
  (foreach dictEname (dlbt:get-nod-dictionary-enames)
    (setq pairs (dlbt:get-datalink-pairs-from-dict-recursive dictEname pairs))
  )
  pairs
)

(defun dlbt:get-datalink-pairs-by-object-scan (/ en ed name pairs)
  (setq pairs '()
        en (entnext))
  (while en
    (setq ed (entget en))
    (if (and ed (= (cdr (assoc 0 ed)) "DATALINK"))
      (progn
        (setq name (dlbt:get-datalink-name-from-ename en))
        (setq pairs (dlbt:add-datalink-pair name en pairs))
      )
    )
    (setq en (entnext en))
  )
  pairs
)

(defun dlbt:get-datalink-pairs (/ pairs)
  (setq pairs (dlbt:get-datalink-pairs-vla))
  (if (null pairs)
    (setq pairs (dlbt:get-datalink-pairs-dict))
  )
  (if (null pairs)
    (setq pairs (dlbt:get-datalink-pairs-by-nod-scan))
  )
  (if (null pairs)
    (setq pairs (dlbt:get-datalink-pairs-by-object-scan))
  )
  pairs
)

(defun dlbt:print-datalinks (/ pairs i)
  (setq pairs (dlbt:get-datalink-pairs))
  (if pairs
    (progn
      (princ "\n目前圖面 Data Link：")
      (setq i 1)
      (foreach p pairs
        (princ (strcat "\n  " (itoa i) ". " (car p)))
        (setq i (1+ i))
      )
    )
    (princ "\n目前圖面找不到可用的 Data Link。")
  )
  pairs
)

(defun dlbt:table-ename-p (e / ed objname)
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

(defun dlbt:find-new-table-after (before / e found last)
  (setq found nil)
  (setq e (if before (entnext before) (entnext)))
  (while e
    (if (dlbt:table-ename-p e)
      (setq found e)
    )
    (setq e (entnext e))
  )
  (if (null found)
    (progn
      (setq last (entlast))
      (if (dlbt:table-ename-p last)
        (setq found last)
      )
    )
  )
  found
)

(defun dlbt:set-entity-table-props (e / ed)
  (if e
    (progn
      (setq ed (entget e))
      (if ed
        (progn
          (setq ed
                (vl-remove-if
                  '(lambda (x) (or (= (car x) 420) (= (car x) 430) (= (car x) 440)))
                  ed))
          (setq ed (dlbt:set-dxf ed 62 dlbt:*table-color-index*))
          (setq ed (dlbt:set-dxf ed 6 dlbt:*table-linetype*))
          (setq ed (dlbt:set-dxf ed 370 dlbt:*table-lineweight*))
          (entmod ed)
          (entupd e)
        )
      )
    )
  )
  T
)

(defun dlbt:apply-grid-format (obj rows cols / clr r c)
  (dlbt:safe-call 'vla-put-Color (list obj dlbt:*table-color-index*))
  (dlbt:safe-call 'vla-put-Linetype (list obj dlbt:*table-linetype*))
  (dlbt:safe-call 'vla-put-Lineweight (list obj dlbt:*table-lineweight*))

  (setq clr (vl-catch-all-apply 'vla-get-TrueColor (list obj)))
  (if (not (vl-catch-all-error-p clr))
    (progn
      (dlbt:safe-put clr 'ColorIndex dlbt:*table-color-index*)
      (dlbt:safe-call 'vla-SetGridColor (list obj 1 dlbt:*table-grid-types-all* clr))
      (dlbt:safe-call 'vla-SetGridColor (list obj 2 dlbt:*table-grid-types-all* clr))
      (dlbt:safe-call 'vla-SetGridColor (list obj 4 dlbt:*table-grid-types-all* clr))
      (dlbt:safe-call 'vla-SetGridColor (list obj 7 dlbt:*table-grid-types-all* clr))
    )
  )

  (dlbt:safe-call 'vla-SetGridColor2 (list obj 1 dlbt:*table-grid-types-all* dlbt:*table-color-index*))
  (dlbt:safe-call 'vla-SetGridColor2 (list obj 2 dlbt:*table-grid-types-all* dlbt:*table-color-index*))
  (dlbt:safe-call 'vla-SetGridColor2 (list obj 4 dlbt:*table-grid-types-all* dlbt:*table-color-index*))
  (dlbt:safe-call 'vla-SetGridColor2 (list obj 7 dlbt:*table-grid-types-all* dlbt:*table-color-index*))

  (dlbt:safe-call 'vla-SetGridLineWeight (list obj 1 dlbt:*table-grid-types-all* dlbt:*table-lineweight*))
  (dlbt:safe-call 'vla-SetGridLineWeight (list obj 2 dlbt:*table-grid-types-all* dlbt:*table-lineweight*))
  (dlbt:safe-call 'vla-SetGridLineWeight (list obj 4 dlbt:*table-grid-types-all* dlbt:*table-lineweight*))
  (dlbt:safe-call 'vla-SetGridLineWeight (list obj 7 dlbt:*table-grid-types-all* dlbt:*table-lineweight*))

  (setq r 0)
  (while (< r rows)
    (setq c 0)
    (while (< c cols)
      (if (not (vl-catch-all-error-p clr))
        (dlbt:safe-call 'vla-SetCellGridColor (list obj r c dlbt:*table-cell-grid-types* clr))
      )
      (dlbt:safe-call 'vla-SetCellGridColor2 (list obj r c dlbt:*table-cell-grid-types* dlbt:*table-color-index*))
      (dlbt:safe-call 'vla-SetCellGridLineWeight (list obj r c dlbt:*table-cell-grid-types* dlbt:*table-lineweight*))
      (setq c (1+ c))
    )
    (setq r (1+ r))
  )

  (if (and clr (not (vl-catch-all-error-p clr)) (= (type clr) 'VLA-OBJECT))
    (dlbt:release clr)
  )
  T
)

(defun dlbt:format-table-object (obj / rows cols r c)
  (setq rows (vla-get-Rows obj))
  (setq cols (vla-get-Columns obj))

  (setq r 0)
  (while (< r rows)
    (dlbt:safe-call 'vla-SetRowHeight (list obj r dlbt:*table-row-height*))
    (setq r (1+ r))
  )

  (setq c 0)
  (while (< c cols)
    (dlbt:safe-call 'vla-SetColumnWidth (list obj c dlbt:*table-col-width*))
    (setq c (1+ c))
  )

  (dlbt:safe-call 'vla-SetTextHeight (list obj 1 dlbt:*table-text-height*))
  (dlbt:safe-call 'vla-SetTextHeight (list obj 2 dlbt:*table-text-height*))
  (dlbt:safe-call 'vla-SetTextHeight (list obj 4 dlbt:*table-text-height*))

  (dlbt:safe-call 'vla-SetAlignment (list obj 1 dlbt:*table-alignment*))
  (dlbt:safe-call 'vla-SetAlignment (list obj 2 dlbt:*table-alignment*))
  (dlbt:safe-call 'vla-SetAlignment (list obj 4 dlbt:*table-alignment*))

  (dlbt:safe-call 'vla-SetTextStyle (list obj 1 dlbt:*table-text-style*))
  (dlbt:safe-call 'vla-SetTextStyle (list obj 2 dlbt:*table-text-style*))
  (dlbt:safe-call 'vla-SetTextStyle (list obj 4 dlbt:*table-text-style*))

  (setq r 0)
  (while (< r rows)
    (setq c 0)
    (while (< c cols)
      (dlbt:safe-call 'vla-SetCellTextHeight (list obj r c dlbt:*table-text-height*))
      (dlbt:safe-call 'vla-SetCellAlignment (list obj r c dlbt:*table-alignment*))
      (dlbt:safe-call 'vla-SetCellTextStyle (list obj r c dlbt:*table-text-style*))
      (setq c (1+ c))
    )
    (setq r (1+ r))
  )

  (dlbt:apply-grid-format obj rows cols)
  (dlbt:safe-call 'vla-GenerateLayout (list obj))
  (dlbt:safe-call 'vla-Update (list obj))
  T
)

(defun dlbt:format-table-ename (e / obj ok)
  (setq ok nil)
  (if (dlbt:table-ename-p e)
    (progn
      (setq obj (vlax-ename->vla-object e))
      (dlbt:format-table-object obj)
      (dlbt:set-entity-table-props e)
      (dlbt:release obj)
      (setq ok T)
    )
  )
  ok
)

(defun dlbt:set-table-datalink (tbl row col dlid / res)
  ;; Do not call vla-SetDataLink directly. Some AutoCAD builds do not generate
  ;; that Visual LISP wrapper, which raises "bad function: VLA-SETDATALINK"
  ;; before normal error handling can recover.
  (setq res (vl-catch-all-apply 'vlax-invoke-method (list tbl 'SetDataLink row col dlid :vlax-true)))
  res
)

(defun dlbt:update-table-datalink (tbl / res)
  ;; SourceToData = 1, UpdateOption.None = 0.
  (setq res (vl-catch-all-apply 'vlax-invoke-method (list tbl 'UpdateDataLink 0 0 1 0)))
  (if (vl-catch-all-error-p res)
    (setq res (vl-catch-all-apply 'vlax-invoke-method (list tbl 'UpdateDataLink 1 0)))
  )
  (if (vl-catch-all-error-p res)
    (setq res (vl-catch-all-apply 'vla-Update (list tbl)))
  )
  res
)

(defun dlbt:finish-command-prompts (/ guard)
  (setq guard 0)
  (while (and (> (getvar "CMDACTIVE") 0) (< guard 12))
    (vl-catch-all-apply 'vl-cmdf (list ""))
    (setq guard (1+ guard))
  )
  (= (getvar "CMDACTIVE") 0)
)

(defun dlbt:create-linked-table-by-command-once (dlpair inspt opt / before res newtab obj)
  (setq before (entlast))
  (setq res (vl-catch-all-apply 'vl-cmdf (list "_.-TABLE" opt (car dlpair) inspt)))
  (dlbt:finish-command-prompts)
  (setq newtab (dlbt:find-new-table-after before))
  (if newtab
    (progn
      (dlbt:format-table-ename newtab)
      (setq obj (vlax-ename->vla-object newtab))
      obj
    )
    (progn
      (setq dlbt:*last-create-error* (dlbt:catch-message res "TABLE command did not create a table"))
      nil
    )
  )
)

(defun dlbt:create-linked-table-by-command (dlpair inspt / oldcmdecho oldcmddia tbl)
  (setq oldcmdecho (getvar "CMDECHO")
        oldcmddia  (getvar "CMDDIA"))
  (setvar "CMDECHO" 0)
  (setvar "CMDDIA" 0)

  (setq tbl (dlbt:create-linked-table-by-command-once dlpair inspt "_L"))
  (if (null tbl)
    (setq tbl (dlbt:create-linked-table-by-command-once dlpair inspt "L"))
  )

  (setvar "CMDECHO" oldcmdecho)
  (setvar "CMDDIA" oldcmddia)
  tbl
)

(defun dlbt:create-linked-table (dlpair inspt / ms tbl dlobj dlid setRes updRes ok)
  (setq dlbt:*last-create-error* nil)
  (setq ms (dlbt:get-ms))
  (setq tbl
        (vl-catch-all-apply
          'vla-AddTable
          (list ms (dlbt:3dpt inspt) 1 1 dlbt:*table-row-height* dlbt:*table-col-width*)))

  (if (vl-catch-all-error-p tbl)
    (progn
      (setq dlbt:*last-create-error* (dlbt:catch-message tbl "AddTable failed"))
      (dlbt:create-linked-table-by-command dlpair inspt)
    )
    (progn
      (if (not (vlax-method-applicable-p tbl 'SetDataLink))
        (progn
          (setq dlbt:*last-create-error* "Table.SetDataLink is not available through ActiveX")
          (dlbt:safe-call 'vla-Delete (list tbl))
          (dlbt:create-linked-table-by-command dlpair inspt)
        )
        (progn
          (setq dlobj (vl-catch-all-apply 'vlax-ename->vla-object (list (cdr dlpair))))
          (if (vl-catch-all-error-p dlobj)
            (progn
              (setq dlbt:*last-create-error* (dlbt:catch-message dlobj "Data Link object failed"))
              (dlbt:safe-call 'vla-Delete (list tbl))
              (dlbt:create-linked-table-by-command dlpair inspt)
            )
            (progn
              (setq dlid (vl-catch-all-apply 'vla-get-ObjectID (list dlobj)))
              (if (vl-catch-all-error-p dlid)
                (progn
                  (setq dlbt:*last-create-error* (dlbt:catch-message dlid "ObjectID failed"))
                  (dlbt:safe-call 'vla-Delete (list tbl))
                  (dlbt:release dlobj)
                  (dlbt:create-linked-table-by-command dlpair inspt)
                )
                (progn
                  (setq setRes (dlbt:set-table-datalink tbl 0 0 dlid)
                        ok (not (vl-catch-all-error-p setRes)))
                  (if ok
                    (progn
                      (setq updRes (dlbt:update-table-datalink tbl))
                      (dlbt:safe-call 'vla-GenerateLayout (list tbl))
                      (dlbt:safe-call 'vla-Update (list tbl))
                      (dlbt:format-table-object tbl)
                      (dlbt:release dlobj)
                      tbl
                    )
                    (progn
                      (setq dlbt:*last-create-error* (dlbt:catch-message setRes "SetDataLink failed"))
                      (dlbt:safe-call 'vla-Delete (list tbl))
                      (dlbt:release dlobj)
                      (dlbt:create-linked-table-by-command dlpair inspt)
                    )
                  )
                )
              )
            )
          )
        )
      )
    )
  )
)

(defun c:DLLIST (/)
  (vl-load-com)
  (dlbt:print-datalinks)
  (princ)
)

(defun c:DLBATCHLIST ()
  (c:DLLIST)
)

(defun c:DLBATCHFIX (/ ent)
  (vl-load-com)
  (setq ent (car (entsel "\n選取要調整格式的表格: ")))
  (if ent
    (if (dlbt:format-table-ename ent)
      (princ "\n表格格式已調整完成。")
      (princ "\n選到的物件不是表格。")
    )
    (princ "\n未選取表格。")
  )
  (princ)
)

(defun c:DLBATCHAUTO (/ pairs base sp i p pt tbl made failed ss en item msg)
  (vl-load-com)
  (setq pairs (dlbt:print-datalinks))
  (if pairs
    (progn
      (setq base (getpoint "\n指定第一張表格插入點: "))
      (if base
        (progn
          (setq sp (getdist (strcat "\n指定表格垂直間距 <" (rtos dlbt:*default-spacing* 2 0) ">: ")))
          (if (or (null sp) (<= sp 0.0))
            (setq sp dlbt:*default-spacing*)
          )

          (setq made 0
                failed nil
                ss (ssadd)
                i 0)

          (foreach p pairs
            (setq pt (list (car base) (- (cadr base) (* i sp)) (if (caddr base) (caddr base) 0.0)))
            (princ (strcat "\n建立 Data Link 表格：" (car p)))
            (setq tbl (dlbt:create-linked-table p pt))
            (if tbl
              (progn
                (setq made (1+ made))
                (setq en (vlax-vla-object->ename tbl))
                (if en
                  (ssadd en ss)
                )
              )
              (setq failed (cons (cons (car p) dlbt:*last-create-error*) failed))
            )
            (setq i (1+ i))
          )

          (if (> (sslength ss) 0)
            (sssetfirst nil ss)
          )

          (princ (strcat "\n完成。成功建立 " (itoa made) " 張 Data Link 表格。"))
          (if failed
            (progn
              (princ "\n以下 Data Link 建立失敗，通常是此 AutoCAD 版本未透過 ActiveX 暴露 Table.SetDataLink：")
              (foreach item (reverse failed)
                (setq msg (dlbt:trim (cdr item)))
                (princ (strcat "\n  - " (car item)))
                (if (/= msg "")
                  (princ (strcat "：" msg))
                )
              )
              (princ "\n失敗者請改用 DLTABLEUI 互動建立後，再用 TBFIX 或 DLBATCHFIX 格式化。")
            )
          )
        )
        (princ "\n未指定插入點。")
      )
    )
  )
  (princ)
)

(defun c:DLBATCHTABLE ()
  (c:DLBATCHAUTO)
)

(defun c:DLBATCHTABLES ()
  (c:DLBATCHAUTO)
)

(princ "\n已載入 datalink_batch_table.lsp")
(princ "\n指令：DLLIST/DLBATCHLIST 列出 Data Link；DLBATCHAUTO/DLBATCHTABLE 一次建立全部表格；DLBATCHFIX 手動格式化表格。")
(princ)
