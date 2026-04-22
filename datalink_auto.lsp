(vl-load-com)

; 安全預設：
; 1) 禁止由 AutoCAD 回寫 Excel 來源檔
; 2) 保留 Excel 格式（含合併儲存格、列高、欄寬）
; UpdateOption 參考：
; AllowSourceUpdate=1048576, SkipFormat=131072, UpdateRowHeight=262144, UpdateColumnWidth=524288,
; OverwriteContentModifiedAfterUpdate=4194304, OverwriteFormatModifiedAfterUpdate=8388608
(setq dl:*safe-update-flag91* 13369344)
(setq dl:*safe-update-flag92* 1)

(defun dl:bit-clear (flags bitMask)
  (if (/= (logand flags bitMask) 0)
    (- flags bitMask)
    flags
  )
)

(defun dl:normalize-update-flag91 (flag91 / out)
  (setq out (if flag91 flag91 0))
  ; 禁止寫回來源檔
  (setq out (dl:bit-clear out 1048576))
  ; 不可跳過格式，否則合併儲存格/格式會丟失
  (setq out (dl:bit-clear out 131072))
  ; 保留並更新 Excel 格式相關資訊
  (setq out (logior out 262144))
  (setq out (logior out 524288))
  (setq out (logior out 4194304))
  (setq out (logior out 8388608))
  out
)

(defun dl:trim (s)
  (if s
    (vl-string-trim " \t\r\n" s)
    ""
  )
)

(defun dl:norm-path (p / out)
  (setq out (dl:trim p))
  (while (vl-string-search "/" out)
    (setq out (vl-string-subst "\\" "/" out))
  )
  (while (and (> (strlen out) 1)
              (= (substr out (strlen out) 1) "\\"))
    (setq out (substr out 1 (1- (strlen out))))
  )
  out
)

(defun dl:split-path (path / parts pos start part)
  (setq parts '()
        start 1)
  (while (setq pos (vl-string-search "\\" path (1- start)))
    (setq part (substr path start (- (+ pos 1) start)))
    (if (/= part "")
      (setq parts (append parts (list part)))
    )
    (setq start (+ pos 2))
  )
  (setq part (substr path start))
  (if (/= part "")
    (setq parts (append parts (list part)))
  )
  parts
)

(defun dl:join-path (parts / out)
  (setq out "")
  (foreach p parts
    (if (= out "")
      (setq out p)
      (setq out (strcat out "\\" p))
    )
  )
  out
)

(defun dl:make-up-list (n / out)
  (setq out '())
  (repeat n
    (setq out (append out (list "..")))
  )
  out
)

(defun dl:drop (lst n)
  (while (and lst (> n 0))
    (setq lst (cdr lst)
          n   (1- n))
  )
  lst
)

(defun dl:relative-path-from-dwg (absPath / base target baseParts targetParts i common up parts rel)
  (setq base   (dl:norm-path (getvar "dwgprefix"))
        target (dl:norm-path absPath))
  (setq baseParts   (dl:split-path base)
        targetParts (dl:split-path target))

  (if (or (null baseParts)
          (null targetParts)
          (/= (strcase (car baseParts)) (strcase (car targetParts))))
    target
    (progn
      (setq i 0
            common (min (length baseParts) (length targetParts)))
      (while (and (< i common)
                  (= (strcase (nth i baseParts))
                     (strcase (nth i targetParts))))
        (setq i (1+ i))
      )
      (setq up    (dl:make-up-list (- (length baseParts) i))
            parts (append up (dl:drop targetParts i))
            rel   (dl:join-path parts))
      (if (= rel "")
        "."
        (if (= (substr rel 1 2) "..")
          rel
          (strcat ".\\" rel)
        )
      )
    )
  )
)

(defun dl:has-non-ascii (s / i ch found)
  (setq i 1
        found nil)
  (while (and s (not found) (<= i (strlen s)))
    (setq ch (ascii (substr s i 1)))
    (if (or (< ch 32) (> ch 126))
      (setq found T)
    )
    (setq i (1+ i))
  )
  found
)

(defun dl:get-short-path (absPath / fso fileObj shortPath)
  (setq shortPath nil)
  (setq fso (vl-catch-all-apply 'vlax-create-object (list "Scripting.FileSystemObject")))
  (if (not (vl-catch-all-error-p fso))
    (progn
      (if (= :vlax-true (vlax-invoke-method fso 'FileExists absPath))
        (progn
          (setq fileObj (vl-catch-all-apply 'vlax-invoke-method (list fso 'GetFile absPath)))
          (if (not (vl-catch-all-error-p fileObj))
            (setq shortPath (vlax-get-property fileObj 'ShortPath))
          )
        )
      )
      (dl:release fileObj)
      (dl:release fso)
    )
  )
  shortPath
)

(defun dl:pick-link-path (absPath / relPath shortAbs shortRel)
  (setq relPath (dl:relative-path-from-dwg absPath))
  (if (dl:has-non-ascii relPath)
    (progn
      (setq shortAbs (dl:get-short-path absPath))
      (if (and shortAbs (/= (dl:trim shortAbs) ""))
        (progn
          (setq shortRel (dl:relative-path-from-dwg shortAbs))
          (if (and shortRel (/= (dl:trim shortRel) ""))
            (setq relPath shortRel)
          )
        )
      )
    )
  )
  relPath
)

(defun dl:release (obj)
  (if (and obj (= (type obj) 'VLA-OBJECT))
    (vl-catch-all-apply 'vlax-release-object (list obj))
  )
)

(defun dl:ename-p (x)
  (= (type x) 'ENAME)
)

(defun dl:sanitize-name (s / out)
  (setq out (dl:trim s))
  (foreach ch '("\\" "/" ":" ";" "," "=" "?" "\"" "<" ">" "|" "*" "!" "\t" "\r" "\n")
    (setq out (vl-string-subst "_" ch out))
  )
  (if (= out "")
    "DATALINK"
    out
  )
)

(defun dl:get-text-from-entity (ename / obj oname txt)
  (setq txt "")
  (if ename
    (progn
      (setq obj (vlax-ename->vla-object ename))
      (if obj
        (progn
          (setq oname (vla-get-objectname obj))
          (if (wcmatch oname
                       "AcDbText,AcDbMText,AcDbAttributeDefinition,AcDbAttribute")
            (setq txt (vla-get-TextString obj))
          )
        )
      )
    )
  )
  (dl:trim txt)
)

(defun dl:get-panel-name (/ ent ename obj txt)
  (setq txt nil)
  (if (setq ent (entsel "\n請選擇盤名文字 (Enter改手動輸入): "))
    (progn
      (setq ename (car ent)
            obj   (vlax-ename->vla-object ename))
      (if obj
        (setq txt (dl:get-text-from-entity ename))
        (prompt "\n所選物件不是文字，改用手動輸入。")
      )
    )
  )
  (if (= (dl:trim txt) "")
    (setq txt (getstring T "\n輸入盤名: "))
  )
  (dl:trim txt)
)

(defun dl:get-sheet-names (filePath / xl books wb sheets sh names openRes)
  (setq names nil)
  (setq xl (vl-catch-all-apply 'vlax-get-or-create-object (list "Excel.Application")))
  (if (vl-catch-all-error-p xl)
    nil
    (progn
      (vlax-put-property xl 'Visible :vlax-false)
      (vlax-put-property xl 'DisplayAlerts :vlax-false)
      (setq books (vlax-get-property xl 'Workbooks))
      ; 以唯讀方式開啟來源檔，僅讀取分頁名稱，不對來源檔做任何寫入
      (setq openRes (vl-catch-all-apply 'vlax-invoke-method (list books 'Open filePath 0 :vlax-true)))
      (if (vl-catch-all-error-p openRes)
        (setq names nil)
        (progn
          (setq wb openRes
                sheets (vlax-get-property wb 'Worksheets))
          (vlax-for sh sheets
            (setq names (cons (vlax-get-property sh 'Name) names))
          )
          (vl-catch-all-apply 'vlax-invoke-method (list wb 'Close :vlax-false))
        )
      )
      (vl-catch-all-apply 'vlax-invoke-method (list xl 'Quit))
      (dl:release sh)
      (dl:release sheets)
      (dl:release wb)
      (dl:release books)
      (dl:release xl)
      (reverse names)
    )
  )
)

(defun dl:find-sheet-by-panel (panel sheetNames / pU exact)
  (setq pU (strcase (dl:trim panel)))
  (setq exact
    (vl-some
      '(lambda (s)
         (if (= (strcase (dl:trim s)) pU)
           s
         )
       )
      sheetNames
    )
  )
  (if exact
    exact
    (vl-some
      '(lambda (s)
         (if (or (wcmatch (strcase s) (strcat "*" pU "*"))
                 (wcmatch pU (strcat "*" (strcase s) "*")))
           s
         )
       )
      sheetNames
    )
  )
)

(defun dl:print-sheet-list (sheetNames / idx)
  (setq idx 1)
  (foreach s sheetNames
    (prompt (strcat "\n  " (itoa idx) ". " s))
    (setq idx (1+ idx))
  )
)

(defun dl:sheet-exists-ci (name sheetNames / target)
  (setq target (strcase (dl:trim name)))
  (vl-some
    '(lambda (s)
       (if (= (strcase (dl:trim s)) target)
         s
       )
     )
    sheetNames
  )
)

(defun dl:list-has-ci (item lst / target)
  (setq target (strcase (dl:trim item)))
  (if (vl-some '(lambda (s) (= (strcase (dl:trim s)) target)) lst)
    T
    nil
  )
)

(defun dl:add-unique-ci (item lst)
  (if (dl:list-has-ci item lst)
    lst
    (append lst (list item))
  )
)

(defun dl:get-panel-names-from-ss (ss / i en txt names)
  (setq i 0
        names '())
  (if ss
    (repeat (sslength ss)
      (setq en  (ssname ss i)
            txt (dl:get-text-from-entity en))
      (if (/= txt "")
        (setq names (dl:add-unique-ci txt names))
      )
      (setq i (1+ i))
    )
  )
  names
)

(defun dl:get-datalink-dict-object (/ acad doc db dicts dictObj nod rec en)
  (setq acad  (vlax-get-acad-object)
        doc   (vla-get-ActiveDocument acad)
        db    (vla-get-Database doc)
        dicts (vla-get-Dictionaries db))

  ; 先走資料庫 Dictionaries 集合，和 Data Link Manager 對齊
  (setq dictObj (vl-catch-all-apply 'vla-Item (list dicts "ACAD_DATALINK")))
  (if (vl-catch-all-error-p dictObj)
    (progn
      ; 回退：嘗試從 Named Object Dictionary 取
      (setq nod (namedobjdict)
            rec (if (dl:ename-p nod) (dictsearch nod "ACAD_DATALINK")))
      (if rec
        (progn
          (setq en (or (cdr (assoc 360 rec))
                       (cdr (assoc 350 rec))))
          (if (dl:ename-p en)
            (setq dictObj (vlax-ename->vla-object en))
            (setq dictObj nil)
          )
        )
        ; 都找不到才建立新的
        (setq dictObj (vla-Add dicts "ACAD_DATALINK"))
      )
    )
  )

  (if (vl-catch-all-error-p dictObj)
    nil
    dictObj
  )
)

(defun dl:get-datalink-dict-ename (/ dictObj en)
  (setq dictObj (dl:get-datalink-dict-object))
  (if dictObj
    (progn
      (setq en (vlax-vla-object->ename dictObj))
      (dl:release dictObj)
      (if (dl:ename-p en) en nil)
    )
    nil
  )
)

(defun dl:get-existing-datalink-dict-ename (/ nod rec)
  (setq nod (namedobjdict)
        rec (if (dl:ename-p nod) (dictsearch nod "ACAD_DATALINK")))
  (if rec
    (progn
      (setq rec (or (cdr (assoc 360 rec))
                    (cdr (assoc 350 rec))))
      (if (dl:ename-p rec) rec nil)
    )
    nil
  )
)

(defun dl:extract-datalink-name-from-detail (detail / p1 p2 out)
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
  (dl:trim out)
)

(defun dl:get-datalink-name-from-object (en / ed out)
  (setq out "")
  (if (dl:ename-p en)
    (progn
      (setq ed (entget en))
      (if (and ed (= (cdr (assoc 0 ed)) "DATALINK"))
        (progn
          (setq out (dl:extract-datalink-name-from-detail (cdr (assoc 301 ed))))
          (if (= out "")
            (setq out (dl:trim (cdr (assoc 3 ed))))
          )
          (if (= out "")
            (setq out (dl:trim (cdr (assoc 300 ed))))
          )
        )
      )
    )
  )
  out
)

(defun dl:get-datalink-names-from-dict (dictEname / rec key objEn objEd out)
  (setq out '())
  (if dictEname
    (progn
      (setq rec (dictnext dictEname T))
      (while rec
        (setq key   (dl:trim (cdr (assoc 3 rec)))
              objEn (or (cdr (assoc 360 rec))
                        (cdr (assoc 350 rec)))
              objEd (if (dl:ename-p objEn) (entget objEn)))
        (if (and objEd (= (cdr (assoc 0 objEd)) "DATALINK"))
          (progn
            (if (= key "")
              (setq key (dl:get-datalink-name-from-object objEn))
            )
            (if (/= key "")
              (setq out (dl:add-unique-ci key out))
            )
          )
        )
        (setq rec (dictnext dictEname))
      )
    )
  )
  out
)

(setq dl:*dict-visited* nil)
(setq dl:*dict-vla-visited* nil)

(defun dl:safe-vla-getname (dictObj childObj / r)
  (setq r (vl-catch-all-apply 'vla-GetName (list dictObj childObj)))
  (if (vl-catch-all-error-p r)
    ""
    (dl:trim r)
  )
)

(defun dl:safe-vla-objectname (obj / r)
  (setq r (vl-catch-all-apply 'vla-get-ObjectName (list obj)))
  (if (vl-catch-all-error-p r)
    ""
    (strcase (dl:trim r))
  )
)

(defun dl:safe-vla-handle (obj / r)
  (setq r (vl-catch-all-apply 'vla-get-Handle (list obj)))
  (if (vl-catch-all-error-p r)
    ""
    (dl:trim r)
  )
)

(defun dl:get-datalink-names-from-vla-dict-recursive (dictObj / out item key oname handle)
  (setq out '())
  (if (and dictObj (= (type dictObj) 'VLA-OBJECT))
    (progn
      (setq handle (dl:safe-vla-handle dictObj))
      (if (and (/= handle "") (member handle dl:*dict-vla-visited*))
        out
        (progn
          (if (/= handle "")
            (setq dl:*dict-vla-visited* (cons handle dl:*dict-vla-visited*))
          )
          (vlax-for item dictObj
            (setq key   (dl:safe-vla-getname dictObj item)
                  oname (dl:safe-vla-objectname item))
            (cond
              ((wcmatch oname "*DICTIONARY*")
               (foreach n (dl:get-datalink-names-from-vla-dict-recursive item)
                 (setq out (dl:add-unique-ci n out))
               )
              )
              (T
               ; 在 DataLink 字典樹中，非字典節點的 key 即可視為 Data Link 名稱
               (if (/= key "")
                 (setq out (dl:add-unique-ci key out))
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

(defun dl:get-datalink-names-vla (/ acad doc db dicts dictObj out)
  (setq out '())
  (setq acad (vlax-get-acad-object)
        doc  (vla-get-ActiveDocument acad)
        db   (vla-get-Database doc)
        dicts (vla-get-Dictionaries db))
  (setq dictObj (vl-catch-all-apply 'vla-Item (list dicts "ACAD_DATALINK")))
  (if (not (vl-catch-all-error-p dictObj))
    (progn
      (setq dl:*dict-vla-visited* nil)
      (setq out (dl:get-datalink-names-from-vla-dict-recursive dictObj))
      (dl:release dictObj)
    )
  )
  (dl:release dicts)
  out
)

(defun dl:get-datalink-names-from-dict-recursive (dictEname / rec key objEn objEd typ out handle)
  (setq out '())
  (if (dl:ename-p dictEname)
    (progn
      (setq handle (cdr (assoc 5 (entget dictEname))))
      (if (and handle (member handle dl:*dict-visited*))
        out
        (progn
          (if handle
            (setq dl:*dict-visited* (cons handle dl:*dict-visited*))
          )
          (setq rec (dictnext dictEname T))
          (while rec
            (setq key   (dl:trim (cdr (assoc 3 rec)))
                  objEn (or (cdr (assoc 360 rec))
                            (cdr (assoc 350 rec)))
                  objEd (if (dl:ename-p objEn) (entget objEn))
                  typ   (if objEd (cdr (assoc 0 objEd))))
            (cond
              ((= typ "DATALINK")
               (if (= key "")
                 (setq key (dl:get-datalink-name-from-object objEn))
               )
               (if (/= key "")
                 (setq out (dl:add-unique-ci key out))
               )
              )
              ((= typ "DICTIONARY")
               (foreach n (dl:get-datalink-names-from-dict-recursive objEn)
                 (setq out (dl:add-unique-ci n out))
               )
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

(defun dl:get-nod-dictionary-enames (/ nod rec en recEd out)
  (setq out '()
        nod (namedobjdict))
  (if (dl:ename-p nod)
    (progn
      (setq rec (dictnext nod T))
      (while rec
        (setq en (or (cdr (assoc 360 rec))
                     (cdr (assoc 350 rec))))
        (if (and (dl:ename-p en)
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

(defun dl:get-datalink-names-by-nod-scan (/ dictEnames dictEname n names)
  (setq names '()
        dictEnames (dl:get-nod-dictionary-enames))
  (setq dl:*dict-visited* nil)
  (foreach dictEname dictEnames
    (foreach n (dl:get-datalink-names-from-dict-recursive dictEname)
      (setq names (dl:add-unique-ci n names))
    )
  )
  names
)

(defun dl:get-datalink-names-by-object-scan (/ en ed nm out)
  (setq out '()
        en  (entnext))
  (while en
    (setq ed (entget en))
    (if (and ed (= (cdr (assoc 0 ed)) "DATALINK"))
      (progn
        (setq nm (dl:get-datalink-name-from-object en))
        (if (/= nm "")
          (setq out (dl:add-unique-ci nm out))
        )
      )
    )
    (setq en (entnext en))
  )
  out
)

(defun dl:get-cdate-part (dc start len)
  (atoi (substr dc start len))
)

(defun dl:set-dxf (ed code val / old)
  (if (setq old (assoc code ed))
    (subst (cons code val) old ed)
    (append ed (list (cons code val)))
  )
)

(defun dl:find-datalink-key-ci-in-dict (dictEname linkName / rec rawKey keyNorm target found)
  (if (not (dl:ename-p dictEname))
    nil
    (progn
      (setq target (strcase (dl:trim linkName))
            rec    (dictnext dictEname T)
            found  nil)
      (while (and rec (null found))
        (if (assoc 3 rec)
          (progn
            (setq rawKey  (cdr (assoc 3 rec))
                  keyNorm (strcase (dl:trim rawKey)))
            (if (= keyNorm target)
              (setq found rawKey)
            )
          )
        )
        (if (null found)
          (setq rec (dictnext dictEname))
        )
      )
      found
    )
  )
)

(defun dl:update-datalink-by-key (dictEname linkKey filePath sheetName / rec en ed conn detail dc newEd)
  (if (or (not (dl:ename-p dictEname))
          (= (dl:trim linkKey) ""))
    nil
    (progn
      (setq rec (dictsearch dictEname linkKey))
      (if (null rec)
        nil
        (progn
          (setq en (or (cdr (assoc 360 rec))
                       (cdr (assoc 350 rec))))
          (if (not (dl:ename-p en))
            nil
            (progn
              (setq ed (entget en))
              (if (or (null ed) (/= (cdr (assoc 0 ed)) "DATALINK"))
                nil
                (progn
                  (setq conn   (strcat filePath "!" sheetName)
                        detail (strcat "Data Link\n" linkKey "\n" filePath "\nLink details: Entire sheet: " sheetName)
                        dc     (rtos (getvar 'cdate) 2 20)
                        newEd  ed)

                  (setq newEd (dl:set-dxf newEd 301 detail))
                  (setq newEd (dl:set-dxf newEd 302 conn))
                  (setq newEd (dl:set-dxf newEd 91 dl:*safe-update-flag91*))
                  (setq newEd (dl:set-dxf newEd 92 dl:*safe-update-flag92*))
                  (setq newEd (dl:set-dxf newEd 170 (dl:get-cdate-part dc 1 4)))
                  (setq newEd (dl:set-dxf newEd 171 (dl:get-cdate-part dc 5 2)))
                  (setq newEd (dl:set-dxf newEd 172 (dl:get-cdate-part dc 7 2)))
                  (setq newEd (dl:set-dxf newEd 173 (dl:get-cdate-part dc 10 2)))
                  (setq newEd (dl:set-dxf newEd 174 (dl:get-cdate-part dc 12 2)))
                  (setq newEd (dl:set-dxf newEd 175 (dl:get-cdate-part dc 14 2)))
                  (setq newEd (dl:set-dxf newEd 176 (dl:get-cdate-part dc 16 2)))

                  (if (entmod newEd)
                    (progn
                      (entupd en)
                      en
                    )
                    nil
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

(defun dl:remove-datalink-by-key (dictEname linkKey / rec removed)
  (if (or (not (dl:ename-p dictEname))
          (= (dl:trim linkKey) ""))
    nil
    (progn
      (setq rec (dictsearch dictEname linkKey))
      (if rec
        (progn
          (setq removed (dictremove dictEname linkKey))
          (if removed T nil)
        )
        nil
      )
    )
  )
)

(defun dl:get-datalink-flags (dictEname / rec en ed flag91 flag92)
  (if (not (dl:ename-p dictEname))
    (list (dl:normalize-update-flag91 dl:*safe-update-flag91*)
          dl:*safe-update-flag92*)
    (progn
      (setq rec (dictnext dictEname T))
      (while rec
        (setq en (or (cdr (assoc 360 rec))
                     (cdr (assoc 350 rec))))
        (if (dl:ename-p en)
          (progn
            (setq ed (entget en))
            (if (and ed (= (cdr (assoc 0 ed)) "DATALINK"))
              (progn
                (if (assoc 91 ed)
                  (setq flag91 (cdr (assoc 91 ed)))
                )
                (if (assoc 92 ed)
                  (setq flag92 (cdr (assoc 92 ed)))
                )
                (setq rec nil)
              )
              (setq rec (dictnext dictEname))
            )
          )
          (setq rec (dictnext dictEname))
        )
      )
      (list (dl:normalize-update-flag91 (if flag91 flag91 dl:*safe-update-flag91*))
            (if flag92 flag92 dl:*safe-update-flag92*))
    )
  )
)

(defun dl:create-datalink (linkName filePath sheetName / dictObj dictEname dc dlem edl tempTC
                                                    dataLinkList connString detail flags flag91 flag92
                                                    existingKey overwriteOK okEntmod okDictadd updatedEn)
  (setq dl:*last-create-error* nil)
  (setq dictObj (dl:get-datalink-dict-object))

  (if (null dictObj)
    (progn
      (setq dl:*last-create-error* "no_datalink_dict")
      nil
    )
    (progn
      (setq dictEname (vlax-vla-object->ename dictObj)
            flags      (dl:get-datalink-flags dictEname)
            existingKey (dl:find-datalink-key-ci-in-dict dictEname linkName)
            overwriteOK T
            flag91      (car flags)
            flag92      (cadr flags))

      (if (not (dl:ename-p dictEname))
        (setq dictEname (dl:get-datalink-dict-ename))
      )
      (if (not (dl:ename-p dictEname))
        (progn
          (setq dl:*last-create-error* "invalid_datalink_dict")
          (dl:release dictObj)
          nil
        )
        (progn

          (setq flags      (dl:get-datalink-flags dictEname)
                existingKey (dl:find-datalink-key-ci-in-dict dictEname linkName)
                flag91      (car flags)
                flag92      (cadr flags))

          (if existingKey
            (progn
              ; 同名優先直接更新，避免 dictadd 在部分版本失敗
              (setq updatedEn (dl:update-datalink-by-key dictEname existingKey filePath sheetName))
              (if updatedEn
                (progn
                  (setq dl:*last-create-error* nil)
                  (dl:release dictObj)
                  updatedEn
                )
                (progn
                  (setq linkName existingKey
                        overwriteOK (dl:remove-datalink-by-key dictEname existingKey))
                )
              )
            )
          )

          (if (and (null updatedEn) (not overwriteOK))
            (progn
              (setq dl:*last-create-error* "overwrite_remove_failed")
              (dl:release dictObj)
              nil
            )
            (if updatedEn
              updatedEn
              (progn
              (setq dlem   (entmakex '((0 . "DATALINK") (100 . "AcDbDataLink"))))

              (setq connString (strcat filePath "!" sheetName)
                    detail     (strcat "Data Link\n" linkName "\n" filePath "\nLink details: Entire sheet: " sheetName))

              (if (null dlem)
                (progn
                  (setq dl:*last-create-error* "entmakex_datalink_failed")
                  (dl:release dictObj)
                  nil
                )
                (progn
                  (setq edl (entget dlem)
                        tempTC
                        (entmakex
                          (list
                            (cons 0 "TABLECONTENT")
                            (cons 100 "AcDbLinkedData")
                            (cons 100 "AcDbLinkedTableData")
                            (cons 92 0)
                            (cons 100 "AcDbFormattedTableData")
                            (cons 300 "TABLEFORMAT")
                            (cons 1 "TABLEFORMAT_BEGIN")
                            (cons 90 4)
                            (cons 170 0)
                            (cons 309 "TABLEFORMAT_END")
                            (cons 90 0)
                            (cons 100 "AcDbTableContent")
                          )
                        )
                  )

                  (if (null tempTC)
                    (progn
                      (setq dl:*last-create-error* "entmakex_tablecontent_failed")
                      (entdel dlem)
                      (dl:release dictObj)
                      nil
                    )
                    (progn
                      (entmod
                        (subst (cons 330 dictEname)
                               (assoc 330 (entget tempTC))
                               (entget tempTC)))

                      (setq dc (rtos (getvar 'cdate) 2 20)
                            dataLinkList
                            (list
                              (assoc -1 edl)
                              (cons 0 "DATALINK")
                              (cons 102 "{ACAD_REACTORS")
                              (cons 330 dictEname)
                              (cons 102 "}")
                              (cons 330 dictEname)
                              (cons 100 "AcDbDataLink")
                              (cons 1 "AcExcel")
                              (cons 300 "")
                              (cons 301 detail)
                              (cons 302 connString)
                              (cons 90 2)
                              (cons 91 flag91)
                              (cons 92 flag92)
                              (cons 170 (dl:get-cdate-part dc 1 4))
                              (cons 171 (dl:get-cdate-part dc 5 2))
                              (cons 172 (dl:get-cdate-part dc 7 2))
                              (cons 173 (dl:get-cdate-part dc 10 2))
                              (cons 174 (dl:get-cdate-part dc 12 2))
                              (cons 175 (dl:get-cdate-part dc 14 2))
                              (cons 176 (dl:get-cdate-part dc 16 2))
                              (cons 177 3)
                              (cons 93 0)
                              (cons 304 "")
                              (cons 94 0)
                              (cons 360 tempTC)
                              (cons 305 "CUSTOMDATA")
                              (cons 1 "DATAMAP_BEGIN")
                              (cons 90 3)
                              (cons 300 "ACEXCEL_UPDATEOPTIONS")
                              (cons 301 "DATAMAP_VALUE")
                              (cons 93 2)
                              (cons 90 1)
                              (cons 91 flag91)
                              (cons 94 0)
                              (cons 300 "")
                              (cons 302 "")
                              (cons 304 "ACVALUE_END")
                              (cons 300 "ACEXCEL_CONNECTION_STRING")
                              (cons 301 "DATAMAP_VALUE")
                              (cons 93 2)
                              (cons 90 4)
                              (cons 1 connString)
                              (cons 94 0)
                              (cons 300 "")
                              (cons 302 "")
                              (cons 304 "ACVALUE_END")
                              (cons 300 "ACEXCEL_SOURCEDATE")
                              (cons 301 "DATAMAP_VALUE")
                              (cons 93 2)
                              (cons 90 1)
                              (cons 92 16)
                              (cons 94 0)
                              (cons 300 "")
                              (cons 302 "")
                              (cons 304 "ACVALUE_END")
                              (cons 309 "DATAMAP_END")
                            )
                      )

                      (setq okEntmod (entmod dataLinkList))
                      (if okEntmod
                        (progn
                          (setq okDictadd (dictadd dictEname linkName dlem))
                          (if (null okDictadd)
                            (progn
                              (if (dl:remove-datalink-by-key dictEname linkName)
                                (setq okDictadd (dictadd dictEname linkName dlem))
                              )
                            )
                          )
                          (if okDictadd
                            (progn
                              (setq dl:*last-create-error* nil)
                              (dl:release dictObj)
                              dlem
                            )
                            (progn
                              (setq dl:*last-create-error* "dictadd_failed")
                              (entdel dlem)
                              (dl:release dictObj)
                              nil
                            )
                          )
                        )
                        (progn
                          (setq dl:*last-create-error* "entmod_failed")
                          (entdel dlem)
                          (dl:release dictObj)
                          nil
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
    )
  )
)

(defun dl:get-datalink-names (/ dictEname out)
  (setq out '())

  ; 1) 優先使用 ActiveX Dictionary 列舉（與資料連結管理員同源）
  (setq out (dl:get-datalink-names-vla))

  ; 2) 回退：讀標準 ACAD_DATALINK（ename/dictnext）
  (setq dictEname (dl:get-existing-datalink-dict-ename))
  (if (and (null out) (null dictEname))
    (setq dictEname (dl:get-datalink-dict-ename))
  )
  (if (and (null out) (not (dl:ename-p dictEname)))
    (setq dictEname nil)
  )
  (if (and (null out) dictEname)
    (progn
      (setq dl:*dict-visited* nil)
      (setq out (dl:get-datalink-names-from-dict-recursive dictEname))
    )
  )

  ; 3) 若沒抓到，掃描 NOD 底下所有字典，尋找實際 DATALINK 物件
  (if (null out)
    (setq out (dl:get-datalink-names-by-nod-scan))
  )

  ; 4) 最後回退：掃描整個 DB 的 DATALINK 物件
  (if (null out)
    (setq out (dl:get-datalink-names-by-object-scan))
  )

  out
)

(defun dl:get-datalink-names-dictonly (/ dictEname out)
  (setq out '())
  (setq dictEname (dl:get-existing-datalink-dict-ename))
  (if (and (null dictEname) (dl:ename-p (namedobjdict)))
    (setq dictEname (dl:get-datalink-dict-ename))
  )
  (if (dl:ename-p dictEname)
    (progn
      (setq dl:*dict-visited* nil)
      (setq out (dl:get-datalink-names-from-dict-recursive dictEname))
    )
  )
  out
)

(defun c:DLLINKLIST (/ names)
  (vl-load-com)
  (setq names (dl:get-datalink-names))
  (if names
    (progn
      (prompt "\n--- Data Link 清單 ---")
      (foreach n names
        (prompt (strcat "\n" n))
      )
    )
    (prompt "\n目前找不到任何 Data Link。")
  )
  (princ)
)

(defun c:DLLINKLISTDBG (/ namesVla namesDict namesNod)
  (vl-load-com)
  (setq namesVla  (dl:get-datalink-names-vla)
        namesDict (dl:get-datalink-names-dictonly)
        namesNod  (dl:get-datalink-names-by-nod-scan))
  (prompt "\n--- DLLINKLIST 偵錯 ---")
  (prompt (strcat "\nVLA(管理員來源) 數量: " (itoa (length namesVla))))
  (foreach n namesVla
    (prompt (strcat "\n  [VLA] " n))
  )
  (prompt (strcat "\nDict(ename) 數量: " (itoa (length namesDict))))
  (foreach n namesDict
    (prompt (strcat "\n  [DICT] " n))
  )
  (prompt (strcat "\nNOD掃描 數量: " (itoa (length namesNod))))
  (foreach n namesNod
    (prompt (strcat "\n  [NOD] " n))
  )
  (princ)
)

(defun c:AUTODATALINKLIST ()
  (c:DLLINKLIST)
)

(defun c:DATALINKLIST ()
  (c:DLLINKLIST)
)

(defun c:DATALINKLISTT ()
  (c:DLLINKLIST)
)

(defun c:DLAUTO (/ panelName filePath linkPath sheetNames sheetName manualSheet linkName)
  (vl-load-com)
  (setvar "cmdecho" 0)

  (setq panelName (dl:get-panel-name))
  (if (= panelName "")
    (prompt "\n未輸入盤名，已取消。")
    (progn
      (setq filePath (getfiled "選擇 Excel 來源檔案" (getvar "dwgprefix") "xlsx;xls;csv" 0))
      (if (null filePath)
        (prompt "\n未選取 Excel 檔案，已取消。")
        (progn
          (setq linkPath (dl:pick-link-path filePath))
          (setq sheetNames (dl:get-sheet-names filePath))
          (if (or (null sheetNames) (null (car sheetNames)))
            (prompt "\n無法讀取 Excel 分頁，請確認已安裝 Excel 且檔案可正常開啟。")
            (progn
              (setq sheetName (dl:find-sheet-by-panel panelName sheetNames))
              (if (null sheetName)
                (progn
                  (prompt "\n找不到與盤名相符的分頁，請手動指定。")
                  (dl:print-sheet-list sheetNames)
                  (setq manualSheet (getstring T (strcat "\n輸入分頁名稱 <" (car sheetNames) ">: ")))
                  (if (= (dl:trim manualSheet) "")
                    (setq sheetName (car sheetNames))
                    (setq sheetName (or (dl:sheet-exists-ci manualSheet sheetNames)
                                        manualSheet))
                  )
                )
              )

              (setq linkName (dl:sanitize-name panelName))

              (if (dl:create-datalink linkName linkPath sheetName)
                (progn
                  (prompt (strcat "\n已建立 Data Link: " linkName))
                  (prompt "\n若已有同名 Data Link，已自動覆蓋舊資料。")
                  (prompt (strcat "\n來源檔案(相對路徑): " linkPath))
                  (prompt (strcat "\n對應分頁: " sheetName))
                )
                (prompt (strcat "\nData Link 建立失敗。原因: " (if dl:*last-create-error* dl:*last-create-error* "unknown")))
              )
            )
          )
        )
      )
    )
  )
  (princ)
)

(defun c:AUTODATALINK ()
  (c:DLAUTO)
)

(defun c:DLAUTOSHEETS (/ filePath linkPath sheetNames sheetName linkName
                        okList failList)
  (vl-load-com)
  (setvar "cmdecho" 0)

  (setq filePath (getfiled "選擇 Excel 來源檔案" (getvar "dwgprefix") "xlsx;xls;csv" 0))
  (if (null filePath)
    (prompt "\n未選取 Excel 檔案，已取消。")
    (progn
      (setq linkPath (dl:pick-link-path filePath))
      (setq sheetNames (dl:get-sheet-names filePath))
      (if (or (null sheetNames) (null (car sheetNames)))
        (prompt "\n無法讀取 Excel 分頁，請確認已安裝 Excel 且檔案可正常開啟。")
        (progn
          (setq okList '()
                failList '())

          (foreach sheetName sheetNames
            (setq linkName sheetName)
            (if (= (dl:trim linkName) "")
              (setq failList (cons "(空白分頁名稱)" failList))
              (if (dl:create-datalink linkName linkPath sheetName)
                (setq okList (cons (strcat sheetName " -> " linkName) okList))
                (setq failList
                      (cons (strcat sheetName " (" (if dl:*last-create-error* dl:*last-create-error* "unknown") ")")
                            failList))
              )
            )
          )

          (prompt (strcat "\n全部分頁建立完成，總分頁數: " (itoa (length sheetNames))))
          (prompt (strcat "\n成功: " (itoa (length okList))
                          "，失敗: " (itoa (length failList))))
          (prompt (strcat "\n來源檔案(相對路徑): " linkPath))
          (prompt "\nData Link 名稱與 Excel 分頁名稱一致；同名 Data Link 會自動覆蓋舊資料。")

          (if okList
            (progn
              (prompt "\n--- 成功清單 ---")
              (foreach x (reverse okList)
                (prompt (strcat "\n" x))
              )
            )
          )
          (if failList
            (progn
              (prompt "\n--- 建立失敗 ---")
              (foreach x (reverse failList)
                (prompt (strcat "\n" x))
              )
            )
          )
        )
      )
    )
  )
  (princ)
)

(defun c:AUTODATALINKSHEETS ()
  (c:DLAUTOSHEETS)
)

(defun c:DLAUTOALLSHEETS ()
  (c:DLAUTOSHEETS)
)

(defun c:DLAUTOBATCH (/ filePath linkPath sheetNames ss panelNames
                        panel sheetName linkName
                        okList failCreate noSheetList)
  (vl-load-com)
  (setvar "cmdecho" 0)

  (setq filePath (getfiled "選擇 Excel 來源檔案" (getvar "dwgprefix") "xlsx;xls;csv" 0))
  (if (null filePath)
    (prompt "\n未選取 Excel 檔案，已取消。")
    (progn
      (setq linkPath (dl:pick-link-path filePath))
      (setq sheetNames (dl:get-sheet-names filePath))
      (if (or (null sheetNames) (null (car sheetNames)))
        (prompt "\n無法讀取 Excel 分頁，請確認已安裝 Excel 且檔案可正常開啟。")
        (progn
          (prompt "\n請選取要批量建立的盤名文字物件...")
          (setq ss (ssget '((0 . "TEXT,MTEXT,ATTRIB,ATTDEF"))))
          (if (null ss)
            (prompt "\n未選取任何文字物件，已取消。")
            (progn
              (setq panelNames (dl:get-panel-names-from-ss ss))
              (if (null panelNames)
                (prompt "\n選取物件中沒有可用盤名，已取消。")
                (progn
                  (setq okList '()
                        failCreate '()
                        noSheetList '())

                  (foreach panel panelNames
                    (setq sheetName (dl:find-sheet-by-panel panel sheetNames))
                    (if sheetName
                      (progn
                        (setq linkName (dl:sanitize-name panel))
                        (if (dl:create-datalink linkName linkPath sheetName)
                          (setq okList (cons (strcat panel " -> " sheetName " -> " linkName) okList))
                          (setq failCreate
                                (cons (strcat panel " (" (if dl:*last-create-error* dl:*last-create-error* "unknown") ")")
                                      failCreate))
                        )
                      )
                      (setq noSheetList (cons panel noSheetList))
                    )
                  )

                  (prompt (strcat "\n批量建立完成，共處理盤名: " (itoa (length panelNames))))
                  (prompt (strcat "\n成功: " (itoa (length okList))
                                  "，找不到分頁: " (itoa (length noSheetList))
                                  "，建立失敗: " (itoa (length failCreate))))
                  (prompt (strcat "\n來源檔案(相對路徑): " linkPath))
                  (prompt "\n同名 Data Link 會自動覆蓋舊資料。")

                  (if okList
                    (progn
                      (prompt "\n--- 成功清單 ---")
                      (foreach x (reverse okList)
                        (prompt (strcat "\n" x))
                      )
                    )
                  )
                  (if noSheetList
                    (progn
                      (prompt "\n--- 找不到對應分頁 ---")
                      (foreach x (reverse noSheetList)
                        (prompt (strcat "\n" x))
                      )
                    )
                  )
                  (if failCreate
                    (progn
                      (prompt "\n--- 建立失敗 ---")
                      (foreach x (reverse failCreate)
                        (prompt (strcat "\n" x))
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
  (princ)
)

(defun c:AUTODATALINKBATCH ()
  (c:DLAUTOBATCH)
)

(princ "\nDLAUTO / AUTODATALINK / DLAUTOSHEETS / AUTODATALINKSHEETS / DLAUTOALLSHEETS / DLAUTOBATCH / AUTODATALINKBATCH / DLLINKLIST / AUTODATALINKLIST / DATALINKLIST / DATALINKLISTT / DLLINKLISTDBG 載入完成。")
(princ)
