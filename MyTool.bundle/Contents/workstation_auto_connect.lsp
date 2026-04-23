(vl-load-com)

(setq wsc:*eps* 1e-9)
(setq wsc:*sort-axis* '(1.0 0.0 0.0))
(setq wsc:*sample-pline-layer* "電氣管線")
(setq wsc:*axis-align-tol* 1e-6)

(defun wsc:trim (s)
  (if s
    (vl-string-trim " \t\r\n" s)
    ""
  )
)

(defun wsc:str-eq-ci (a b)
  (= (strcase (wsc:trim a)) (strcase (wsc:trim b)))
)

(defun wsc:pt3 (p / z)
  (setq z (if (> (length p) 2) (nth 2 p) 0.0))
  (list (float (car p)) (float (cadr p)) (float z))
)

(defun wsc:dist2 (a b / dx dy dz)
  (setq dx (- (car a) (car b))
        dy (- (cadr a) (cadr b))
        dz (- (caddr a) (caddr b)))
  (+ (* dx dx) (* dy dy) (* dz dz))
)

(defun wsc:axis-aligned-kind (p1 p2 / dx dy)
  (setq dx (abs (- (car p1) (car p2)))
        dy (abs (- (cadr p1) (cadr p2))))
  (cond
    ((<= dx wsc:*axis-align-tol*) 'vertical)
    ((<= dy wsc:*axis-align-tol*) 'horizontal)
    (T nil)
  )
)

(defun wsc:normalize-axis-pair (p1 p2 kind / x y)
  (cond
    ((= kind 'vertical)
      (setq x (car p1))
      (list
        (list x (cadr p1) (caddr p1))
        (list x (cadr p2) (caddr p2))
      )
    )
    ((= kind 'horizontal)
      (setq y (cadr p1))
      (list
        (list (car p1) y (caddr p1))
        (list (car p2) y (caddr p2))
      )
    )
    (T nil)
  )
)

(defun wsc:dot2d (pt axis)
  (+ (* (car pt) (car axis))
     (* (cadr pt) (cadr axis)))
)

(defun wsc:ensure-layer (name / lay)
  (setq lay (tblsearch "LAYER" name))
  (if lay
    T
    (progn
      (entmake
        (list
          '(0 . "LAYER")
          '(100 . "AcDbSymbolTableRecord")
          '(100 . "AcDbLayerTableRecord")
          (cons 2 name)
          '(70 . 0)
          '(62 . 7)
          '(6 . "Continuous")
        )
      )
      (if (tblsearch "LAYER" name)
        T
        nil
      )
    )
  )
)

(defun wsc:normalize2d (vec / len)
  (setq len (sqrt (+ (* (car vec) (car vec))
                     (* (cadr vec) (cadr vec)))))
  (if (> len wsc:*eps*)
    (list (/ (car vec) len) (/ (cadr vec) len) 0.0)
    nil
  )
)

(defun wsc:get-effective-block-name (ename / obj val ed)
  (setq obj (vlax-ename->vla-object ename))
  (cond
    ((and obj (vlax-property-available-p obj 'EffectiveName))
      (setq val (vl-catch-all-apply 'vla-get-EffectiveName (list obj)))
      (if (vl-catch-all-error-p val)
        ""
        (wsc:trim val)
      )
    )
    ((setq ed (entget ename))
      (wsc:trim (cdr (assoc 2 ed)))
    )
    (T "")
  )
)

(defun wsc:text-like-type-p (typ)
  (if typ
    (wcmatch (strcase (wsc:trim typ)) "TEXT,MTEXT,ATTRIB,ATTDEF")
    nil
  )
)

(defun wsc:get-text-from-entity (ename / ed typ obj txt)
  (setq txt "")
  (if ename
    (progn
      (setq ed  (entget ename)
            typ (cdr (assoc 0 ed)))
      (if (wsc:text-like-type-p typ)
        (progn
          (setq obj (vlax-ename->vla-object ename))
          (if obj
            (setq txt (vla-get-TextString obj))
          )
        )
      )
    )
  )
  (wsc:trim txt)
)

(defun wsc:get-first-nonempty-text-in-ss (ss / i total en txt out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out "")
  (while (and (< i total) (= out ""))
    (setq en  (ssname ss i)
          txt (wsc:get-text-from-entity en))
    (if (/= txt "")
      (setq out txt)
    )
    (setq i (1+ i))
  )
  out
)

(defun wsc:spec-text-token (spec)
  (cond
    ((= (car spec) 'insert)
      (if (> (length spec) 2) (nth 2 spec) "")
    )
    ((= (car spec) 'entity)
      (if (> (length spec) 3) (nth 3 spec) "")
    )
    ((= (car spec) 'text)
      (if (> (length spec) 3) (nth 3 spec) "")
    )
    (T "")
  )
)

(defun wsc:entity-has-nearby-text-p (ename token / bbox mn mx size margin p1 p2 ss i total en txt found)
  (if (= (wsc:trim token) "")
    T
    (progn
      (setq bbox (wsc:get-bbox ename))
      (if bbox
        (progn
          (setq mn (car bbox)
                mx (cadr bbox)
                size (max (abs (- (car mx) (car mn)))
                          (abs (- (cadr mx) (cadr mn))))
                margin (max 50.0 (* size 0.8))
                p1 (list (- (car mn) margin) (- (cadr mn) margin) 0.0)
                p2 (list (+ (car mx) margin) (+ (cadr mx) margin) 0.0)
                ss (ssget "_C" p1 p2 '((0 . "TEXT,MTEXT,ATTRIB,ATTDEF"))))

          (setq i 0
                total (if ss (sslength ss) 0)
                found nil)
          (while (and (< i total) (not found))
            (setq en  (ssname ss i)
                  txt (wsc:get-text-from-entity en))
            (if (wsc:str-eq-ci token txt)
              (setq found T)
            )
            (setq i (1+ i))
          )
          found
        )
        nil
      )
    )
  )
)

(defun wsc:make-station-spec (ename legendText / ed typ lyr txt token)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed))
        lyr (cdr (assoc 8 ed))
        txt (if (wsc:text-like-type-p typ) (wsc:get-text-from-entity ename) "")
        token (strcase (wsc:trim legendText)))
  (cond
    ((= typ "INSERT")
      (if (= token "")
        (list 'insert (strcase (wsc:get-effective-block-name ename)))
        (list 'insert (strcase (wsc:get-effective-block-name ename)) token)
      )
    )
    ((and (wsc:text-like-type-p typ) (/= txt ""))
      ; 文字圖例：用 類型+圖層+文字內容 比對（例如 E3）
      (list 'text
            (strcase (wsc:trim typ))
            (strcase (wsc:trim lyr))
            (strcase txt))
    )
    (T
      ; 非 block 時，用「圖元類型 + 圖層」比對；若圖例框選有文字則一併比對文字
      (if (= token "")
        (list 'entity (strcase (wsc:trim typ)) (strcase (wsc:trim lyr)))
        (list 'entity (strcase (wsc:trim typ)) (strcase (wsc:trim lyr)) token)
      )
    )
  )
)

(defun wsc:station-match-p (ename spec / ed typ lyr txt token)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed))
        lyr (cdr (assoc 8 ed))
        txt (wsc:get-text-from-entity ename)
        token (wsc:spec-text-token spec))
  (cond
    ((= (car spec) 'insert)
      (and (= typ "INSERT")
           (wsc:str-eq-ci (cadr spec) (wsc:get-effective-block-name ename))
           (wsc:entity-has-nearby-text-p ename token))
    )
    ((= (car spec) 'text)
      (and (wsc:str-eq-ci (cadr spec) typ)
           (wsc:str-eq-ci (caddr spec) lyr)
           (wsc:str-eq-ci (cadddr spec) txt))
    )
    ((= (car spec) 'entity)
      (and (wsc:str-eq-ci (cadr spec) typ)
           (wsc:str-eq-ci (caddr spec) lyr)
           (wsc:entity-has-nearby-text-p ename token))
    )
    (T nil)
  )
)

(defun wsc:station-match-relaxed-p (ename spec / ed typ txt token)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed))
        txt (wsc:get-text-from-entity ename)
        token (wsc:spec-text-token spec))
  (cond
    ((= (car spec) 'text)
      ; 寬鬆比對：文字型只看 類型+文字內容（忽略圖層）
      (and (wsc:str-eq-ci (cadr spec) typ)
           (wsc:str-eq-ci (cadddr spec) txt))
    )
    ((= (car spec) 'entity)
      ; 寬鬆比對：一般圖元只看類型（忽略圖層）
      (and (wsc:str-eq-ci (cadr spec) typ)
           (wsc:entity-has-nearby-text-p ename token))
    )
    (T nil)
  )
)

(defun wsc:get-selection-like-pl90f (msg / ss)
  (prompt msg)
  (setq ss (ssget))
  ss
)

(defun wsc:legend-type-score (typ / t2)
  (setq t2 (strcase (wsc:trim typ)))
  (cond
    ((= t2 "INSERT") 0)
    ((= t2 "CIRCLE") 1)
    ((= t2 "ELLIPSE") 2)
    ((wsc:text-like-type-p t2) 3)
    ((= t2 "ARC") 4)
    ((wcmatch t2 "LWPOLYLINE,POLYLINE") 5)
    ((= t2 "LINE") 6)
    (T 9)
  )
)

(defun wsc:legend-entity-score (ename / typ txt score)
  (setq typ   (cdr (assoc 0 (entget ename)))
        score (wsc:legend-type-score typ))
  (if (and (wsc:text-like-type-p typ)
           (= (wsc:get-text-from-entity ename) ""))
    (+ score 2)
    score
  )
)

(defun wsc:pick-legend-from-ss (ss / i total en bestEnt bestScore score)
  (setq i 0
        total (if ss (sslength ss) 0)
        bestEnt nil
        bestScore nil)
  (while (< i total)
    (setq en    (ssname ss i)
          score (wsc:legend-entity-score en))
    (if (or (null bestScore) (< score bestScore))
      (progn
        (setq bestEnt en)
        (setq bestScore score)
      )
    )
    (setq i (1+ i))
  )
  bestEnt
)

(defun wsc:get-legend-selection (/ ss en txt)
  (setq ss
        (wsc:get-selection-like-pl90f
          "\n1/5 框選/多選工作站圖例: "))
  (if ss
    (progn
      (setq en (wsc:pick-legend-from-ss ss))
      (setq txt (wsc:get-first-nonempty-text-in-ss ss))
      (if en
        (progn
          (if (> (sslength ss) 1)
            (prompt "\n已從框選結果自動判定圖例（優先 Block/圓/文字）。")
          )
          (if (/= (wsc:trim txt) "")
            (prompt (strcat "\n圖例文字條件: " txt))
          )
          (list en txt)
        )
        nil
      )
    )
    nil
  )
)

(defun wsc:get-area-selection ()
  (wsc:get-selection-like-pl90f
    "\n2/5 框選/多選要連線區域: ")
)

(defun wsc:collect-stations-by-rule (ss spec rule / i total en out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out '())
  (while (< i total)
    (setq en (ssname ss i))
    (if (and en
             (or (and (= rule 'exact) (wsc:station-match-p en spec))
                 (and (= rule 'relaxed) (wsc:station-match-relaxed-p en spec))))
      (setq out (cons en out))
    )
    (setq i (1+ i))
  )
  (reverse out)
)

(defun wsc:collect-stations (ss spec / out token)
  (setq token (wsc:spec-text-token spec))
  (setq out (wsc:collect-stations-by-rule ss spec 'exact))
  (if (and (null out) (not (= (car spec) 'insert)))
    (progn
      (setq out (wsc:collect-stations-by-rule ss spec 'relaxed))
      (if out
        (prompt "\n3/5 偵測到精準比對為 0，已改用寬鬆比對（忽略圖層）重試。")
      )
    )
  )
  (if (and (null out) (/= (wsc:trim token) ""))
    (prompt (strcat "\n3/5 目前已啟用圖例文字條件: " token "，請確認區域內物件附近有相同文字。"))
  )
  out
)

(defun wsc:get-base-point (ename / ed p)
  (setq ed (entget ename)
        p  (cdr (assoc 10 ed)))
  (if p
    (wsc:pt3 p)
    nil
  )
)

(defun wsc:get-bbox (ename / obj ret minPt maxPt)
  (setq obj (vlax-ename->vla-object ename))
  (if obj
    (progn
      (setq ret (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'minPt 'maxPt)))
      (if (vl-catch-all-error-p ret)
        nil
        (list
          (wsc:pt3 (vlax-safearray->list minPt))
          (wsc:pt3 (vlax-safearray->list maxPt))
        )
      )
    )
    nil
  )
)

(defun wsc:center-from-bbox (bbox / mn mx)
  (setq mn (car bbox)
        mx (cadr bbox))
  (list
    (/ (+ (car mn) (car mx)) 2.0)
    (/ (+ (cadr mn) (cadr mx)) 2.0)
    (/ (+ (caddr mn) (caddr mx)) 2.0)
  )
)

(defun wsc:size-from-bbox (bbox / mn mx w h)
  (setq mn (car bbox)
        mx (cadr bbox)
        w  (abs (- (car mx) (car mn)))
        h  (abs (- (cadr mx) (cadr mn))))
  (max w h)
)

(defun wsc:quarters-from-bbox (bbox / mn mx cx cy cz)
  (setq mn (car bbox)
        mx (cadr bbox)
        cx (/ (+ (car mn) (car mx)) 2.0)
        cy (/ (+ (cadr mn) (cadr mx)) 2.0)
        cz (/ (+ (caddr mn) (caddr mx)) 2.0))
  ; 四分點：上/右/下/左中點
  (list
    (list cx (cadr mx) cz)
    (list (car mx) cy cz)
    (list cx (cadr mn) cz)
    (list (car mn) cy cz)
  )
)

(defun wsc:node-from-entity (ename / bbox c q p s)
  (setq bbox (wsc:get-bbox ename))
  (if bbox
    (progn
      (setq c (wsc:center-from-bbox bbox)
            q (wsc:quarters-from-bbox bbox)
            s (wsc:size-from-bbox bbox))
      (list ename c q s)
    )
    (progn
      (setq p (wsc:get-base-point ename))
      (if p
        (list ename p (list p p p p) 0.0)
        nil
      )
    )
  )
)

(defun wsc:build-nodes (entities / out node)
  (setq out '())
  (foreach en entities
    (setq node (wsc:node-from-entity en))
    (if node
      (setq out (cons node out))
    )
  )
  (reverse out)
)

(defun wsc:avg-node-size (nodes / sum cnt sz)
  (setq sum 0.0
        cnt 0)
  (foreach n nodes
    (setq sz (nth 3 n))
    (if (and sz (> sz wsc:*eps*))
      (progn
        (setq sum (+ sum sz))
        (setq cnt (1+ cnt))
      )
    )
  )
  (if (> cnt 0)
    (/ sum cnt)
    0.0
  )
)

(defun wsc:entity-in-list-p (ename lst)
  (if (vl-some '(lambda (x) (eq x ename)) lst)
    T
    nil
  )
)

(defun wsc:find-node-by-ename (ename nodes)
  (vl-some
    '(lambda (n)
       (if (eq (car n) ename)
         n
       )
     )
    nodes
  )
)

(defun wsc:pick-first-station-in-ss (ss stations / i total en first cnt)
  (setq i 0
        total (if ss (sslength ss) 0)
        first nil
        cnt 0)
  (while (< i total)
    (setq en (ssname ss i))
    (if (wsc:entity-in-list-p en stations)
      (progn
        (setq cnt (1+ cnt))
        (if (null first)
          (setq first en)
        )
      )
    )
    (setq i (1+ i))
  )
  (list first cnt)
)

(defun wsc:pick-station-from-list (msg stations / ss hit en cnt)
  (setq en nil)
  (while (null en)
    (prompt msg)
    (setq ss (ssget))
    (cond
      ((null ss)
        (setq en 'wsc_cancel)
      )
      (T
        (setq hit (wsc:pick-first-station-in-ss ss stations)
              en  (car hit)
              cnt (cadr hit))
        (if (null en)
          (prompt "\n框選結果不含可用工作站，請重選。")
          (if (> cnt 1)
            (prompt "\n已框到多個工作站，將使用第一個符合者。")
          )
        )
      )
    )
  )
  (if (eq en 'wsc_cancel)
    nil
    en
  )
)

(defun wsc:extract-style-from-entity (ename / ed out)
  (setq ed (entget ename)
        out '())
  (foreach d ed
    (if (member (car d) '(8 6 48 62 370 420 430 440))
      (setq out (append out (list d)))
    )
  )
  out
)

(defun wsc:default-style ()
  (list
    (cons 8 (getvar "CLAYER"))
    (cons 6 "ByLayer")
    (cons 62 256)
    (cons 370 -1)
  )
)

(defun wsc:find-new-polyline (before / e typ out)
  (setq e (if before (entnext before) (entnext))
        out nil)
  (while e
    (setq typ (cdr (assoc 0 (entget e))))
    (if (wcmatch typ "LWPOLYLINE,POLYLINE")
      (setq out e)
    )
    (setq e (entnext e))
  )
  out
)

(defun wsc:get-connection-style (/ oldLayer before drawRes pl style)
  (setq oldLayer (getvar "CLAYER")
        style nil)
  (if (not (wsc:ensure-layer wsc:*sample-pline-layer*))
    (progn
      (prompt (strcat "\n無法建立圖層「" wsc:*sample-pline-layer* "」，已取消。"))
      nil
    )
    (progn
      (setvar "CLAYER" wsc:*sample-pline-layer*)
      (prompt (strcat "\n4/5 請在圖層「" wsc:*sample-pline-layer* "」使用 PLINE 畫範例線段，完成請 Enter，取消請 ESC。"))
      (setq before (entlast))
      (setq drawRes
            (vl-catch-all-apply
              '(lambda ()
                 (command "_.PLINE")
                 (while (/= 0 (logand 1 (getvar "CMDACTIVE")))
                   (command pause)
                 )
               )
            )
      )
      (setvar "CLAYER" oldLayer)
      (if (vl-catch-all-error-p drawRes)
        (progn
          (prompt "\n未完成範例 PLINE，已取消。")
          nil
        )
        (progn
          (setq pl (wsc:find-new-polyline before))
          (if pl
            (progn
              (setq style (wsc:extract-style-from-entity pl))
              (if (null style)
                (setq style (wsc:default-style))
              )
              style
            )
            (progn
              (prompt "\n未偵測到新建的 PLINE，請重試。")
              nil
            )
          )
        )
      )
    )
  )
)

(defun wsc:get-sample-config (stations nodes / s1 s2 n1 n2 dir perp size1 size2 base tol style)
  (prompt "\n4/5 設定連線範例：先框選 2 個工作站決定方向，再用 PLINE 畫範例線。")
  (setq s1 (wsc:pick-station-from-list "\n4/5 框選/多選範例起點工作站: " stations))
  (if s1
    (progn
      (setq s2 (wsc:pick-station-from-list "\n4/5 框選/多選範例終點工作站: " stations))
      (if (or (null s2) (eq s1 s2))
        (progn
          (if (eq s1 s2)
            (prompt "\n範例起訖不可為同一個工作站。")
          )
          nil
        )
        (progn
          (setq n1 (wsc:find-node-by-ename s1 nodes)
                n2 (wsc:find-node-by-ename s2 nodes))
          (if (or (null n1) (null n2))
            (progn
              (prompt "\n找不到範例工作站幾何資訊。")
              nil
            )
            (progn
              (setq dir
                    (wsc:normalize2d
                      (list
                        (- (car (cadr n2)) (car (cadr n1)))
                        (- (cadr (cadr n2)) (cadr (cadr n1)))
                        0.0
                      )
                    )
              )
              (if (null dir)
                (progn
                  (prompt "\n範例方向過短，請重試。")
                  nil
                )
                (progn
                  (setq perp  (list (- (cadr dir)) (car dir) 0.0)
                        size1 (nth 3 n1)
                        size2 (nth 3 n2)
                        base  (/ (+ size1 size2) 2.0))
                  (if (<= base wsc:*eps*)
                    (setq base (wsc:avg-node-size nodes))
                  )
                  (if (<= base wsc:*eps*)
                    (setq base 100.0)
                  )
                  (setq tol (* base 0.75))
                  (prompt (strcat "\n分列/分欄容差: " (rtos tol 2 2)))
                  (setq style (wsc:get-connection-style))
                  (if style
                    (list dir perp tol style)
                    nil
                  )
                )
              )
            )
          )
        )
      )
    )
    nil
  )
)

(defun wsc:node-axis-key (node axis)
  (wsc:dot2d (cadr node) axis)
)

(defun wsc:node-less-by-current-axis (a b)
  (< (wsc:node-axis-key a wsc:*sort-axis*)
     (wsc:node-axis-key b wsc:*sort-axis*))
)

(defun wsc:sort-nodes-by-axis (nodes axis)
  (setq wsc:*sort-axis* axis)
  (vl-sort nodes 'wsc:node-less-by-current-axis)
)

(defun wsc:group-by-perp-axis (nodes perpAxis tol / sorted groups current rowKey rowCount n key)
  (setq sorted   (wsc:sort-nodes-by-axis nodes perpAxis)
        groups   '()
        current  '()
        rowKey   nil
        rowCount 0)

  (foreach n sorted
    (setq key (wsc:node-axis-key n perpAxis))
    (if (or (null rowKey) (<= (abs (- key rowKey)) tol))
      (progn
        (setq current (append current (list n)))
        (if (null rowKey)
          (progn
            (setq rowKey key)
            (setq rowCount 1)
          )
          (progn
            (setq rowCount (1+ rowCount))
            (setq rowKey (/ (+ (* rowKey (1- rowCount)) key) rowCount))
          )
        )
      )
      (progn
        (setq groups (append groups (list current)))
        (setq current  (list n)
              rowKey   key
              rowCount 1)
      )
    )
  )

  (if current
    (setq groups (append groups (list current)))
  )
  groups
)

(defun wsc:best-quarter-pair (nodeA nodeB / qa qb pa pb best bestD d kind norm)
  (setq qa    (caddr nodeA)
        qb    (caddr nodeB)
        best  nil
        bestD nil)
  (foreach pa qa
    (foreach pb qb
      (setq kind (wsc:axis-aligned-kind pa pb))
      (if kind
        (progn
          (setq norm (wsc:normalize-axis-pair pa pb kind)
                d    (wsc:dist2 (car norm) (cadr norm)))
          (if (or (null bestD) (< d bestD))
            (setq bestD d
                  best  norm)
          )
        )
      )
    )
  )
  best
)

(defun wsc:make-line-with-style (p1 p2 style / data)
  (setq data
        (append
          (list
            '(0 . "LINE")
            (cons 10 p1)
            (cons 11 p2)
          )
          style))
  (entmakex data)
)

(defun wsc:connect-order (order style / i total nodeA nodeB pair created skipped)
  (setq i 0
        total (length order)
        created 0
        skipped 0)
  (while (< i (1- total))
    (setq nodeA (nth i order)
          nodeB (nth (1+ i) order)
          pair  (wsc:best-quarter-pair nodeA nodeB))
    (if (and pair (wsc:make-line-with-style (car pair) (cadr pair) style))
      (setq created (1+ created))
      (setq skipped (1+ skipped))
    )
    (setq i (1+ i))
  )
  (list created skipped)
)

(defun wsc:connect-groups-by-direction (groups dirAxis style / created skipped row sorted rowResult)
  (setq created 0
        skipped 0)
  (foreach row groups
    (setq sorted (wsc:sort-nodes-by-axis row dirAxis))
    (if (> (length sorted) 1)
      (progn
        (setq rowResult (wsc:connect-order sorted style))
        (setq created (+ created (car rowResult))
              skipped (+ skipped (cadr rowResult)))
      )
    )
  )
  (list created skipped)
)

(defun c:WSAUTOCONNECT (/ oldErr oldEcho legendInfo legend legendText spec areaSS stations nodes cfg dir perp tol style groups connectResult created skipped)
  (vl-load-com)
  (setq oldErr  *error*
        oldEcho (getvar "CMDECHO"))

  (defun *error* (msg)
    (setvar "CMDECHO" oldEcho)
    (if oldErr
      (setq *error* oldErr)
      (setq *error* nil)
    )
    (cond
      ((or (= msg "Function cancelled")
           (= msg "quit / exit abort"))
        (prompt "\nWSAUTOCONNECT 已取消。")
      )
      ((and msg (/= msg ""))
        (prompt (strcat "\n錯誤: " msg))
      )
    )
    (princ)
  )

  (setvar "CMDECHO" 0)
  (prompt "\n=== 工作站自動連線 ===")

  (setq legendInfo (wsc:get-legend-selection))
  (if (null legendInfo)
    (prompt "\n未選取工作站圖例，已取消。")
    (progn
      (setq legend     (car legendInfo)
            legendText (cadr legendInfo)
            spec       (wsc:make-station-spec legend legendText))

      (setq areaSS (wsc:get-area-selection))
      (if (null areaSS)
        (prompt "\n未選取連線區域，已取消。")
        (progn
          (setq stations (wsc:collect-stations areaSS spec))
          (prompt (strcat "\n3/5 已過濾工作站數量: " (itoa (length stations))))

          (if (< (length stations) 2)
            (prompt "\n區域內可連線工作站不足 2 個，已取消。")
            (progn
              (setq nodes (wsc:build-nodes stations))
              (if (< (length nodes) 2)
                (prompt "\n可用工作站節點不足 2 個，已取消。")
                (progn
                  (setq cfg (wsc:get-sample-config stations nodes))
                  (if (null cfg)
                    (prompt "\n未完成連線範例設定，已取消。")
                    (progn
                      (setq dir   (nth 0 cfg)
                            perp  (nth 1 cfg)
                            tol   (nth 2 cfg)
                            style (nth 3 cfg)
                            groups (wsc:group-by-perp-axis nodes perp tol))

                      (prompt "\n5/5 開始自動連線（僅水平/垂直，全部由四分點起訖）...")
                      (setq connectResult (wsc:connect-groups-by-direction groups dir style)
                            created       (car connectResult)
                            skipped       (cadr connectResult))
                      (prompt
                        (strcat
                          "\n完成連線，共建立 "
                          (itoa created)
                          " 條線段；略過斜線 "
                          (itoa skipped)
                          " 條；分組數量: "
                          (itoa (length groups))
                          "。"
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

  (setvar "CMDECHO" oldEcho)
  (if oldErr
    (setq *error* oldErr)
    (setq *error* nil)
  )
  (princ)
)

(defun c:AUTOWSCONNECT ()
  (c:WSAUTOCONNECT)
)

(defun c:WORKSTATIONAUTOCONNECT ()
  (c:WSAUTOCONNECT)
)

(prompt "\nWSAUTOCONNECT / AUTOWSCONNECT / WORKSTATIONAUTOCONNECT 載入完成。")
(princ)
